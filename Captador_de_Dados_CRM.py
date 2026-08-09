# ==============================================================
# INTERFACE GRÁFICA MODERNA - CAPTADOR DE DADOS CRM
# ==============================================================
# Requer: pip install customtkinter selenium webdriver-manager
#         pandas openpyxl pillow
# ==============================================================

import os
import io
import re
import sys
import json
import time
import glob
import shutil
import queue
import threading
import subprocess
from datetime import datetime

import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys


URL_CRM = (
    "https://kroton.crm2.dynamics.com/main.aspx?"
    "appid=bd6da805-c090-ec11-b400-0022483841d9"
    "&pagetype=custom"
    "&name=kcs_buscarinscricaov2_2fc14"
)

NOME_PRODUTO = "Captador de Dados CRM"
NOME_ARQUIVO_MANUAL = "Manual_do_Usuario_Captador_de_Dados_CRM.pdf"


def caminho_recurso(nome_arquivo):
    """
    Resolve o caminho de um arquivo empacotado junto do .exe (ícone, etc).
    Funciona rodando o .py direto, dentro do .exe gerado pelo PyInstaller
    (que descompacta os recursos numa pasta temporária em sys._MEIPASS) e
    também não quebra se rodado num notebook/REPL, onde "__file__" não
    existe.
    """
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            base = os.getcwd()

    return os.path.join(base, nome_arquivo)


def _pasta_utilizavel(pasta):
    """
    Tenta criar a pasta e confirma que dá pra escrever de verdade nela
    (criar a pasta pode "funcionar" sem erro mesmo quando o caminho está
    quebrado — por exemplo pastas "Documentos" redirecionadas pelo OneDrive
    que apontam pra um link morto). Só volta True se conseguir criar e
    gravar um arquivo de teste dentro dela.
    """
    try:
        os.makedirs(pasta, exist_ok=True)
        arquivo_teste = os.path.join(pasta, ".teste_escrita.tmp")
        with open(arquivo_teste, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(arquivo_teste)
        return True
    except Exception:
        return False


def pasta_base_dados():
    """
    Pasta GERAL onde toda a automação guarda seus dados (Entrada, Saida,
    Logs, ChromeProfile). Por padrão fica em "Documentos/Captador de Dados
    CRM" do usuário — mas em algumas máquinas essa pasta está redirecionada
    (OneDrive) para um caminho quebrado, o que faz até o os.makedirs
    falhar com WinError 2. Por isso tentamos várias opções, em ordem, e
    usamos a primeira que realmente aceitar leitura/escrita.
    """
    home = os.path.expanduser("~")

    candidatos = [
        os.path.join(home, "Documents", NOME_PRODUTO),
        os.path.join(home, "Documentos", NOME_PRODUTO),  # Windows em PT-BR
        os.path.join(home, NOME_PRODUTO),                 # direto na pasta do usuário
        os.path.join(caminho_recurso("."), NOME_PRODUTO),  # ao lado do .py/.exe
    ]

    for candidato in candidatos:
        if _pasta_utilizavel(candidato):
            return candidato

    # Último recurso: pasta temporária do sistema (sempre gravável).
    import tempfile
    candidato = os.path.join(tempfile.gettempdir(), NOME_PRODUTO)
    os.makedirs(candidato, exist_ok=True)
    return candidato


PASTA_DADOS = pasta_base_dados()

PASTA_ENTRADA = os.path.join(PASTA_DADOS, "Entrada")
PASTA_SAIDA = os.path.join(PASTA_DADOS, "Saida")
PASTA_LOGS = os.path.join(PASTA_DADOS, "Logs")
PASTA_SCREENSHOTS = os.path.join(PASTA_LOGS, "Screenshots")
PASTA_PERFIS = os.path.join(PASTA_DADOS, "ChromeProfile")

ARQUIVO_CPFS_PADRAO = os.path.join(PASTA_ENTRADA, "cpfs.xlsx")
NOME_ARQUIVO_PARCIAL = "Resultado_parcial.xlsx"
NOME_ARQUIVO_BACKUP_CSV = "Resultado_backup.csv"
NOME_ARQUIVO_SAIDA_FINAL = "Resultado_inscricoes.xlsx"
NOME_ARQUIVO_SAIDA_FINAL_CSV = "Resultado_inscricoes_backup.csv"

ARQUIVO_LOG_EXCEL = os.path.join(PASTA_LOGS, "log_processamento.xlsx")
ARQUIVO_LOG_CSV = os.path.join(PASTA_LOGS, "log_processamento.csv")

# Nova coluna "Email" incluída logo após "Celular", conforme solicitado.
ORDEM_COLUNAS = [
    "CPF", "Nome_Completo", "Data_Nascimento", "Celular", "Email", "CEP", "Bairro",
    "Estado", "Cidade", "Quantidade_Inscricoes", "Possui_Multiplas", "Status",
    "Curso", "Unidade", "Formato_Oferta", "Data primeira insc",
    "Canal primeira ins", "Data segunda insc", "Canal segunda ins",
    "Pagamento", "Vestibular", "Contrato", "Matricula", "Data_Processamento",
]

# Trava usada para proteger a leitura/escrita dos arquivos de saída e de log
# quando mais de um perfil (mais de uma janela do Chrome) roda em paralelo.
LOCK_ARQUIVOS = threading.Lock()


# ==============================================================
# BLOCO 1 - NAVEGAÇÃO NO CRM
# ==============================================================

def criar_pasta_execucao():
    """
    Cria (e retorna) uma subpasta dentro de Saida com a data/hora dessa
    importação (ex: Saida/20-07-2026_15-26-40) — assim os resultados de
    cada importação ficam separados dos de importações anteriores, em vez
    de todos se misturarem na mesma pasta Saida.
    """
    marca_tempo = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    pasta = os.path.join(PASTA_SAIDA, marca_tempo)
    os.makedirs(pasta, exist_ok=True)
    return pasta


def esperar_crm(driver):
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )


def ler_cpfs(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)

    if "CPF" not in df.columns:
        raise Exception("Coluna CPF não encontrada")

    return (
        df["CPF"]
        .dropna()
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.zfill(11)
        .tolist()
    )


def preencher_cpf(driver, cpf):
    campo = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable(
            (By.ID, "PAPrefix0ee7fe8d392a5afield-3__control")
        )
    )

    campo.click()
    campo.send_keys(Keys.CONTROL, "a")
    campo.send_keys(Keys.DELETE)

    cpf = str(cpf).zfill(11)
    cpf_formatado = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

    campo.send_keys(cpf_formatado)
    campo.send_keys(Keys.TAB)

    WebDriverWait(driver, 10).until(
        lambda d: campo.get_attribute("value") == cpf_formatado
    )


def clicar_buscar(driver):
    botoes = driver.find_elements(By.TAG_NAME, "button")

    for botao in botoes:
        texto = botao.text.strip()

        if "Buscar inscrição" in texto:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", botao
            )
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", botao)
            return

    raise Exception("Botão Buscar inscrição não encontrado")


def verificar_resultado_busca(driver):
    tempo_inicio = datetime.now()

    while True:
        cards = driver.find_elements(
            By.CSS_SELECTOR, "[data-control-name='container_inscricao']"
        )
        if len(cards) > 0:
            return "INSCRICAO"

        textos = driver.find_elements(By.TAG_NAME, "pre")
        for texto in textos:
            mensagem = texto.text.strip()
            if "Não foram encontradas inscrições" in mensagem:
                return "SEM_INSCRICAO"

        segundos = (datetime.now() - tempo_inicio).total_seconds()
        if segundos > 25:
            print("⚠️ Tempo excedido aguardando retorno")
            return "TIMEOUT"

        time.sleep(0.5)


def voltar_tela_inicial(driver):
    driver.get(URL_CRM)
    WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable(
            (By.ID, "PAPrefix0ee7fe8d392a5afield-3__control")
        )
    )


# ==============================================================
# BLOCO 2 - SCREENSHOT DE ERRO (para depuração)
# ==============================================================

def capturar_screenshot_erro(driver, cpf):
    """Salva um print da tela no momento de uma falha definitiva."""
    try:
        os.makedirs(PASTA_SCREENSHOTS, exist_ok=True)
        marca_tempo = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho = os.path.join(PASTA_SCREENSHOTS, f"{cpf}_{marca_tempo}.png")
        driver.save_screenshot(caminho)
        print(f"🖼️ Screenshot de erro salva: {caminho}")
    except Exception:
        pass


def limpar_imagens_log():
    """Apaga todos os screenshots de erro salvos em Logs/Screenshots."""
    if not os.path.exists(PASTA_SCREENSHOTS):
        return 0

    arquivos = glob.glob(os.path.join(PASTA_SCREENSHOTS, "*.png"))

    apagados = 0
    for arquivo in arquivos:
        try:
            os.remove(arquivo)
            apagados += 1
        except Exception:
            pass

    return apagados


# ==============================================================
# BLOCO 2B - FUNÇÕES DE LIMPEZA / FORMATAÇÃO DE DADOS (NOVO)
# ==============================================================

def limpar_celular(numero):
    """
    Remove espaços e caracteres especiais do telefone e garante o
    prefixo do Brasil (55) na frente, ex: '(11) 94567-9896' -> '5511945679896'.
    """
    if not numero:
        return ""

    digitos = re.sub(r"\D", "", str(numero))

    if not digitos:
        return ""

    # Se já vier com o DDI 55 e o tamanho for compatível com DDD+numero+DDI,
    # não duplica o prefixo.
    if digitos.startswith("55") and len(digitos) >= 12:
        return digitos

    return "55" + digitos


def somente_data(valor):
    """
    Recebe algo como '18/07/2026 15:50' e devolve apenas '18/07/2026'.
    Também aceita valores já sem hora, ou vazios.
    """
    if not valor:
        return ""

    valor = str(valor).strip()
    return valor.split(" ")[0].strip()


# ==============================================================
# BLOCO 3 - CAPTURA DE DADOS (INSCRIÇÕES + DADOS PESSOAIS)
# ==============================================================

def capturar_inscricoes(driver, cpf):
    resultados = []
    try:
        cards = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "[data-control-name='container_inscricao']")
            )
        )

        quantidade = len(cards)

        for card in cards:
            textos = card.find_elements(By.TAG_NAME, "pre")
            valores = [t.text.strip() for t in textos if t.text.strip()]

            dados = {
                "CPF_Pesquisado": cpf,
                "Quantidade_Inscricoes": quantidade,
                "Possui_Multiplas": "SIM" if quantidade > 1 else "NÃO",
                "Status": "", "Curso": "", "Unidade": "", "Formato_Oferta": "",
                "Pagamento": "", "Vestibular": "", "Contrato": "", "Matricula": "",
                "Data_Inscricao": "", "Canal": "",
            }

            for i, texto in enumerate(valores):
                if i + 1 >= len(valores):
                    continue

                campo = texto.strip()

                if campo == "Data de Inscrição:":
                    dados["Data_Inscricao"] = valores[i + 1].strip()
                elif campo == "Canal:":
                    dados["Canal"] = valores[i + 1].strip()
                elif campo == "Curso:":
                    dados["Curso"] = valores[i + 1].strip()
                elif campo == "Unidade:":
                    dados["Unidade"] = valores[i + 1].strip()
                elif campo == "Formato de Oferta:":
                    dados["Formato_Oferta"] = valores[i + 1].strip()
                elif campo == "Pagamento:":
                    dados["Pagamento"] = valores[i + 1].strip()
                elif campo == "Vestibular:":
                    dados["Vestibular"] = valores[i + 1].strip()
                elif campo == "Contrato:":
                    dados["Contrato"] = valores[i + 1].strip()
                elif campo == "Matricula:":
                    dados["Matricula"] = valores[i + 1].strip()

            if len(valores) > 1:
                dados["Status"] = valores[1].strip()

            resultados.append(dados)

        return resultados

    except Exception:
        return []


def abrir_primeiro_cadastro(driver):
    try:
        botao = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[.//span[contains(normalize-space(),'Abrir')]]")
            )
        )

        abas_antes = driver.window_handles
        driver.execute_script("arguments[0].click();", botao)
        driver.abas_antes_cadastro = abas_antes
        return True

    except Exception:
        return False


def capturar_dados_pessoais(driver):
    dados = {
        "Nome": "", "Sobrenome": "", "Data_Nascimento": "", "Celular": "",
        "Email": "", "CEP": "", "Bairro": "", "Estado": "", "Cidade": "",
    }

    def capturar_input(xpath):
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            def valor_preenchido(d):
                valor = d.find_element(By.XPATH, xpath).get_attribute("value")
                return valor.strip() if valor else False

            try:
                return WebDriverWait(driver, 8).until(valor_preenchido)
            except Exception:
                return driver.find_element(By.XPATH, xpath).get_attribute("value") or ""

        except Exception:
            return ""

    try:
        nome = driver.find_element(By.XPATH, "//input[@aria-label='Nome']")
        valor = ""
        for _ in range(20):
            valor = nome.get_attribute("value")
            if valor and valor.strip():
                break
            time.sleep(0.5)
        dados["Nome"] = valor
    except Exception:
        pass

    try:
        sobrenome = driver.find_element(By.XPATH, "//input[@aria-label='Sobrenome']")
        valor = ""
        for _ in range(20):
            valor = sobrenome.get_attribute("value")
            if valor and valor.strip():
                break
            time.sleep(0.5)
        dados["Sobrenome"] = valor
    except Exception:
        pass

    dados["Data_Nascimento"] = capturar_input("//input[@aria-label='Data de Nascimento']")

    # Telefone já sai limpo e com o 55 na frente.
    dados["Celular"] = limpar_celular(capturar_input("//input[@aria-label='Celular']"))

    # NOVO: captura do e-mail (mesmo padrão de input usado pelo CRM,
    # identificado pelo aria-label "E-mail").
    dados["Email"] = capturar_input("//input[@aria-label='E-mail']").strip()

    dados["CEP"] = capturar_input("//input[@aria-label='CEP']")
    dados["Bairro"] = capturar_input("//input[@aria-label='Bairro']")

    try:
        dados["Estado"] = driver.find_element(
            By.XPATH,
            "//label[contains(text(),'Estado')]/ancestor::div[contains(@data-lp-id,'FieldSectionItem')]//div[@role='link']/div",
        ).text
    except Exception:
        pass

    try:
        dados["Cidade"] = driver.find_element(
            By.XPATH,
            "//label[contains(text(),'Cidade')]/ancestor::div[contains(@data-lp-id,'FieldSectionItem')]//div[@role='link']/div",
        ).text
    except Exception:
        pass

    return dados


def abrir_e_capturar_dados(driver):
    aba_principal = driver.current_window_handle

    dados = {
        "Nome": "", "Sobrenome": "", "Data_Nascimento": "", "Celular": "",
        "Email": "", "CEP": "", "Bairro": "", "Estado": "", "Cidade": "",
    }

    try:
        WebDriverWait(driver, 30).until(lambda d: len(d.window_handles) > 1)

        novas_abas = [aba for aba in driver.window_handles if aba != aba_principal]

        if not novas_abas:
            raise Exception("Nova aba não encontrada")

        driver.switch_to.window(novas_abas[0])

        WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Nome']"))
        )

        dados = capturar_dados_pessoais(driver)

    except Exception:
        pass

    finally:
        abas = driver.window_handles
        for aba in abas:
            if aba != aba_principal:
                try:
                    driver.switch_to.window(aba)
                    driver.close()
                except Exception:
                    pass
        driver.switch_to.window(aba_principal)

    return dados


# ==============================================================
# BLOCO 4 - LOG DE CONTROLE DE CPF
# ==============================================================

def carregar_log():
    if not os.path.exists(ARQUIVO_LOG_EXCEL):
        return set()

    try:
        df_log = pd.read_excel(ARQUIVO_LOG_EXCEL)

        if "CPF" not in df_log.columns:
            return set()

        cpfs = (
            df_log[df_log["Status"].isin(["Sucesso", "Sem inscrição"])]["CPF"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.zfill(11)
            .tolist()
        )

        return set(cpfs)

    except Exception:
        return set()


def salvar_log(cpf, status, tempo="", mensagem_erro=""):
    """
    Protegida por LOCK_ARQUIVOS pelo chamador quando houver mais de
    um perfil rodando ao mesmo tempo.
    """
    novo_registro = pd.DataFrame(
        [
            {
                "CPF": cpf,
                "Data_Processamento": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "Status": status,
                "Tempo": tempo,
                "Mensagem_Erro": mensagem_erro,
            }
        ]
    )

    if os.path.exists(ARQUIVO_LOG_EXCEL):
        try:
            df_antigo = pd.read_excel(ARQUIVO_LOG_EXCEL)
            df_final = pd.concat([df_antigo, novo_registro], ignore_index=True)
        except Exception:
            df_final = novo_registro
    else:
        df_final = novo_registro

    df_final = df_final.drop_duplicates(subset=["CPF"], keep="last")

    df_final.to_excel(ARQUIVO_LOG_EXCEL, index=False)
    df_final.to_csv(ARQUIVO_LOG_CSV, index=False, encoding="utf-8-sig")


# ==============================================================
# BLOCO 4B - IDENTIDADE DOS PERFIS (apelido / login detectado)
# ==============================================================

ARQUIVO_PERFIS_CONFIG = os.path.join(PASTA_DADOS, "perfis_config.json")
LOCK_PERFIS_CONFIG = threading.Lock()

PERFIS_CONFIG_PADRAO = {
    str(i): {
        "apelido": "",
        "apelido_personalizado": False,
        "login_detectado": "",
        "email": "",
        "status": "Desconhecido",
        "configurado": False,
    }
    for i in (1, 2, 3)
}


def carregar_perfis_config():
    """Lê a identidade salva de cada perfil (apelido, e-mail, status de login)."""
    if not os.path.exists(ARQUIVO_PERFIS_CONFIG):
        return {k: dict(v) for k, v in PERFIS_CONFIG_PADRAO.items()}

    try:
        with open(ARQUIVO_PERFIS_CONFIG, "r", encoding="utf-8") as f:
            dados = json.load(f)
        for chave, valor_padrao in PERFIS_CONFIG_PADRAO.items():
            if chave not in dados or not isinstance(dados[chave], dict):
                dados[chave] = dict(valor_padrao)
            else:
                # Preenche com os campos padrão quaisquer chaves novas que
                # ainda não existiam num arquivo salvo por uma versão antiga.
                for campo, valor_campo in valor_padrao.items():
                    dados[chave].setdefault(campo, valor_campo)
        return dados
    except Exception:
        return {k: dict(v) for k, v in PERFIS_CONFIG_PADRAO.items()}


def salvar_perfis_config(config):
    with LOCK_PERFIS_CONFIG:
        try:
            os.makedirs(PASTA_DADOS, exist_ok=True)
            arquivo_temp = ARQUIVO_PERFIS_CONFIG + ".tmp"
            with open(arquivo_temp, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            os.replace(arquivo_temp, ARQUIVO_PERFIS_CONFIG)
        except Exception:
            pass


def definir_apelido_perfil(numero_perfil, apelido):
    """
    Salva o apelido digitado manualmente pelo usuário para um perfil (via
    lápis de edição). A partir daqui o apelido fica "travado": detecções
    automáticas de e-mail (registrar_email_detectado) não vão mais
    sobrescrevê-lo, só atualizam o e-mail salvo.
    """
    config = carregar_perfis_config()
    chave = str(numero_perfil)
    config.setdefault(chave, dict(PERFIS_CONFIG_PADRAO[chave]))
    config[chave]["apelido"] = apelido.strip()
    config[chave]["apelido_personalizado"] = True
    if apelido.strip():
        config[chave]["configurado"] = True
    salvar_perfis_config(config)


def registrar_email_detectado(numero_perfil, email):
    """
    Chamado quando a automação detecta (best-effort) o e-mail da conta
    logada naquele perfil. Por padrão o apelido do perfil passa a ser esse
    e-mail — a menos que o usuário já tenha personalizado o apelido na mão
    (nesse caso só o e-mail salvo é atualizado, o apelido não é mexido).
    """
    email = (email or "").strip()
    if not email:
        return
    config = carregar_perfis_config()
    chave = str(numero_perfil)
    config.setdefault(chave, dict(PERFIS_CONFIG_PADRAO[chave]))
    config[chave]["email"] = email
    config[chave]["login_detectado"] = email
    if not config[chave].get("apelido_personalizado"):
        config[chave]["apelido"] = email
    config[chave]["configurado"] = True
    salvar_perfis_config(config)


def definir_email_manual(numero_perfil, email):
    """Salva um e-mail digitado manualmente pelo usuário (via lápis de edição)."""
    config = carregar_perfis_config()
    chave = str(numero_perfil)
    config.setdefault(chave, dict(PERFIS_CONFIG_PADRAO[chave]))
    config[chave]["email"] = email.strip()
    config[chave]["login_detectado"] = email.strip()
    salvar_perfis_config(config)


def definir_status_perfil(numero_perfil, status):
    """Salva o status de login mais recente detectado para o perfil."""
    config = carregar_perfis_config()
    chave = str(numero_perfil)
    config.setdefault(chave, dict(PERFIS_CONFIG_PADRAO[chave]))
    config[chave]["status"] = status
    salvar_perfis_config(config)


def rotulo_perfil_configurado(numero_perfil):
    """Retorna o melhor nome disponível pra exibir/usar como rótulo desse perfil."""
    config = carregar_perfis_config()
    dados = config.get(str(numero_perfil), {})
    apelido = (dados.get("apelido") or "").strip()
    if apelido:
        return apelido
    login = (dados.get("login_detectado") or dados.get("email") or "").strip()
    if login:
        return login
    return f"Perfil {numero_perfil}"


def limpar_perfil_individual(numero_perfil):
    """
    Apaga a pasta de login do Chrome daquele perfil (força novo login) e
    reseta a identidade salva (apelido / login detectado) dele. Não mexe
    nos outros perfis.
    """
    pasta = os.path.join(PASTA_PERFIS, f"perfil_{numero_perfil}")
    if os.path.exists(pasta):
        shutil.rmtree(pasta, ignore_errors=True)

    config = carregar_perfis_config()
    config[str(numero_perfil)] = dict(PERFIS_CONFIG_PADRAO[str(numero_perfil)])
    salvar_perfis_config(config)


REGEX_EMAIL = re.compile(r"[\w.+-]+@[\w.-]+\.\w+")

# Seletores que abrem o painel de conta do cabeçalho padrão Office 365.
# Nem todos existem sempre — tenta em sequência e para no primeiro clique
# que efetivamente abrir o painel.
_SELETORES_ABRIR_PAINEL_CONTA = [
    (By.ID, "O365_MainLink_Me"),
    (By.ID, "meControlTrigger"),
    (By.CSS_SELECTOR, ".mectrl_trigger"),
    (By.ID, "mectrl_header_account"),
    (By.ID, "mectrl_currentAccount_picture"),
    (By.CSS_SELECTOR, "[aria-label*='conta' i]"),
    (By.CSS_SELECTOR, "[aria-label*='account' i]"),
]

# Domínios/URLs que indicam que o navegador está numa tela de login (ainda
# não autenticado) da Microsoft, e não dentro do CRM.
_DOMINIOS_LOGIN_MICROSOFT = (
    "login.microsoftonline.com",
    "login.live.com",
    "/oauth2/",
    "adfs/ls",
)


def _ler_email_painel_conta(driver):
    """Tenta ler #mectrl_currentAccount_secondary, se já estiver no DOM."""
    try:
        elemento = driver.find_element(By.ID, "mectrl_currentAccount_secondary")
        texto = (elemento.text or "").strip()
        if texto and REGEX_EMAIL.search(texto):
            return REGEX_EMAIL.search(texto).group(0)
    except Exception:
        pass
    return None


def _varrer_pagina_em_busca_de_email(driver):
    """
    Fallback genérico: varre atributos title/aria-label de todos os
    elementos e o texto visível da página inteira, procurando algo com
    formato de e-mail.
    """
    try:
        for atributo in ("title", "aria-label"):
            elementos = driver.find_elements(By.CSS_SELECTOR, f"[{atributo}*='@']")
            for elemento in elementos:
                valor = (elemento.get_attribute(atributo) or "").strip()
                encontrado = REGEX_EMAIL.search(valor)
                if encontrado:
                    return encontrado.group(0)
    except Exception:
        pass

    try:
        texto_pagina = driver.find_element(By.TAG_NAME, "body").text or ""
        encontrado = REGEX_EMAIL.search(texto_pagina)
        if encontrado:
            return encontrado.group(0)
    except Exception:
        pass

    return None


def capturar_email_conta(driver):
    """
    Captura (melhor esforço) o e-mail da conta logada, usando o painel de
    conta padrão do cabeçalho Office 365 do CRM.

    O painel com o e-mail (#mectrl_currentAccount_secondary) só existe no
    HTML depois que alguém clica no ícone/foto de perfil no canto superior
    direito — antes do clique, esse elemento nem está no DOM. Por isso:
      1. tenta ler o painel direto (caso já esteja aberto por algum motivo);
      2. se não achar, tenta clicar num dos seletores conhecidos de
         abertura do painel, um de cada vez, parando no primeiro que
         funcionar;
      3. tenta ler o painel de novo depois do clique;
      4. se ainda assim não achar, cai no fallback genérico (varredura de
         title/aria-label/texto da página em busca de algo com formato de
         e-mail).

    Não trava nem levanta erro se não conseguir — nesse caso a automação
    segue normalmente e o usuário pode corrigir o e-mail manualmente na
    tela Perfis.
    """
    email = _ler_email_painel_conta(driver)
    if email:
        return email

    for by, seletor in _SELETORES_ABRIR_PAINEL_CONTA:
        try:
            elemento = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((by, seletor))
            )
            elemento.click()
        except Exception:
            continue

        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, "mectrl_currentAccount_secondary"))
            )
        except Exception:
            pass

        email = _ler_email_painel_conta(driver)
        if email:
            return email

    return _varrer_pagina_em_busca_de_email(driver)


def detectar_status_login(driver, tempo_espera=4):
    """
    Decide (melhor esforço) se o perfil está logado no CRM, sem precisar de
    credenciais: parte do princípio que o driver já navegou pra uma URL que
    exige login (ver esperar_crm / iniciar_chrome), espera alguns segundos
    e olha pra onde o navegador foi parar.

    Retorna uma das strings: "Logado", "Não logado", "Não foi possível
    verificar". Nunca levanta exceção.
    """
    try:
        time.sleep(tempo_espera)
        url_atual = (driver.current_url or "").lower()

        if any(dominio in url_atual for dominio in _DOMINIOS_LOGIN_MICROSOFT):
            return "Não logado"

        try:
            campo_senha = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
            if campo_senha.is_displayed():
                return "Não logado"
        except Exception:
            pass

        if "crm2.dynamics.com" in url_atual:
            return "Logado"

        return "Não foi possível verificar"
    except Exception:
        return "Não foi possível verificar"


# ==============================================================
# BLOCO 5 - RESULTADO PARCIAL / FINAL
# ==============================================================

def salvar_resultado_parcial(resultados, pasta_saida):
    df = pd.DataFrame(resultados)

    if df.empty:
        return

    if "Nome_Completo" not in df.columns:
        df["Nome_Completo"] = ""

    if "Nome" in df.columns:
        nome_novo = df["Nome"].fillna("").astype(str).str.strip()

        if "Sobrenome" in df.columns:
            nome_novo = (
                nome_novo + " " + df["Sobrenome"].fillna("").astype(str).str.strip()
            ).str.strip()

        df["Nome_Completo"] = (
            df["Nome_Completo"]
            .fillna("")
            .astype(str)
            .str.strip()
            .where(lambda x: x != "", nome_novo)
        )

    df.drop(columns=["Nome", "Sobrenome"], inplace=True, errors="ignore")

    if "Data_Processamento" not in df.columns:
        df["Data_Processamento"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    if "CPF" in df.columns:
        df["CPF"] = (
            df["CPF"].astype(str).str.replace(".0", "", regex=False).str.zfill(11)
        )
        df = df.drop_duplicates(subset=["CPF"], keep="last")

    # Garante que as colunas de data fiquem apenas com a data (sem hora),
    # mesmo que algum registro antigo do arquivo parcial ainda tenha hora.
    for coluna in ["Data primeira insc", "Data segunda insc"]:
        if coluna in df.columns:
            df[coluna] = df[coluna].apply(somente_data)

    # Garante telefone limpo mesmo em registros antigos do arquivo parcial.
    if "Celular" in df.columns:
        df["Celular"] = df["Celular"].apply(limpar_celular)

    df = df[[coluna for coluna in ORDEM_COLUNAS if coluna in df.columns]]

    os.makedirs(pasta_saida, exist_ok=True)
    arquivo_parcial = os.path.join(pasta_saida, NOME_ARQUIVO_PARCIAL)
    arquivo_temp = os.path.join(pasta_saida, "Resultado_parcial_temp.xlsx")
    df.to_excel(arquivo_temp, index=False)
    os.replace(arquivo_temp, arquivo_parcial)

    df.to_csv(os.path.join(pasta_saida, NOME_ARQUIVO_BACKUP_CSV), index=False, encoding="utf-8-sig")


def carregar_resultado_parcial(pasta_saida):
    arquivo_parcial = os.path.join(pasta_saida, NOME_ARQUIVO_PARCIAL)
    if not os.path.exists(arquivo_parcial):
        return []

    try:
        df = pd.read_excel(arquivo_parcial)
        return df.to_dict("records")
    except Exception:
        return []


def gerar_resultado_final(resultado_final, pasta_saida):
    df_final = pd.DataFrame(resultado_final)

    if df_final.empty:
        print("⚠️ Nenhum registro processado")
        return

    if "Nome_Completo" not in df_final.columns:
        if "Nome" in df_final.columns and "Sobrenome" in df_final.columns:
            df_final["Nome_Completo"] = (
                df_final["Nome"].fillna("").astype(str)
                + " "
                + df_final["Sobrenome"].fillna("").astype(str)
            ).str.strip()

    df_final.drop(columns=["Nome", "Sobrenome"], inplace=True, errors="ignore")

    for coluna in ["Data primeira insc", "Data segunda insc"]:
        if coluna in df_final.columns:
            df_final[coluna] = df_final[coluna].apply(somente_data)

    if "Celular" in df_final.columns:
        df_final["Celular"] = df_final["Celular"].apply(limpar_celular)

    df_final = df_final[[c for c in ORDEM_COLUNAS if c in df_final.columns]]

    os.makedirs(pasta_saida, exist_ok=True)
    df_final.to_excel(os.path.join(pasta_saida, NOME_ARQUIVO_SAIDA_FINAL), index=False)
    df_final.to_csv(os.path.join(pasta_saida, NOME_ARQUIVO_SAIDA_FINAL_CSV), index=False, encoding="utf-8-sig")

    print("📁 Arquivo final atualizado")


def finalizar_processamento_geral(pasta_saida):
    """
    Consolida o arquivo final UMA ÚNICA VEZ, depois que TODOS os perfis
    ativos já terminaram suas execuções. Chamada pela interface (App),
    nunca de dentro de executar_processamento.
    """
    print()
    print("==============================")
    print("🏁 TODOS OS PERFIS CONCLUÍDOS")
    print("==============================")

    with LOCK_ARQUIVOS:
        gerar_resultado_final(carregar_resultado_parcial(pasta_saida), pasta_saida)


# ==============================================================
# EXECUÇÃO PRINCIPAL (usada pela interface)
# ==============================================================

def iniciar_chrome(pasta_perfil, posicao=0, headless=False):
    options = webdriver.ChromeOptions()
    profile = os.path.abspath(pasta_perfil)
    os.makedirs(profile, exist_ok=True)
    options.add_argument(f"--user-data-dir={profile}")
    options.add_argument(f"--window-position={posicao * 60},{posicao * 60}")

    if headless:
        # "--headless=new" mantém o mesmo motor de renderização do Chrome
        # normal (o painel de conta e os seletores continuam funcionando
        # igual), só não abre uma janela visível na tela.
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1366,900")
        options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )

    driver.get(URL_CRM)
    return driver


def executar_processamento(
    cpfs,
    pasta_perfil,
    stop_event,
    atualizar_progresso,
    atualizar_stats,
    pasta_saida,
    posicao_janela=0,
    rotulo_perfil="",
    numero_perfil=None,
):
    """
    Processa uma lista de CPFs usando um perfil (pasta) específico do Chrome.
    Pode ser chamada várias vezes em paralelo (uma thread por perfil) —
    por isso toda leitura/escrita de arquivo compartilhado passa pelo
    LOCK_ARQUIVOS. Os resultados dessa importação são salvos em
    pasta_saida (uma subpasta com data/hora, criada por criar_pasta_execucao,
    exclusiva desta importação).
    """
    os.makedirs(PASTA_ENTRADA, exist_ok=True)
    os.makedirs(pasta_saida, exist_ok=True)
    os.makedirs(PASTA_LOGS, exist_ok=True)

    prefixo = f"[{rotulo_perfil}] " if rotulo_perfil else ""
    driver = None

    try:
        driver = iniciar_chrome(pasta_perfil, posicao=posicao_janela)
        esperar_crm(driver)

        if numero_perfil is not None:
            try:
                status = detectar_status_login(driver, tempo_espera=2)
                definir_status_perfil(numero_perfil, status)
                if status == "Logado":
                    email_detectado = capturar_email_conta(driver)
                    if email_detectado:
                        registrar_email_detectado(numero_perfil, email_detectado)
                        print(f"{prefixo}🔑 Login detectado: {email_detectado}")
            except Exception:
                pass

        with LOCK_ARQUIVOS:
            cpfs_processados = carregar_log()

        total_original = len(cpfs)
        cpfs = [cpf for cpf in cpfs if cpf not in cpfs_processados]

        print(f"{prefixo}📋 CPFs recebidos: {total_original}")
        print(f"{prefixo}♻️ Ignorados pelo log: {total_original - len(cpfs)}")
        print(f"{prefixo}🚀 Processando: {len(cpfs)}")

        inicio_execucao = datetime.now()

        sucessos = 0
        erros = 0
        sem_inscricao = 0
        total = len(cpfs)

        atualizar_progresso(0, total)
        atualizar_stats(sucessos, sem_inscricao, erros)

        for indice, cpf in enumerate(cpfs):
            if stop_event.is_set():
                print(f"{prefixo}⏹ Processamento interrompido pelo usuário.")
                break

            inicio_cpf = datetime.now()
            processado = False
            tentativa = 0

            while tentativa < 3 and not processado:
                tentativa += 1

                try:
                    if tentativa > 1:
                        try:
                            driver.refresh()
                            time.sleep(3)
                            esperar_crm(driver)
                        except Exception:
                            voltar_tela_inicial(driver)

                    preencher_cpf(driver, cpf)
                    clicar_buscar(driver)

                    resultado_busca = verificar_resultado_busca(driver)

                    if resultado_busca in ["TIMEOUT", "erro"]:
                        raise Exception("Busca sem retorno")

                    if resultado_busca == "SEM_INSCRICAO":
                        with LOCK_ARQUIVOS:
                            salvar_log(cpf, "Sem inscrição")
                        sem_inscricao += 1
                        processado = True
                        voltar_tela_inicial(driver)
                        continue

                    inscricoes = capturar_inscricoes(driver, cpf)

                    if not inscricoes:
                        raise Exception("Nenhuma inscrição capturada")

                    dados_inscricao = {
                        "Quantidade_Inscricoes": len(inscricoes),
                        "Possui_Multiplas": "SIM" if len(inscricoes) > 1 else "NÃO",
                        "Status": "", "Curso": "", "Unidade": "", "Formato_Oferta": "",
                        "Data primeira insc": "", "Canal primeira ins": "",
                        "Data segunda insc": "", "Canal segunda ins": "",
                        "Pagamento": "", "Vestibular": "", "Contrato": "", "Matricula": "",
                    }

                    primeira = inscricoes[0]

                    for campo in [
                        "Status", "Curso", "Unidade", "Formato_Oferta",
                        "Pagamento", "Vestibular", "Contrato", "Matricula",
                    ]:
                        dados_inscricao[campo] = primeira.get(campo, "")

                    # Datas gravadas somente com o dia (sem hora), conforme solicitado.
                    dados_inscricao["Data primeira insc"] = somente_data(primeira.get("Data_Inscricao", ""))
                    dados_inscricao["Canal primeira ins"] = primeira.get("Canal", "")

                    if len(inscricoes) > 1:
                        segunda = inscricoes[1]
                        dados_inscricao["Data segunda insc"] = somente_data(segunda.get("Data_Inscricao", ""))
                        dados_inscricao["Canal segunda ins"] = segunda.get("Canal", "")

                    dados_pessoais = {
                        "Nome": "", "Sobrenome": "", "Data_Nascimento": "",
                        "Celular": "", "Email": "", "CEP": "", "Bairro": "",
                        "Estado": "", "Cidade": "",
                    }

                    if abrir_primeiro_cadastro(driver):
                        dados_pessoais = abrir_e_capturar_dados(driver)
                    else:
                        raise Exception("Cadastro não abriu")

                    if not dados_pessoais.get("Nome") and not dados_pessoais.get("Sobrenome"):
                        raise Exception("Nome não capturado")

                    voltar_tela_inicial(driver)

                    linha_final = {
                        "CPF": cpf,
                        **dados_pessoais,
                        **dados_inscricao,
                        "Data_Processamento": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    }

                    with LOCK_ARQUIVOS:
                        resultado_atual = carregar_resultado_parcial(pasta_saida)
                        resultado_atual.append(linha_final)
                        salvar_resultado_parcial(resultado_atual, pasta_saida)
                        salvar_log(cpf, "Sucesso", tempo=str(datetime.now() - inicio_cpf))

                    sucessos += 1
                    processado = True

                except Exception as erro:
                    if tentativa >= 3:
                        erros += 1
                        capturar_screenshot_erro(driver, cpf)
                        with LOCK_ARQUIVOS:
                            salvar_log(
                                cpf, "Erro",
                                tempo=str(datetime.now() - inicio_cpf),
                                mensagem_erro=str(erro),
                            )
                        processado = True

            if processado:
                atualizar_progresso(indice + 1, total)
                atualizar_stats(sucessos, sem_inscricao, erros)

        tempo_total = datetime.now() - inicio_execucao

        print()
        print(f"{prefixo}==============================")
        print(f"{prefixo}✅ Perfil concluído. Aguardando os demais perfis...")
        print(f"{prefixo}==============================")
        print(f"{prefixo}✅ Sucesso: {sucessos}")
        print(f"{prefixo}⚪ Sem inscrição: {sem_inscricao}")
        print(f"{prefixo}❌ Erros: {erros}")
        print(f"{prefixo}⏱ Tempo total: {tempo_total}")

        # IMPORTANTE: o arquivo final NÃO é gerado aqui. Cada perfil só
        # atualiza o resultado parcial (compartilhado). O arquivo final só
        # é consolidado depois que TODOS os perfis ativos tiverem terminado
        # — isso é feito de forma centralizada pela interface, na função
        # finalizar_processamento_geral().

    except Exception as erro:
        print(f"{prefixo}💥 Erro fatal: {erro}")

    finally:
        # Fecha a janela deste perfil assim que ele termina, para liberar
        # recursos — os demais perfis continuam rodando normalmente.
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass



# ==============================================================
# BLOCO 6 - HISTÓRICO DE EXECUÇÕES (para o painel "Início")
# ==============================================================

ARQUIVO_HISTORICO = os.path.join(PASTA_LOGS, "historico_execucoes.json")
LOCK_HISTORICO = threading.Lock()


def carregar_historico():
    """Lê o histórico de execuções salvo em disco (lista de dicts)."""
    if not os.path.exists(ARQUIVO_HISTORICO):
        return []
    try:
        with open(ARQUIVO_HISTORICO, "r", encoding="utf-8") as f:
            dados = json.load(f)
            return dados if isinstance(dados, list) else []
    except Exception:
        return []


def salvar_historico(lista):
    """Grava o histórico de execuções, mantendo só os últimos 50 registros."""
    try:
        os.makedirs(PASTA_LOGS, exist_ok=True)
        lista_reduzida = lista[-50:]
        arquivo_temp = ARQUIVO_HISTORICO + ".tmp"
        with open(arquivo_temp, "w", encoding="utf-8") as f:
            json.dump(lista_reduzida, f, ensure_ascii=False, indent=2)
        os.replace(arquivo_temp, ARQUIVO_HISTORICO)
    except Exception:
        pass


def registrar_inicio_execucao(qtd_perfis):
    """Cria e persiste um novo registro de execução com status 'Em andamento'."""
    with LOCK_HISTORICO:
        historico = carregar_historico()
        registro = {
            "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "inicio": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "fim": None,
            "perfis": qtd_perfis,
            "status": "Em andamento",
            "processados": 0,
            "sucessos": 0,
            "sem_inscricao": 0,
            "erros": 0,
            "duracao": None,
        }
        historico.append(registro)
        salvar_historico(historico)
        return registro["id"]


def atualizar_execucao(id_execucao, **campos):
    """Atualiza campos de um registro específico do histórico (por id)."""
    with LOCK_HISTORICO:
        historico = carregar_historico()
        for registro in historico:
            if registro.get("id") == id_execucao:
                registro.update(campos)
                break
        salvar_historico(historico)


def finalizar_execucao(id_execucao, sucessos, sem_inscricao, erros, interrompida=False):
    """Marca um registro do histórico como concluído (ou interrompido) e calcula a duração."""
    with LOCK_HISTORICO:
        historico = carregar_historico()
        for registro in historico:
            if registro.get("id") == id_execucao:
                agora = datetime.now()
                try:
                    inicio_dt = datetime.strptime(registro["inicio"], "%d/%m/%Y %H:%M:%S")
                    duracao = agora - inicio_dt
                    horas, resto = divmod(int(duracao.total_seconds()), 3600)
                    minutos, segundos = divmod(resto, 60)
                    duracao_fmt = f"{horas:02d}:{minutos:02d}:{segundos:02d}"
                except Exception:
                    duracao_fmt = ""

                registro.update({
                    "fim": agora.strftime("%d/%m/%Y %H:%M:%S"),
                    "status": "Interrompido" if interrompida else "Concluído",
                    "processados": sucessos + sem_inscricao + erros,
                    "sucessos": sucessos,
                    "sem_inscricao": sem_inscricao,
                    "erros": erros,
                    "duracao": duracao_fmt,
                })
                break
        salvar_historico(historico)


def total_dados_capturados():
    """Total histórico de CPFs processados com sucesso (todas as execuções já feitas)."""
    if not os.path.exists(ARQUIVO_LOG_EXCEL):
        return 0
    try:
        df = pd.read_excel(ARQUIVO_LOG_EXCEL)
        if "Status" not in df.columns:
            return 0
        return int((df["Status"] == "Sucesso").sum())
    except Exception:
        return 0


def calcular_estatisticas_hoje():
    """Retorna (execuções hoje, tempo total hoje formatado H:MM:SS)."""
    historico = carregar_historico()
    hoje = datetime.now().strftime("%d/%m/%Y")

    execucoes_hoje = 0
    segundos_totais = 0

    for registro in historico:
        inicio = registro.get("inicio") or ""
        if not inicio.startswith(hoje):
            continue
        execucoes_hoje += 1

        duracao = registro.get("duracao")
        if duracao:
            try:
                partes = [int(p) for p in duracao.split(":")]
                if len(partes) == 3:
                    segundos_totais += partes[0] * 3600 + partes[1] * 60 + partes[2]
            except Exception:
                pass
        elif registro.get("status") == "Em andamento":
            try:
                inicio_dt = datetime.strptime(registro["inicio"], "%d/%m/%Y %H:%M:%S")
                segundos_totais += int((datetime.now() - inicio_dt).total_seconds())
            except Exception:
                pass

    horas, resto = divmod(segundos_totais, 3600)
    minutos, segundos = divmod(resto, 60)
    tempo_fmt = f"{horas:02d}:{minutos:02d}:{segundos:02d}"

    return execucoes_hoje, tempo_fmt

# ==============================================================
# INTERFACE GRÁFICA (CUSTOMTKINTER)
# ==============================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ---------------------------------------------------------------
# Paleta de temas. Cada tema define modo (light/dark) e cores.
# Trocar de tema chama ctk.set_appearance_mode(...) e reconfigura
# todos os widgets registrados em self._temaveis. Os dois temas
# escuros mantêm as cores originais; o tema Claro usa a paleta
# oficial do produto (laranja / creme / cinza escuro / gelo) — nele,
# qualquer fundo claro SEMPRE usa texto em cinza escuro (#525252),
# nunca texto claro, para garantir contraste e legibilidade.
# ---------------------------------------------------------------
TEMAS = {
    "Roxo Escuro": {
        "modo": "dark",
        "fundo": "#0f1117",
        "card": "#181b24",
        "sidebar": "#14161f",
        "accent": "#6c5ce7",
        "accent_hover": "#5843d6",
        "accent_texto": "#ffffff",
        "secundario": "#2b2f3a",
        "secundario_hover": "#3a3f4d",
        "texto_principal": "#ffffff",
        "texto_secundario": "#9aa0ac",
        "texto_no_accent": "#ffffff",
        "log_fundo": "#0b0d12",
        "log_texto": "#c7cbd4",
        "log_hora": "#5c6270",
        "borda": "#2b2f3a",
    },
    "Meia-noite Azul": {
        "modo": "dark",
        "fundo": "#0a0e1a",
        "card": "#111827",
        "sidebar": "#0d1320",
        "accent": "#3b82f6",
        "accent_hover": "#2563eb",
        "accent_texto": "#ffffff",
        "secundario": "#1f2937",
        "secundario_hover": "#2b394d",
        "texto_principal": "#f1f5f9",
        "texto_secundario": "#94a3b8",
        "texto_no_accent": "#ffffff",
        "log_fundo": "#060a12",
        "log_texto": "#cbd5e1",
        "log_hora": "#475569",
        "borda": "#1f2937",
    },
    "Claro": {
        # Paleta oficial do produto (print de referência).
        "modo": "light",
        "fundo": "#FEF8F1",          # creme
        "card": "#FCFAF9",           # gelo
        "sidebar": "#FCFAF9",        # gelo
        "accent": "#FE8310",         # laranja
        "accent_hover": "#E3730A",
        "accent_texto": "#ffffff",   # laranja é escuro o suficiente p/ texto branco
        "secundario": "#FBEEE1",     # laranja bem clarinho
        "secundario_hover": "#F3DFC8",
        "texto_principal": "#525252",   # cinza escuro (regra de contraste)
        "texto_secundario": "#8A8A8A",
        "texto_no_accent": "#ffffff",
        "log_fundo": "#FCFAF9",
        "log_texto": "#525252",
        "log_hora": "#ADADAD",
        "borda": "#F0E4D8",
    },
}

COR_SUCESSO = "#2ecc71"
COR_ALERTA = "#f1c40f"
COR_ERRO = "#e74c3c"
COR_PARAR = "#e74c3c"
COR_PARAR_HOVER = "#c0392b"

ITENS_MENU = [
    ("inicio", "🏠", "Início"),
    ("perfis", "👤", "Perfis"),
    ("execucoes", "✅", "Execuções"),
    ("configuracoes", "⚙️", "Configurações"),
    ("logs", "🖥️", "Logs"),
    ("suporte", "❓", "Suporte"),
    ("sobre", "ℹ️", "Sobre"),
]


class TextoRedirecionado(io.TextIOBase):
    """
    Redireciona chamadas de print() para uma fila lida pela GUI.
    Guarda o texto em buffer e só envia para a fila quando uma linha
    é fechada (quebra de linha). Como várias threads (perfis) podem
    imprimir ao mesmo tempo, o write() é protegido por uma trava.
    """

    def __init__(self, fila):
        self.fila = fila
        self._buffer = ""
        self._lock = threading.Lock()

    def write(self, texto):
        with self._lock:
            self._buffer += texto
            while "\n" in self._buffer:
                linha, self._buffer = self._buffer.split("\n", 1)
                self.fila.put(("log", linha))
        return len(texto)

    def flush(self):
        with self._lock:
            if self._buffer:
                self.fila.put(("log", self._buffer))
                self._buffer = ""


class CartaoIcone(ctk.CTkFrame):
    """Card do topo do dashboard: ícone num quadrado colorido, valor grande e título."""

    def __init__(self, master, emoji, titulo, **kwargs):
        super().__init__(master, corner_radius=14, **kwargs)

        self.icone_frame = ctk.CTkFrame(self, width=46, height=46, corner_radius=12)
        self.icone_frame.pack(anchor="w", padx=18, pady=(18, 10))
        self.icone_frame.pack_propagate(False)

        self.icone_label = ctk.CTkLabel(self.icone_frame, text=emoji, font=ctk.CTkFont(size=20))
        self.icone_label.pack(expand=True)

        self.valor_label = ctk.CTkLabel(self, text="0", font=ctk.CTkFont(size=25, weight="bold"))
        self.valor_label.pack(anchor="w", padx=18)

        self.titulo_label = ctk.CTkLabel(
            self, text=titulo, font=ctk.CTkFont(size=12), justify="left", wraplength=190, anchor="w",
        )
        self.titulo_label.pack(anchor="w", padx=18, pady=(2, 18), fill="x")

    def atualizar(self, valor):
        self.valor_label.configure(text=str(valor))

    def aplicar_tema(self, paleta):
        self.configure(fg_color=paleta["card"])
        self.icone_frame.configure(fg_color=paleta["accent"])
        self.icone_label.configure(text_color=paleta["accent_texto"])
        self.valor_label.configure(text_color=paleta["texto_principal"])
        self.titulo_label.configure(text_color=paleta["texto_secundario"])


class CartaoStat(ctk.CTkFrame):
    """Card pequeno de estatística (usado em Processados / Sucessos / Pendentes / Erros)."""

    def __init__(self, master, titulo, cor, emoji, **kwargs):
        super().__init__(master, corner_radius=14, **kwargs)

        self.emoji_label = ctk.CTkLabel(self, text=emoji, font=ctk.CTkFont(size=24))
        self.emoji_label.pack(pady=(14, 0))

        self.valor_label = ctk.CTkLabel(
            self, text="0", font=ctk.CTkFont(size=24, weight="bold"), text_color=cor
        )
        self.valor_label.pack()

        self.titulo_label = ctk.CTkLabel(self, text=titulo, font=ctk.CTkFont(size=12))
        self.titulo_label.pack(pady=(0, 14))

    def atualizar(self, valor):
        self.valor_label.configure(text=str(valor))

    def aplicar_tema(self, paleta):
        self.configure(fg_color=paleta["card"])
        self.titulo_label.configure(text_color=paleta["texto_secundario"])


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(NOME_PRODUTO)
        self.geometry("1180x860")
        self.minsize(1020, 700)
        self._aplicar_icone_janela()

        # ---------------- Estado geral ----------------
        self.fila = queue.Queue()
        self.stop_event = threading.Event()
        self.threads_execucao = []
        self.threads_ativas = 0
        self.caminho_cpfs = ctk.StringVar(value=ARQUIVO_CPFS_PADRAO)
        self.pasta_saida_atual = None
        self.tema_atual = ctk.StringVar(value="Roxo Escuro")

        # Um checkbox por perfil (permite qualquer combinação: 1; 1 e 3; 2 e 3; etc.)
        self.perfil_selecionado = {1: ctk.BooleanVar(value=True), 2: ctk.BooleanVar(value=False), 3: ctk.BooleanVar(value=False)}
        # Campo de apelido editável de cada perfil (ex: "Ricardo")
        self.perfil_apelido_var = {1: ctk.StringVar(), 2: ctk.StringVar(), 3: ctk.StringVar()}

        self._temaveis = []           # (widget, {parametro: chave_paleta})
        self._progresso_por_perfil = {}
        self._stats_por_perfil = {}
        self._botoes_menu = {}
        self._paginas = {}
        self.pagina_atual = "inicio"

        self.execucao_atual = None    # dict com dados da execução em andamento (ou None)
        self._historico_cache = carregar_historico()

        self._montar_layout()
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)
        self._aplicar_tema("Roxo Escuro")
        self._carregar_perfis_na_tela()
        self._mostrar_pagina("inicio")
        self._atualizar_cards_topo()
        self._atualizar_tabelas_execucoes()
        self._processar_fila()
        self._tick_relogio()

    # =====================================================
    # ÍCONE / LOGO
    # =====================================================
    def _carregar_logo(self, tamanho=(42, 42)):
        try:
            from PIL import Image
            caminho_png = caminho_recurso(os.path.join("assets", "icone_captador.png"))
            if not os.path.exists(caminho_png):
                return None
            imagem_pil = Image.open(caminho_png)
            return ctk.CTkImage(light_image=imagem_pil, dark_image=imagem_pil, size=tamanho)
        except Exception:
            return None

    def _aplicar_icone_janela(self):
        try:
            caminho_ico = caminho_recurso(os.path.join("assets", "icone_captador.ico"))
            if os.path.exists(caminho_ico):
                self.iconbitmap(caminho_ico)
                return
        except Exception:
            pass

        try:
            caminho_png = caminho_recurso(os.path.join("assets", "icone_captador.png"))
            if os.path.exists(caminho_png):
                from tkinter import PhotoImage
                self._icone_png_ref = PhotoImage(file=caminho_png)
                self.iconphoto(True, self._icone_png_ref)
        except Exception:
            pass

    # =====================================================
    # TEMA
    # =====================================================
    def _registrar_tema(self, widget, **mapa):
        self._temaveis.append((widget, mapa))

    def _aplicar_tema(self, nome):
        paleta = TEMAS[nome]
        ctk.set_appearance_mode(paleta["modo"])
        self.configure(fg_color=paleta["fundo"])

        for widget, mapa in self._temaveis:
            props = {parametro: paleta[chave] for parametro, chave in mapa.items()}
            try:
                widget.configure(**props)
            except Exception:
                pass

        for cartao in getattr(self, "_cartoes_tematizaveis", []):
            cartao.aplicar_tema(paleta)

        if hasattr(self, "texto_log"):
            self.texto_log.configure(fg_color=paleta["log_fundo"])
            caixa = self.texto_log._textbox
            caixa.tag_config("hora", foreground=paleta["log_hora"])
            caixa.tag_config("info", foreground=paleta["log_texto"])
            caixa.tag_config("destaque", foreground=paleta["accent"])

        self.tema_atual.set(nome)
        self._atualizar_estilo_menu()
        # Tabelas usam cor "crua" nos widgets (não passam por _temaveis),
        # então precisam ser reconstruídas quando o tema muda.
        self._atualizar_tabelas_execucoes()

    # =====================================================
    # LAYOUT GERAL: SIDEBAR + ÁREA DE CONTEÚDO
    # =====================================================
    def _montar_layout(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True)

        self._construir_sidebar(container)

        self.area_conteudo = ctk.CTkFrame(container, fg_color="transparent")
        self.area_conteudo.pack(side="left", fill="both", expand=True)

        self._cartoes_tematizaveis = []

        self._pagina_inicio()
        self._pagina_perfis()
        self._pagina_execucoes()
        self._pagina_configuracoes()
        self._pagina_logs()
        self._pagina_suporte()
        self._pagina_sobre()

    def _construir_sidebar(self, container):
        self.sidebar = ctk.CTkFrame(container, corner_radius=0, width=250)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self._registrar_tema(self.sidebar, fg_color="sidebar")

        bloco_logo = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bloco_logo.pack(fill="x", padx=22, pady=(28, 18))

        logo_img = self._carregar_logo((64, 64))
        if logo_img is not None:
            logo_label = ctk.CTkLabel(bloco_logo, image=logo_img, text="")
            logo_label.pack(anchor="w")
            self._logo_ref_sidebar = logo_img

        linha1 = ctk.CTkLabel(bloco_logo, text="Captador de", font=ctk.CTkFont(size=18, weight="bold"))
        linha1.pack(anchor="w", pady=(10, 0))
        self._registrar_tema(linha1, text_color="texto_principal")

        linha2 = ctk.CTkLabel(bloco_logo, text="Dados CRM", font=ctk.CTkFont(size=18, weight="bold"))
        linha2.pack(anchor="w")
        self._registrar_tema(linha2, text_color="accent")

        subtitulo = ctk.CTkLabel(
            bloco_logo, text="Automação para consulta\ne extração de dados",
            font=ctk.CTkFont(size=11), justify="left",
        )
        subtitulo.pack(anchor="w", pady=(6, 0))
        self._registrar_tema(subtitulo, text_color="texto_secundario")

        # ---------- Itens do menu ----------
        bloco_menu = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bloco_menu.pack(fill="x", padx=14, pady=(10, 0))

        for chave, emoji, rotulo in ITENS_MENU:
            botao = ctk.CTkButton(
                bloco_menu, text=f"  {emoji}   {rotulo}", anchor="w",
                height=42, corner_radius=10, font=ctk.CTkFont(size=13),
                command=lambda c=chave: self._mostrar_pagina(c),
            )
            botao.pack(fill="x", pady=3)
            self._botoes_menu[chave] = botao

        # ---------- Rodapé ----------
        rodape = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        rodape.pack(side="bottom", fill="x", padx=22, pady=20)

        linha_rodape1 = ctk.CTkLabel(
            rodape, text=f"© {datetime.now().year} {NOME_PRODUTO}", font=ctk.CTkFont(size=10),
        )
        linha_rodape1.pack(anchor="w")
        self._registrar_tema(linha_rodape1, text_color="texto_secundario")

        linha_rodape2 = ctk.CTkLabel(
            rodape, text="Todos os direitos reservados", font=ctk.CTkFont(size=10),
        )
        linha_rodape2.pack(anchor="w")
        self._registrar_tema(linha_rodape2, text_color="texto_secundario")

    def _atualizar_estilo_menu(self):
        paleta = TEMAS[self.tema_atual.get()]
        for chave, botao in self._botoes_menu.items():
            ativo = (chave == self.pagina_atual)
            if ativo:
                botao.configure(
                    fg_color=paleta["accent"], hover_color=paleta["accent_hover"],
                    text_color=paleta["accent_texto"],
                )
            else:
                botao.configure(
                    fg_color="transparent", hover_color=paleta["secundario"],
                    text_color=paleta["texto_principal"],
                )

    def _mostrar_pagina(self, nome):
        for frame in self._paginas.values():
            frame.pack_forget()
        self._paginas[nome].pack(fill="both", expand=True)
        self.pagina_atual = nome
        self._atualizar_estilo_menu()
        if nome in ("inicio", "execucoes"):
            self._atualizar_tabelas_execucoes()
        if nome == "inicio":
            self._atualizar_cards_topo()

    # =====================================================
    # CABEÇALHO PADRÃO DE PÁGINA
    # =====================================================
    def _cabecalho_pagina(self, pai, titulo, subtitulo):
        header = ctk.CTkFrame(pai, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(26, 16))

        bloco_texto = ctk.CTkFrame(header, fg_color="transparent")
        bloco_texto.pack(side="left", anchor="w")

        titulo_label = ctk.CTkLabel(bloco_texto, text=titulo, font=ctk.CTkFont(size=24, weight="bold"))
        titulo_label.pack(anchor="w")
        self._registrar_tema(titulo_label, text_color="texto_principal")

        subtitulo_label = ctk.CTkLabel(bloco_texto, text=subtitulo, font=ctk.CTkFont(size=13))
        subtitulo_label.pack(anchor="w", pady=(2, 0))
        self._registrar_tema(subtitulo_label, text_color="texto_secundario")

        return header

    def _criar_card(self, pai, **kwargs):
        card = ctk.CTkFrame(pai, corner_radius=16, **kwargs)
        self._registrar_tema(card, fg_color="card")
        return card

    # =====================================================
    # PÁGINA: INÍCIO (dashboard)
    # =====================================================
    def _pagina_inicio(self):
        pagina = ctk.CTkScrollableFrame(self.area_conteudo, fg_color="transparent")
        self._registrar_tema(pagina, fg_color="fundo")
        self._paginas["inicio"] = pagina

        header = self._cabecalho_pagina(pagina, "Bem-vindo!", "Gerencie as execuções e acompanhe o progresso da captura de dados.")

        self.botao_nova_execucao = ctk.CTkButton(
            header, text="➕  Nova Execução", height=40, corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._mostrar_pagina("execucoes"),
        )
        self.botao_nova_execucao.pack(side="right", anchor="e", padx=(0, 4))
        self._registrar_tema(self.botao_nova_execucao, fg_color="accent", hover_color="accent_hover", text_color="accent_texto")

        self.botao_parar_inicio = ctk.CTkButton(
            header, text="⏹  Parar", height=40, corner_radius=10, width=100,
            fg_color=COR_PARAR, hover_color=COR_PARAR_HOVER, state="disabled",
            command=self._parar,
        )
        self.botao_parar_inicio.pack(side="right", anchor="e", padx=(0, 10))

        # ---------- Cards do topo ----------
        linha_cards = ctk.CTkFrame(pagina, fg_color="transparent")
        linha_cards.pack(fill="x", padx=30, pady=(0, 18))
        linha_cards.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.card_dados_capturados = CartaoIcone(linha_cards, "📄", "Dados Capturados (histórico total)")
        self.card_dados_capturados.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        self.card_perfis_configurados = CartaoIcone(linha_cards, "👤", "Perfis Selecionados")
        self.card_perfis_configurados.grid(row=0, column=1, sticky="nsew", padx=8)

        self.card_execucoes_hoje = CartaoIcone(linha_cards, "🔄", "Execuções Hoje")
        self.card_execucoes_hoje.grid(row=0, column=2, sticky="nsew", padx=8)

        self.card_tempo_hoje = CartaoIcone(linha_cards, "⏱", "Tempo Total Hoje")
        self.card_tempo_hoje.grid(row=0, column=3, sticky="nsew", padx=(8, 0))

        self._cartoes_tematizaveis.extend([
            self.card_dados_capturados, self.card_perfis_configurados,
            self.card_execucoes_hoje, self.card_tempo_hoje,
        ])

        # ---------- Execução em andamento ----------
        self.card_execucao = self._criar_card(pagina)
        self.card_execucao.pack(fill="x", padx=30, pady=(0, 18))

        faixa = ctk.CTkFrame(self.card_execucao, width=4, corner_radius=0)
        faixa.pack(side="left", fill="y")
        self._registrar_tema(faixa, fg_color="accent")

        corpo_execucao = ctk.CTkFrame(self.card_execucao, fg_color="transparent")
        corpo_execucao.pack(side="left", fill="both", expand=True, padx=20, pady=18)

        titulo_execucao = ctk.CTkLabel(corpo_execucao, text="Execução em andamento", font=ctk.CTkFont(size=15, weight="bold"))
        titulo_execucao.pack(anchor="w")
        self._registrar_tema(titulo_execucao, text_color="accent")

        # -- placeholder (sem execução rodando) --
        self.frame_execucao_placeholder = ctk.CTkFrame(corpo_execucao, fg_color="transparent")
        self.label_execucao_placeholder = ctk.CTkLabel(
            self.frame_execucao_placeholder,
            text="Nenhuma execução em andamento no momento. Clique em \"Nova Execução\" para começar.",
            font=ctk.CTkFont(size=12),
        )
        self.label_execucao_placeholder.pack(anchor="w", pady=(10, 4))
        self._registrar_tema(self.label_execucao_placeholder, text_color="texto_secundario")

        # -- detalhe (execução rodando) --
        self.frame_execucao_detalhe = ctk.CTkFrame(corpo_execucao, fg_color="transparent")

        linha_info1 = ctk.CTkFrame(self.frame_execucao_detalhe, fg_color="transparent")
        linha_info1.pack(fill="x", pady=(12, 2))

        self.label_perfil_execucao = ctk.CTkLabel(linha_info1, text="Perfil: —", font=ctk.CTkFont(size=12))
        self.label_perfil_execucao.pack(side="left")
        self._registrar_tema(self.label_perfil_execucao, text_color="texto_principal")

        self.label_inicio_execucao = ctk.CTkLabel(linha_info1, text="Início: —", font=ctk.CTkFont(size=12))
        self.label_inicio_execucao.pack(side="right")
        self._registrar_tema(self.label_inicio_execucao, text_color="texto_principal")

        linha_info2 = ctk.CTkFrame(self.frame_execucao_detalhe, fg_color="transparent")
        linha_info2.pack(fill="x", pady=(2, 10))

        self.label_status_execucao = ctk.CTkLabel(linha_info2, text="Status: —", font=ctk.CTkFont(size=12, weight="bold"))
        self.label_status_execucao.pack(side="left")
        self._registrar_tema(self.label_status_execucao, text_color="accent")

        self.label_previsao_execucao = ctk.CTkLabel(linha_info2, text="Previsão de término: —", font=ctk.CTkFont(size=12))
        self.label_previsao_execucao.pack(side="right")
        self._registrar_tema(self.label_previsao_execucao, text_color="texto_principal")

        label_progresso_geral = ctk.CTkLabel(self.frame_execucao_detalhe, text="Progresso geral", font=ctk.CTkFont(size=12))
        label_progresso_geral.pack(anchor="w")
        self._registrar_tema(label_progresso_geral, text_color="texto_secundario")

        linha_barra = ctk.CTkFrame(self.frame_execucao_detalhe, fg_color="transparent")
        linha_barra.pack(fill="x", pady=(4, 16))

        self.barra_progresso = ctk.CTkProgressBar(linha_barra, height=14, corner_radius=8)
        self.barra_progresso.pack(side="left", fill="x", expand=True)
        self.barra_progresso.set(0)
        self._registrar_tema(self.barra_progresso, progress_color="accent", fg_color="secundario")

        self.label_percentual = ctk.CTkLabel(linha_barra, text="0%", font=ctk.CTkFont(size=12, weight="bold"), width=44)
        self.label_percentual.pack(side="left", padx=(10, 0))
        self._registrar_tema(self.label_percentual, text_color="texto_principal")

        linha_subcards = ctk.CTkFrame(self.frame_execucao_detalhe, fg_color="transparent")
        linha_subcards.pack(fill="x")
        linha_subcards.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.cartao_processados = CartaoStat(linha_subcards, "Processados", COR_SUCESSO, "📄")
        self.cartao_processados.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.cartao_sucesso = CartaoStat(linha_subcards, "Sucessos", COR_SUCESSO, "✅")
        self.cartao_sucesso.grid(row=0, column=1, sticky="nsew", padx=6)
        self.cartao_pendentes = CartaoStat(linha_subcards, "Pendentes", COR_ALERTA, "⏳")
        self.cartao_pendentes.grid(row=0, column=2, sticky="nsew", padx=6)
        self.cartao_erros = CartaoStat(linha_subcards, "Erros", COR_ERRO, "❌")
        self.cartao_erros.grid(row=0, column=3, sticky="nsew", padx=(6, 0))

        self._cartoes_tematizaveis.extend([
            self.cartao_processados, self.cartao_sucesso, self.cartao_pendentes, self.cartao_erros,
        ])

        self.frame_execucao_placeholder.pack(fill="x")

        # ---------- Execuções recentes ----------
        card_recentes = self._criar_card(pagina)
        card_recentes.pack(fill="x", padx=30, pady=(0, 24))

        cabecalho_recentes = ctk.CTkFrame(card_recentes, fg_color="transparent")
        cabecalho_recentes.pack(fill="x", padx=20, pady=(18, 6))

        titulo_recentes = ctk.CTkLabel(cabecalho_recentes, text="Execuções recentes", font=ctk.CTkFont(size=15, weight="bold"))
        titulo_recentes.pack(side="left")
        self._registrar_tema(titulo_recentes, text_color="accent")

        self.frame_tabela_recentes = ctk.CTkFrame(card_recentes, fg_color="transparent")
        self.frame_tabela_recentes.pack(fill="x", padx=20, pady=(0, 8))

        botao_ver_todas = ctk.CTkButton(
            card_recentes, text="📊  Ver todas as execuções", height=34, corner_radius=8,
            font=ctk.CTkFont(size=12), command=lambda: self._mostrar_pagina("execucoes"),
        )
        botao_ver_todas.pack(padx=20, pady=(0, 18))
        self._registrar_tema(botao_ver_todas, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")

    # =====================================================
    # PÁGINA: PERFIS (identidade e status de cada perfil)
    # =====================================================
    def _pagina_perfis(self):
        pagina = ctk.CTkScrollableFrame(self.area_conteudo, fg_color="transparent")
        self._registrar_tema(pagina, fg_color="fundo")
        self._paginas["perfis"] = pagina

        self._cabecalho_pagina(
            pagina, "Perfis",
            "Cada perfil guarda um login independente do Chrome. O nome do perfil é preenchido automaticamente "
            "com o e-mail detectado no login — use os lápis (✎) pra dar um apelido personalizado ou corrigir o "
            "e-mail manualmente, caso a detecção automática não funcione.",
        )

        linha_cards = ctk.CTkFrame(pagina, fg_color="transparent")
        linha_cards.pack(fill="x", padx=30, pady=(0, 24))
        for coluna in range(3):
            linha_cards.grid_columnconfigure(coluna, weight=1, uniform="perfil")

        self._cartoes_identidade_perfil = {}
        self._labels_titulo_perfil = {}
        self._labels_badge_perfil = {}
        self._labels_email_perfil = {}
        self._botoes_limpar_perfil = {}
        self._botoes_verificar_perfil = {}
        self._botoes_login_manual_perfil = {}

        for indice, numero in enumerate((1, 2, 3)):
            card = self._criar_card(linha_cards)
            card.grid(row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0))
            self._cartoes_identidade_perfil[numero] = card

            cabecalho = ctk.CTkFrame(card, fg_color="transparent")
            cabecalho.pack(fill="x", padx=18, pady=(18, 2))

            titulo = ctk.CTkLabel(cabecalho, text=f"Perfil {numero}", font=ctk.CTkFont(size=15, weight="bold"))
            titulo.pack(side="left")
            self._registrar_tema(titulo, text_color="texto_principal")
            self._labels_titulo_perfil[numero] = titulo

            botao_editar_apelido = ctk.CTkButton(
                cabecalho, text="✎", width=24, height=24, corner_radius=6,
                font=ctk.CTkFont(size=12), fg_color="transparent",
                command=lambda n=numero: self._editar_apelido_perfil(n),
            )
            botao_editar_apelido.pack(side="left", padx=(6, 0))
            self._registrar_tema(botao_editar_apelido, text_color="texto_secundario", hover_color="secundario")

            badge = ctk.CTkLabel(
                cabecalho, text="sem login salvo", font=ctk.CTkFont(size=11, weight="bold"),
                corner_radius=8, width=0, height=22, padx=8,
            )
            badge.pack(side="right")
            self._labels_badge_perfil[numero] = badge

            linha_email = ctk.CTkFrame(card, fg_color="transparent")
            linha_email.pack(fill="x", padx=18, pady=(10, 2))

            label_email = ctk.CTkLabel(
                linha_email, text="E-mail: ainda não detectado", font=ctk.CTkFont(size=12),
                anchor="w", justify="left", wraplength=280,
            )
            label_email.pack(side="left", fill="x", expand=True)
            self._registrar_tema(label_email, text_color="texto_secundario")
            self._labels_email_perfil[numero] = label_email

            botao_editar_email = ctk.CTkButton(
                linha_email, text="✎", width=24, height=24, corner_radius=6,
                font=ctk.CTkFont(size=12), fg_color="transparent",
                command=lambda n=numero: self._editar_email_perfil(n),
            )
            botao_editar_email.pack(side="right")
            self._registrar_tema(botao_editar_email, text_color="texto_secundario", hover_color="secundario")

            label_pasta = ctk.CTkLabel(
                card, text=f"Pasta local: perfil_{numero}", font=ctk.CTkFont(size=11), anchor="w",
            )
            label_pasta.pack(fill="x", padx=18, pady=(0, 14))
            self._registrar_tema(label_pasta, text_color="texto_secundario")

            linha_botoes = ctk.CTkFrame(card, fg_color="transparent")
            linha_botoes.pack(fill="x", padx=18, pady=(0, 18))

            botao_verificar = ctk.CTkButton(
                linha_botoes, text="Verificar login", height=32, corner_radius=8,
                font=ctk.CTkFont(size=11),
                command=lambda n=numero: self._verificar_login(n),
            )
            botao_verificar.pack(side="left", fill="x", expand=True, padx=(0, 6))
            self._registrar_tema(botao_verificar, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")
            self._botoes_verificar_perfil[numero] = botao_verificar

            botao_login_manual = ctk.CTkButton(
                linha_botoes, text="Login manual", height=32, corner_radius=8,
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda n=numero: self._login_manual(n),
            )
            botao_login_manual.pack(side="left", fill="x", expand=True, padx=(0, 6))
            self._registrar_tema(botao_login_manual, fg_color="accent", hover_color="accent_hover", text_color="accent_texto")
            self._botoes_login_manual_perfil[numero] = botao_login_manual

            botao_limpar = ctk.CTkButton(
                linha_botoes, text="Limpar perfil", height=32, corner_radius=8,
                font=ctk.CTkFont(size=11), fg_color=COR_PARAR, hover_color=COR_PARAR_HOVER,
                command=lambda n=numero: self._limpar_perfil(n),
            )
            botao_limpar.pack(side="left", fill="x", expand=True)
            self._botoes_limpar_perfil[numero] = botao_limpar

    def _carregar_perfis_na_tela(self):
        """Lê o arquivo de configuração de perfis e reflete apelido/e-mail/status na tela."""
        config = carregar_perfis_config()
        for numero in (1, 2, 3):
            dados = config.get(str(numero), {})
            self.perfil_apelido_var[numero].set(dados.get("apelido", ""))
            self._atualizar_status_perfil_label(numero, dados)

    def _atualizar_status_perfil_label(self, numero, dados=None):
        if dados is None:
            dados = carregar_perfis_config().get(str(numero), {})

        apelido = (dados.get("apelido") or "").strip()
        email = (dados.get("email") or dados.get("login_detectado") or "").strip()
        status = dados.get("status", "Desconhecido")

        titulo_label = self._labels_titulo_perfil.get(numero)
        if titulo_label is not None:
            titulo_label.configure(text=apelido or f"Perfil {numero}")

        email_label = self._labels_email_perfil.get(numero)
        if email_label is not None:
            email_label.configure(text=f"E-mail: {email}" if email else "E-mail: ainda não detectado")

        badge = self._labels_badge_perfil.get(numero)
        if badge is not None:
            paleta = TEMAS[self.tema_atual.get()]
            if status == "Logado":
                badge.configure(text="Logado", text_color=COR_SUCESSO)
            else:
                badge.configure(text="sem login salvo", text_color=COR_ALERTA)
            badge.configure(fg_color=paleta["secundario"])

    def _editar_apelido_perfil(self, numero):
        atual = self.perfil_apelido_var[numero].get().strip()
        dialogo = ctk.CTkInputDialog(
            text=f"Novo apelido para o Perfil {numero} (deixe em branco para usar o e-mail detectado):",
            title="Editar apelido",
        )
        novo = dialogo.get_input()
        if novo is None:
            return
        novo = novo.strip()
        definir_apelido_perfil(numero, novo)
        self.perfil_apelido_var[numero].set(novo)
        self._atualizar_status_perfil_label(numero)
        self._atualizar_cards_topo()
        if hasattr(self, "_checkboxes_perfil_execucao"):
            self._atualizar_rotulos_checkbox_perfil()

    def _editar_email_perfil(self, numero):
        dialogo = ctk.CTkInputDialog(
            text=f"E-mail do Perfil {numero} (edite manualmente se a detecção automática não funcionou):",
            title="Editar e-mail",
        )
        novo = dialogo.get_input()
        if novo is None:
            return
        definir_email_manual(numero, novo.strip())
        self._atualizar_status_perfil_label(numero)

    def _verificar_login(self, numero):
        """Abre o navegador (headless) naquele perfil e roda a detecção de status/e-mail sob demanda."""
        botao = self._botoes_verificar_perfil.get(numero)
        if botao is not None:
            botao.configure(state="disabled", text="Verificando...")

        def trabalho():
            pasta_perfil = os.path.join(PASTA_PERFIS, f"perfil_{numero}")
            driver = None
            try:
                driver = iniciar_chrome(pasta_perfil, headless=True)
                esperar_crm(driver)
                status = detectar_status_login(driver)
                definir_status_perfil(numero, status)
                if status == "Logado":
                    email = capturar_email_conta(driver)
                    if email:
                        registrar_email_detectado(numero, email)
            except Exception as erro:
                self.fila.put(("erro_verificar_perfil", (numero, str(erro))))
            finally:
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:
                        pass
                self.fila.put(("fim_verificar_perfil", numero))

        threading.Thread(target=trabalho, daemon=True).start()

    def _login_manual(self, numero):
        """
        Abre uma janela visível do Chrome naquele perfil, direto na tela de
        login do CRM, pro usuário logar à mão. Em paralelo, monitora o
        campo de e-mail da tela de login da Microsoft pra capturar em
        tempo real, sem precisar esperar o usuário chegar no painel de
        conta depois.
        """
        pasta_perfil = os.path.join(PASTA_PERFIS, f"perfil_{numero}")

        def monitorar(driver):
            email_ja_capturado = False
            fim = time.time() + 600  # até 10 minutos de janela aberta
            while time.time() < fim:
                try:
                    _ = driver.title  # dispara exceção se a janela foi fechada
                except Exception:
                    break

                if not email_ja_capturado:
                    try:
                        campo = driver.find_element(
                            By.CSS_SELECTOR,
                            "input[type='email'], input[name='loginfmt'], #i0116",
                        )
                        valor = (campo.get_attribute("value") or "").strip()
                        if valor and REGEX_EMAIL.match(valor):
                            registrar_email_detectado(numero, valor)
                            self.fila.put(("email_detectado_perfil", numero))
                            email_ja_capturado = True
                    except Exception:
                        pass

                try:
                    status = detectar_status_login(driver, tempo_espera=0)
                    if status == "Logado":
                        definir_status_perfil(numero, status)
                        if not email_ja_capturado:
                            email = capturar_email_conta(driver)
                            if email:
                                registrar_email_detectado(numero, email)
                                email_ja_capturado = True
                        self.fila.put(("email_detectado_perfil", numero))
                        break
                except Exception:
                    pass

                time.sleep(1.5)

            self.fila.put(("fim_login_manual_perfil", numero))

        def trabalho():
            try:
                driver = iniciar_chrome(pasta_perfil, headless=False)
                esperar_crm(driver)
            except Exception as erro:
                self.fila.put(("erro_verificar_perfil", (numero, str(erro))))
                return
            monitorar(driver)

        threading.Thread(target=trabalho, daemon=True).start()

    def _limpar_perfil(self, numero):
        if self.threads_ativas > 0:
            messagebox.showwarning("Execução em andamento", "Espere a execução atual terminar antes de limpar um perfil.")
            return

        apelido_atual = self.perfil_apelido_var[numero].get().strip() or f"Perfil {numero}"
        confirmar = messagebox.askyesno(
            "Limpar perfil",
            f"Isso vai apagar a sessão de login salva do \"{apelido_atual}\" (Perfil {numero}).\n\n"
            "Na próxima execução usando esse perfil, será necessário fazer login novamente na janela do Chrome.\n\n"
            "Os outros perfis não são afetados. Deseja continuar?",
        )
        if not confirmar:
            return

        limpar_perfil_individual(numero)
        self.perfil_apelido_var[numero].set("")
        self._atualizar_status_perfil_label(numero)
        self._atualizar_cards_topo()
        if hasattr(self, "_checkboxes_perfil_execucao"):
            self._atualizar_rotulos_checkbox_perfil()
        messagebox.showinfo("Perfil limpo", f"O Perfil {numero} foi resetado. Será pedido login novamente na próxima execução.")

    # =====================================================
    # PÁGINA: EXECUÇÕES (importar base, escolher perfis, histórico)
    # =====================================================
    def _pagina_execucoes(self):
        pagina = ctk.CTkScrollableFrame(self.area_conteudo, fg_color="transparent")
        self._registrar_tema(pagina, fg_color="fundo")
        self._paginas["execucoes"] = pagina

        self._cabecalho_pagina(pagina, "Execuções", "Importe a base de CPF's, escolha os perfis e inicie a captura.")

        # ---------- Nova execução ----------
        card_nova_execucao = self._criar_card(pagina)
        card_nova_execucao.pack(fill="x", padx=30, pady=(0, 20))

        titulo_nova_execucao = ctk.CTkLabel(card_nova_execucao, text="Nova execução", font=ctk.CTkFont(size=16, weight="bold"))
        titulo_nova_execucao.pack(anchor="w", padx=20, pady=(18, 14))
        self._registrar_tema(titulo_nova_execucao, text_color="texto_principal")

        linha_arquivo = ctk.CTkFrame(card_nova_execucao, fg_color="transparent")
        linha_arquivo.pack(fill="x", padx=20, pady=(0, 6))

        rotulo_base = ctk.CTkLabel(linha_arquivo, text="Base com CPF's:", font=ctk.CTkFont(size=13, weight="bold"))
        rotulo_base.pack(side="left", padx=(0, 14))
        self._registrar_tema(rotulo_base, text_color="texto_secundario")

        nome_arquivo_atual = os.path.basename(self.caminho_cpfs.get().strip()) or "nenhuma base importada"
        self.label_arquivo_selecionado = ctk.CTkLabel(linha_arquivo, text=nome_arquivo_atual, font=ctk.CTkFont(size=13))
        self.label_arquivo_selecionado.pack(side="left", fill="x", expand=True, anchor="w")
        self._registrar_tema(self.label_arquivo_selecionado, text_color="texto_principal")

        botao_selecionar = ctk.CTkButton(
            linha_arquivo, text="Importar base...", width=150, height=34, corner_radius=8,
            command=self._selecionar_arquivo,
        )
        botao_selecionar.pack(side="right")
        self._registrar_tema(botao_selecionar, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")

        nota_importar = ctk.CTkLabel(
            card_nova_execucao,
            text=(
                "Não é preciso separar a planilha por perfil: a automação divide a base de CPF's automaticamente "
                "entre os perfis marcados abaixo e busca cada CPF direto no CRM, capturando os dados de inscrição "
                "disponíveis."
            ),
            font=ctk.CTkFont(size=11), justify="left", wraplength=900,
        )
        nota_importar.pack(anchor="w", padx=20, pady=(4, 20))
        self._registrar_tema(nota_importar, text_color="texto_secundario")

        # ---------- Quais perfis rodam essa base ----------
        linha_perfis = ctk.CTkFrame(card_nova_execucao, fg_color="transparent")
        linha_perfis.pack(fill="x", padx=20, pady=(0, 20))
        linha_perfis.grid_columnconfigure(0, weight=0)
        linha_perfis.grid_columnconfigure(1, weight=1)

        coluna_checkboxes = ctk.CTkFrame(linha_perfis, fg_color="transparent")
        coluna_checkboxes.grid(row=0, column=0, sticky="nw")

        rotulo_perfis_execucao = ctk.CTkLabel(coluna_checkboxes, text="Perfis:", font=ctk.CTkFont(size=13, weight="bold"))
        rotulo_perfis_execucao.pack(anchor="w", pady=(0, 10))
        self._registrar_tema(rotulo_perfis_execucao, text_color="texto_principal")

        self._checkboxes_perfil_execucao = {}
        for numero in (1, 2, 3):
            checkbox = ctk.CTkCheckBox(
                coluna_checkboxes, text=f"Perfil {numero}", variable=self.perfil_selecionado[numero],
                font=ctk.CTkFont(size=13),
                command=self._atualizar_cards_topo,
            )
            checkbox.pack(anchor="w", pady=(0, 10))
            self._registrar_tema(checkbox, text_color="texto_principal", fg_color="accent", hover_color="accent_hover")
            self._checkboxes_perfil_execucao[numero] = checkbox

        nota_checkboxes = ctk.CTkLabel(
            linha_perfis,
            text="Marque um ou mais — cada um vira uma janela\ntrabalhando em paralelo.",
            font=ctk.CTkFont(size=11), justify="right", anchor="ne",
        )
        nota_checkboxes.grid(row=0, column=1, sticky="ne", padx=(20, 0))
        self._registrar_tema(nota_checkboxes, text_color="texto_secundario")

        self._atualizar_rotulos_checkbox_perfil()

        # ---------- Botões de ação ----------
        linha_botoes_exec = ctk.CTkFrame(card_nova_execucao, fg_color="transparent")
        linha_botoes_exec.pack(fill="x", padx=20, pady=(0, 20))

        self.botao_importar_exec = ctk.CTkButton(
            linha_botoes_exec, text="⬇  Importar", height=40, width=140, corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"), command=self._iniciar,
        )
        self.botao_importar_exec.pack(side="left", padx=(0, 10))
        self._registrar_tema(self.botao_importar_exec, fg_color="accent", hover_color="accent_hover", text_color="accent_texto")

        self.botao_parar_exec = ctk.CTkButton(
            linha_botoes_exec, text="⏹  Parar", height=40, width=120, corner_radius=10,
            fg_color=COR_PARAR, hover_color=COR_PARAR_HOVER, state="disabled",
            command=self._parar,
        )
        self.botao_parar_exec.pack(side="left")

        # ---------- Histórico de execuções ----------
        rotulo_historico = ctk.CTkLabel(pagina, text="Histórico de execuções", font=ctk.CTkFont(size=15, weight="bold"))
        rotulo_historico.pack(anchor="w", padx=30, pady=(0, 8))
        self._registrar_tema(rotulo_historico, text_color="texto_principal")

        card_tabela = self._criar_card(pagina)
        card_tabela.pack(fill="both", expand=True, padx=30, pady=(0, 24))

        self.frame_tabela_completa = ctk.CTkFrame(card_tabela, fg_color="transparent")
        self.frame_tabela_completa.pack(fill="both", expand=True, padx=20, pady=18)

    def _atualizar_rotulos_checkbox_perfil(self):
        """Atualiza o texto de cada checkbox pra mostrar o apelido salvo do perfil."""
        for numero, checkbox in self._checkboxes_perfil_execucao.items():
            apelido = self.perfil_apelido_var[numero].get().strip()
            texto = f"Perfil {numero} — {apelido}" if apelido else f"Perfil {numero}"
            checkbox.configure(text=texto)

    # =====================================================
    # PÁGINA: CONFIGURAÇÕES
    # =====================================================
    def _pagina_configuracoes(self):
        pagina = ctk.CTkScrollableFrame(self.area_conteudo, fg_color="transparent")
        self._registrar_tema(pagina, fg_color="fundo")
        self._paginas["configuracoes"] = pagina

        self._cabecalho_pagina(pagina, "Configurações", "Ajustes de aparência e atalhos para as pastas do programa.")

        card_tema = self._criar_card(pagina)
        card_tema.pack(fill="x", padx=30, pady=(0, 16))

        rotulo_tema = ctk.CTkLabel(card_tema, text="🎨 Tema", font=ctk.CTkFont(size=15, weight="bold"))
        rotulo_tema.pack(anchor="w", padx=20, pady=(18, 10))
        self._registrar_tema(rotulo_tema, text_color="texto_principal")

        self.seletor_tema = ctk.CTkSegmentedButton(
            card_tema, values=list(TEMAS.keys()), command=self._aplicar_tema, variable=self.tema_atual,
        )
        self.seletor_tema.pack(anchor="w", padx=20, pady=(0, 20))

        card_pastas = self._criar_card(pagina)
        card_pastas.pack(fill="x", padx=30, pady=(0, 16))

        rotulo_pastas = ctk.CTkLabel(card_pastas, text="📁 Pastas e arquivos", font=ctk.CTkFont(size=15, weight="bold"))
        rotulo_pastas.pack(anchor="w", padx=20, pady=(18, 10))
        self._registrar_tema(rotulo_pastas, text_color="texto_principal")

        linha_botoes = ctk.CTkFrame(card_pastas, fg_color="transparent")
        linha_botoes.pack(fill="x", padx=20, pady=(0, 20))

        self.botao_saida = ctk.CTkButton(linha_botoes, text="📁 Abrir pasta Saída", height=40, corner_radius=10,
                                          command=lambda: self._abrir_pasta(PASTA_SAIDA))
        self.botao_saida.pack(side="left", padx=(0, 10))
        self._registrar_tema(self.botao_saida, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")

        self.botao_logs_pasta = ctk.CTkButton(linha_botoes, text="📁 Abrir pasta Logs", height=40, corner_radius=10,
                                               command=lambda: self._abrir_pasta(PASTA_LOGS))
        self.botao_logs_pasta.pack(side="left", padx=(0, 10))
        self._registrar_tema(self.botao_logs_pasta, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")

        self.botao_limpar_imagens = ctk.CTkButton(linha_botoes, text="🧹 Limpar imagens de log", height=40, corner_radius=10,
                                                    command=self._limpar_imagens)
        self.botao_limpar_imagens.pack(side="left")
        self._registrar_tema(self.botao_limpar_imagens, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")

        card_estatisticas = self._criar_card(pagina)
        card_estatisticas.pack(fill="x", padx=30, pady=(0, 16))

        rotulo_estatisticas = ctk.CTkLabel(card_estatisticas, text="📊 Estatísticas do painel", font=ctk.CTkFont(size=15, weight="bold"))
        rotulo_estatisticas.pack(anchor="w", padx=20, pady=(18, 6))
        self._registrar_tema(rotulo_estatisticas, text_color="texto_principal")

        nota_estatisticas = ctk.CTkLabel(
            card_estatisticas,
            text=(
                "\"Dados Capturados\" soma todo o histórico de sucessos já registrado em "
                "Logs\\log_processamento.xlsx (desde a primeira vez que o programa rodou nesta máquina). "
                "Já a tabela de Execuções só mostra o que rodou depois que esse histórico passou a existir — "
                "por isso os números podem não bater exatamente. Isso é esperado, não é erro de contagem.\n\n"
                "Se quiser começar a contagem do zero, o botão abaixo apaga apenas os arquivos de log e "
                "histórico do painel — os resultados já capturados na pasta Saída não são afetados."
            ),
            font=ctk.CTkFont(size=11), justify="left", wraplength=820,
        )
        nota_estatisticas.pack(anchor="w", padx=20, pady=(0, 14))
        self._registrar_tema(nota_estatisticas, text_color="texto_secundario")

        self.botao_zerar_estatisticas = ctk.CTkButton(
            card_estatisticas, text="♻️ Zerar estatísticas do painel", height=40, corner_radius=10,
            command=self._zerar_estatisticas,
        )
        self.botao_zerar_estatisticas.pack(anchor="w", padx=20, pady=(0, 20))
        self._registrar_tema(self.botao_zerar_estatisticas, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")

    # =====================================================
    # PÁGINA: LOGS
    # =====================================================
    def _pagina_logs(self):
        pagina = ctk.CTkFrame(self.area_conteudo, fg_color="transparent")
        self._registrar_tema(pagina, fg_color="fundo")
        self._paginas["logs"] = pagina

        header = self._cabecalho_pagina(pagina, "Logs", "Acompanhe em tempo real o que a automação está fazendo.")

        botao_limpar_log = ctk.CTkButton(header, text="🗑️ Limpar log", width=120, height=34, corner_radius=8,
                                          font=ctk.CTkFont(size=12), command=self._limpar_log)
        botao_limpar_log.pack(side="right", anchor="e")
        self._registrar_tema(botao_limpar_log, fg_color="secundario", hover_color="secundario_hover", text_color="texto_principal")

        card_log = self._criar_card(pagina)
        card_log.pack(fill="both", expand=True, padx=30, pady=(0, 24))

        self.texto_log = ctk.CTkTextbox(
            card_log, corner_radius=10, height=320,
            font=ctk.CTkFont(family="Consolas", size=13), wrap="word", spacing1=3, spacing3=7,
        )
        self.texto_log.pack(fill="both", expand=True, padx=18, pady=18)
        self.texto_log.configure(state="disabled")
        self._configurar_tags_log()

    # =====================================================
    # PÁGINA: SUPORTE
    # =====================================================
    def _pagina_suporte(self):
        pagina = ctk.CTkScrollableFrame(self.area_conteudo, fg_color="transparent")
        self._registrar_tema(pagina, fg_color="fundo")
        self._paginas["suporte"] = pagina

        self._cabecalho_pagina(pagina, "Suporte", "Dúvidas ou problemas com a automação? Comece por aqui.")

        card = self._criar_card(pagina)
        card.pack(fill="x", padx=30, pady=(0, 16))

        texto = (
            "• Verifique primeiro a aba Logs — a maioria dos erros aparece ali com uma mensagem clara.\n\n"
            "• Screenshots de erro ficam salvos em Logs\\Screenshots, um por CPF que falhou.\n\n"
            "• Se um perfil ficar pedindo login de novo, verifique se a sessão do Google expirou.\n\n"
            "• Para reportar um problema, envie o arquivo de log (Logs\\log_processamento.xlsx) e uma "
            "descrição do que aconteceu para o contato abaixo."
        )
        label = ctk.CTkLabel(card, text=texto, font=ctk.CTkFont(size=12), justify="left", wraplength=820)
        label.pack(anchor="w", padx=20, pady=20)
        self._registrar_tema(label, text_color="texto_principal")

        # ---------- Card de contato ----------
        card_contato = self._criar_card(pagina)
        card_contato.pack(fill="x", padx=30, pady=(0, 16))

        rotulo_telefone_titulo = ctk.CTkLabel(card_contato, text="📞 TELEFONE", font=ctk.CTkFont(size=11, weight="bold"))
        rotulo_telefone_titulo.pack(anchor="w", padx=20, pady=(18, 0))
        self._registrar_tema(rotulo_telefone_titulo, text_color="texto_secundario")

        rotulo_telefone = ctk.CTkLabel(card_contato, text="(11) 94727-8128", font=ctk.CTkFont(size=20, weight="bold"))
        rotulo_telefone.pack(anchor="w", padx=20, pady=(0, 14))
        self._registrar_tema(rotulo_telefone, text_color="texto_principal")

        rotulo_email_titulo = ctk.CTkLabel(card_contato, text="✉️ E-MAIL", font=ctk.CTkFont(size=11, weight="bold"))
        rotulo_email_titulo.pack(anchor="w", padx=20, pady=(0, 0))
        self._registrar_tema(rotulo_email_titulo, text_color="texto_secundario")

        rotulo_email = ctk.CTkLabel(card_contato, text="samueldayvid5@icloud.com", font=ctk.CTkFont(size=20, weight="bold"))
        rotulo_email.pack(anchor="w", padx=20, pady=(0, 20))
        self._registrar_tema(rotulo_email, text_color="texto_principal")

        # ---------- Manual em anexo ----------
        card_manual = self._criar_card(pagina)
        card_manual.pack(fill="x", padx=30, pady=(0, 24))

        rotulo_manual = ctk.CTkLabel(card_manual, text="📘 Manual de uso", font=ctk.CTkFont(size=15, weight="bold"))
        rotulo_manual.pack(anchor="w", padx=20, pady=(18, 6))
        self._registrar_tema(rotulo_manual, text_color="texto_principal")

        nota_manual = ctk.CTkLabel(
            card_manual,
            text="O manual completo, com o passo a passo de cada tela, já vem junto com o programa.",
            font=ctk.CTkFont(size=12),
        )
        nota_manual.pack(anchor="w", padx=20, pady=(0, 10))
        self._registrar_tema(nota_manual, text_color="texto_secundario")

        botao_manual = ctk.CTkButton(
            card_manual, text="📄 Abrir manual de uso", height=40, corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"), command=self._abrir_manual,
        )
        botao_manual.pack(anchor="w", padx=20, pady=(0, 20))
        self._registrar_tema(botao_manual, fg_color="accent", hover_color="accent_hover", text_color="accent_texto")

    # =====================================================
    # PÁGINA: SOBRE
    # =====================================================
    def _pagina_sobre(self):
        pagina = ctk.CTkScrollableFrame(self.area_conteudo, fg_color="transparent")
        self._registrar_tema(pagina, fg_color="fundo")
        self._paginas["sobre"] = pagina

        self._cabecalho_pagina(pagina, "Sobre", "Informações do programa.")

        card = self._criar_card(pagina)
        card.pack(fill="x", padx=30, pady=(0, 24))

        titulo = ctk.CTkLabel(card, text=NOME_PRODUTO, font=ctk.CTkFont(size=18, weight="bold"))
        titulo.pack(anchor="w", padx=20, pady=(20, 4))
        self._registrar_tema(titulo, text_color="texto_principal")

        versao = ctk.CTkLabel(card, text="Versão 2.0.0", font=ctk.CTkFont(size=12))
        versao.pack(anchor="w", padx=20)
        self._registrar_tema(versao, text_color="texto_secundario")

        descricao = ctk.CTkLabel(
            card,
            text=(
                "Automação para consulta e extração de dados de inscrições no CRM Dynamics, "
                "a partir de uma lista de CPFs. Suporta múltiplos perfis do Chrome em paralelo."
            ),
            font=ctk.CTkFont(size=12), justify="left", wraplength=820,
        )
        descricao.pack(anchor="w", padx=20, pady=(10, 20))
        self._registrar_tema(descricao, text_color="texto_principal")

    # =====================================================
    # TABELA DE EXECUÇÕES (usada em Início e Execuções)
    # =====================================================
    def _atualizar_tabelas_execucoes(self):
        self._historico_cache = carregar_historico()
        registros_ordenados = list(reversed(self._historico_cache))
        paleta = TEMAS[self.tema_atual.get()]

        if hasattr(self, "frame_tabela_recentes"):
            self._preencher_tabela(self.frame_tabela_recentes, registros_ordenados[:5], paleta)
        if hasattr(self, "frame_tabela_completa"):
            self._preencher_tabela(self.frame_tabela_completa, registros_ordenados[:50], paleta)

    def _preencher_tabela(self, frame, registros, paleta):
        for widget in frame.winfo_children():
            widget.destroy()

        self._linha_tabela_bruta(frame, ["Data/Hora", "Perfis", "Status", "Registros", "Duração"], paleta, cabecalho=True)

        if not registros:
            vazio = ctk.CTkLabel(frame, text="Nenhuma execução registrada ainda.", font=ctk.CTkFont(size=12))
            vazio.configure(text_color=paleta["texto_secundario"])
            vazio.pack(anchor="w", pady=10)
            return

        for registro in registros:
            status = registro.get("status", "—")
            if status == "Concluído":
                cor_status = COR_SUCESSO
            elif status in ("Erro", "Interrompido"):
                cor_status = COR_ERRO
            else:
                cor_status = paleta["accent"]

            valores = [
                registro.get("inicio", "—"),
                f'{registro.get("perfis", 1)} perfil(is)',
                status,
                str(registro.get("processados", 0)),
                registro.get("duracao") or "em andamento",
            ]
            self._linha_tabela_bruta(frame, valores, paleta, cor_status=cor_status)

    def _linha_tabela_bruta(self, pai, valores, paleta, cabecalho=False, cor_status=None):
        larguras = [155, 100, 110, 90, 100]
        linha = ctk.CTkFrame(pai, fg_color="transparent")
        linha.pack(fill="x", pady=(0, 10) if cabecalho else (0, 6))

        for i, valor in enumerate(valores):
            fonte = ctk.CTkFont(size=11, weight="bold") if cabecalho else ctk.CTkFont(size=12)
            label = ctk.CTkLabel(linha, text=str(valor), font=fonte, width=larguras[i], anchor="w")
            label.pack(side="left")
            if cabecalho:
                label.configure(text_color=paleta["texto_secundario"])
            elif i == 2 and cor_status:
                label.configure(text_color=cor_status, font=ctk.CTkFont(size=12, weight="bold"))
            else:
                label.configure(text_color=paleta["texto_principal"])

        if not cabecalho:
            botao = ctk.CTkButton(
                linha, text="📁", width=32, height=26, corner_radius=6,
                fg_color=paleta["secundario"], hover_color=paleta["secundario_hover"],
                text_color=paleta["texto_principal"],
                command=lambda: self._abrir_pasta(PASTA_SAIDA),
            )
            botao.pack(side="left")

    # =====================================================
    # CARDS DO TOPO / ESTATÍSTICAS GERAIS
    # =====================================================
    def _atualizar_cards_topo(self):
        self.card_dados_capturados.atualizar(total_dados_capturados())
        qtd_selecionados = sum(1 for v in self.perfil_selecionado.values() if v.get())
        self.card_perfis_configurados.atualizar(qtd_selecionados)
        execucoes_hoje, tempo_hoje = calcular_estatisticas_hoje()
        self.card_execucoes_hoje.atualizar(execucoes_hoje)
        self.card_tempo_hoje.atualizar(tempo_hoje)

    # =====================================================
    # EXECUÇÃO EM ANDAMENTO (widgets ao vivo)
    # =====================================================
    def _mostrar_execucao_ativa(self, ativo):
        if ativo:
            self.frame_execucao_placeholder.pack_forget()
            self.frame_execucao_detalhe.pack(fill="x")
        else:
            self.frame_execucao_detalhe.pack_forget()
            self.frame_execucao_placeholder.pack(fill="x")

    @staticmethod
    def _rotulo_perfis_por_numeros(numeros):
        nomes = [rotulo_perfil_configurado(n) for n in numeros]
        if not nomes:
            return "—"
        if len(nomes) == 1:
            return nomes[0]
        return ", ".join(nomes[:-1]) + f" e {nomes[-1]}"

    def _atualizar_widgets_execucao_atual(self):
        if self.execucao_atual is None:
            return

        soma_atual = sum(a for a, t in self._progresso_por_perfil.values())
        soma_total = sum(t for a, t in self._progresso_por_perfil.values())
        fracao = (soma_atual / soma_total) if soma_total else 0

        total_s = sum(v[0] for v in self._stats_por_perfil.values())
        total_se = sum(v[1] for v in self._stats_por_perfil.values())
        total_e = sum(v[2] for v in self._stats_por_perfil.values())
        processados = total_s + total_se + total_e
        pendentes = max(soma_total - processados, 0)

        self.barra_progresso.set(fracao)
        self.label_percentual.configure(text=f"{int(fracao * 100)}%")

        self.cartao_processados.atualizar(processados)
        self.cartao_sucesso.atualizar(total_s)
        self.cartao_pendentes.atualizar(pendentes)
        self.cartao_erros.atualizar(total_e)

        numeros_perfil = self.execucao_atual.get("numeros_perfil", [])
        self.label_perfil_execucao.configure(text=f"Perfil: {self._rotulo_perfis_por_numeros(numeros_perfil)}")
        self.label_inicio_execucao.configure(text=f"Início: {self.execucao_atual['inicio'].strftime('%d/%m/%Y %H:%M:%S')}")

        status_texto = "Interrompendo..." if self.stop_event.is_set() else "Capturando dados..."
        self.label_status_execucao.configure(text=f"Status: {status_texto}")

        decorrido = datetime.now() - self.execucao_atual["inicio"]
        if fracao > 0.02:
            total_estimado = decorrido / fracao
            previsao = self.execucao_atual["inicio"] + total_estimado
            texto_previsao = previsao.strftime("%d/%m/%Y %H:%M:%S")
        else:
            texto_previsao = "Calculando..."
        self.label_previsao_execucao.configure(text=f"Previsão de término: {texto_previsao}")

    def _tick_relogio(self):
        if self.execucao_atual is not None:
            self._atualizar_widgets_execucao_atual()
            self._contador_tick = getattr(self, "_contador_tick", 0) + 1
            if self._contador_tick % 5 == 0:
                execucoes_hoje, tempo_hoje = calcular_estatisticas_hoje()
                self.card_execucoes_hoje.atualizar(execucoes_hoje)
                self.card_tempo_hoje.atualizar(tempo_hoje)

        self.after(1000, self._tick_relogio)

    # =====================================================
    # CONTROLES DE EXECUÇÃO (iniciar / parar)
    # =====================================================
    def _definir_estado_execucao(self, rodando):
        estado_iniciar = "disabled" if rodando else "normal"
        estado_parar = "normal" if rodando else "disabled"

        botao = getattr(self, "botao_importar_exec", None)
        if botao is not None:
            botao.configure(state=estado_iniciar)

        for nome_botao in ("botao_parar_inicio", "botao_parar_exec"):
            botao = getattr(self, nome_botao, None)
            if botao is not None:
                botao.configure(state=estado_parar)

        for checkbox in getattr(self, "_checkboxes_perfil_execucao", {}).values():
            checkbox.configure(state=estado_iniciar)

        for botao in getattr(self, "_botoes_limpar_perfil", {}).values():
            botao.configure(state=estado_iniciar)

        for nome in ("botao_saida", "botao_logs_pasta", "botao_limpar_imagens", "botao_zerar_estatisticas"):
            widget = getattr(self, nome, None)
            if widget is not None:
                widget.configure(state=estado_iniciar)

    @staticmethod
    def _dividir_lista(itens, n_partes):
        n_partes = max(1, min(n_partes, len(itens)) if itens else 1)
        tamanho = -(-len(itens) // n_partes) if n_partes else len(itens)
        return [itens[i:i + tamanho] for i in range(0, len(itens), tamanho)] or [[]]

    def _iniciar(self):
        if self.threads_ativas > 0:
            return

        caminho = self.caminho_cpfs.get().strip()
        if not caminho or not os.path.exists(caminho):
            messagebox.showerror(
                "Arquivo não encontrado",
                f"Não encontrei o arquivo:\n{caminho}\n\nClique em \"Importar base...\" e escolha a planilha de RAs.",
            )
            return

        try:
            cpfs = ler_cpfs(caminho)
        except Exception as erro:
            messagebox.showerror("Erro ao ler CPFs", str(erro))
            return

        if not cpfs:
            messagebox.showinfo("Nada a processar", "A planilha não tem CPFs para processar.")
            return

        perfis_marcados = [n for n in (1, 2, 3) if self.perfil_selecionado[n].get()]
        if not perfis_marcados:
            messagebox.showerror(
                "Nenhum perfil selecionado",
                "Marque pelo menos um perfil (1, 2 ou 3) na seção \"Quais perfis vão rodar essa base\".",
            )
            return

        lotes = self._dividir_lista(cpfs, len(perfis_marcados))
        lotes = [lote for lote in lotes if lote]
        # se sobrarem menos lotes que perfis marcados (poucos CPFs pra dividir),
        # usa só os primeiros perfis marcados, um lote por perfil.
        perfis_marcados = perfis_marcados[:len(lotes)]

        self.stop_event.clear()
        self._definir_estado_execucao(True)
        self._limpar_log()

        sys.stdout = TextoRedirecionado(self.fila)

        self.pasta_saida_atual = criar_pasta_execucao()
        print(f"📁 Resultados desta importação serão salvos em: {self.pasta_saida_atual}")

        id_execucao = registrar_inicio_execucao(len(lotes))
        self.execucao_atual = {"id": id_execucao, "inicio": datetime.now(), "perfis": len(lotes), "numeros_perfil": perfis_marcados}
        self._mostrar_execucao_ativa(True)
        self._atualizar_cards_topo()

        self._progresso_por_perfil = {n: (0, len(lote)) for n, lote in zip(perfis_marcados, lotes)}
        self._stats_por_perfil = {n: (0, 0, 0) for n in perfis_marcados}
        self.threads_execucao = []
        self.threads_ativas = len(lotes)

        for posicao, (numero_perfil, lote) in enumerate(zip(perfis_marcados, lotes)):
            pasta_perfil = os.path.join(PASTA_PERFIS, f"perfil_{numero_perfil}")
            t = threading.Thread(
                target=self._executar_em_thread, args=(numero_perfil, lote, pasta_perfil, posicao), daemon=True,
            )
            self.threads_execucao.append(t)
            t.start()

    def _executar_em_thread(self, numero_perfil, lote, pasta_perfil, posicao_janela):
        rotulo = rotulo_perfil_configurado(numero_perfil)
        try:
            executar_processamento(
                cpfs=lote,
                pasta_perfil=pasta_perfil,
                stop_event=self.stop_event,
                atualizar_progresso=lambda atual, total, n=numero_perfil: self._callback_progresso(n, atual, total),
                atualizar_stats=lambda s, se, e, n=numero_perfil: self._callback_stats(n, s, se, e),
                pasta_saida=self.pasta_saida_atual,
                posicao_janela=posicao_janela,
                rotulo_perfil=rotulo,
                numero_perfil=numero_perfil,
            )
        finally:
            self.fila.put(("fim_perfil", numero_perfil))

    def _callback_progresso(self, idx, atual, total):
        self.fila.put(("progresso_perfil", (idx, atual, total)))

    def _callback_stats(self, idx, sucessos, sem_inscricao, erros):
        self.fila.put(("stats_perfil", (idx, sucessos, sem_inscricao, erros)))

    def _parar(self):
        self.stop_event.set()
        for botao in (self.botao_parar_inicio, self.botao_parar_exec):
            botao.configure(state="disabled")
        if hasattr(self, "label_status_execucao"):
            self.label_status_execucao.configure(text="Status: Parando após o CPF atual de cada perfil...")

    # =====================================================
    # FILA DE EVENTOS (thread -> interface)
    # =====================================================
    def _processar_fila(self):
        try:
            while True:
                tipo, valor = self.fila.get_nowait()

                if tipo == "log":
                    self._inserir_log(valor)

                elif tipo == "progresso_perfil":
                    idx, atual, total = valor
                    self._progresso_por_perfil[idx] = (atual, total)
                    self._atualizar_widgets_execucao_atual()

                elif tipo == "stats_perfil":
                    idx, s, se, e = valor
                    self._stats_por_perfil[idx] = (s, se, e)
                    self._atualizar_widgets_execucao_atual()

                elif tipo == "fim_perfil":
                    self.threads_ativas -= 1
                    if self.threads_ativas <= 0:
                        if hasattr(self, "label_status_execucao"):
                            self.label_status_execucao.configure(text="Status: Consolidando arquivo final...")
                        threading.Thread(target=self._finalizar_processamento_thread, daemon=True).start()

                elif tipo == "finalizado_geral":
                    self._ao_finalizar_execucao()

                elif tipo == "email_detectado_perfil":
                    numero = valor
                    self.perfil_apelido_var[numero].set(
                        carregar_perfis_config().get(str(numero), {}).get("apelido", "")
                    )
                    self._atualizar_status_perfil_label(numero)
                    self._atualizar_cards_topo()
                    if hasattr(self, "_checkboxes_perfil_execucao"):
                        self._atualizar_rotulos_checkbox_perfil()

                elif tipo == "fim_verificar_perfil":
                    numero = valor
                    self._atualizar_status_perfil_label(numero)
                    botao = self._botoes_verificar_perfil.get(numero)
                    if botao is not None:
                        botao.configure(state="normal", text="Verificar login")

                elif tipo == "fim_login_manual_perfil":
                    numero = valor
                    self._atualizar_status_perfil_label(numero)

                elif tipo == "erro_verificar_perfil":
                    numero, mensagem = valor
                    botao = self._botoes_verificar_perfil.get(numero)
                    if botao is not None:
                        botao.configure(state="normal", text="Verificar login")
                    print(f"⚠️ Perfil {numero}: não foi possível verificar o login ({mensagem})")

        except queue.Empty:
            pass

        self.after(150, self._processar_fila)

    def _finalizar_processamento_thread(self):
        try:
            finalizar_processamento_geral(self.pasta_saida_atual)
        except Exception as erro:
            print(f"💥 Erro ao consolidar arquivo final: {erro}")
        finally:
            self.fila.put(("finalizado_geral", None))

    def _ao_finalizar_execucao(self):
        sys.stdout = sys.__stdout__

        if self.execucao_atual is not None:
            total_s = sum(v[0] for v in self._stats_por_perfil.values())
            total_se = sum(v[1] for v in self._stats_por_perfil.values())
            total_e = sum(v[2] for v in self._stats_por_perfil.values())
            finalizar_execucao(
                self.execucao_atual["id"], total_s, total_se, total_e,
                interrompida=self.stop_event.is_set(),
            )

        self.execucao_atual = None
        self._definir_estado_execucao(False)
        self._mostrar_execucao_ativa(False)
        self._atualizar_cards_topo()
        self._atualizar_tabelas_execucoes()
        self._carregar_perfis_na_tela()
        self._atualizar_rotulos_checkbox_perfil()
        messagebox.showinfo("Concluído", "Processamento finalizado. Confira a pasta Saída.")

    # =====================================================
    # LOG DE EXECUÇÃO
    # =====================================================
    def _configurar_tags_log(self):
        caixa = self.texto_log._textbox
        caixa.tag_config("hora", foreground="#5c6270")
        caixa.tag_config("sucesso", foreground=COR_SUCESSO)
        caixa.tag_config("erro", foreground=COR_ERRO)
        caixa.tag_config("alerta", foreground=COR_ALERTA)
        caixa.tag_config("destaque", foreground=TEMAS["Roxo Escuro"]["accent"], font=ctk.CTkFont(
            family="Consolas", size=12, weight="bold"
        ))
        caixa.tag_config("info", foreground="#c7cbd4")

    def _classificar_linha(self, linha):
        if "✅" in linha or "Sucesso" in linha:
            return "sucesso"
        if "❌" in linha or "Erro" in linha:
            return "erro"
        if any(marca in linha for marca in ("⚠️", "⏹", "⚪")):
            return "alerta"
        if any(linha.startswith(m) for m in ("🏁", "📁", "🧹", "🖼️", "🚀", "📋", "♻️", "⏱")) or "=" in linha[:6]:
            return "destaque"
        return "info"

    def _limpar_log(self):
        if not hasattr(self, "texto_log"):
            return
        self.texto_log.configure(state="normal")
        self.texto_log.delete("1.0", "end")
        self.texto_log.configure(state="disabled")

    def _inserir_log(self, linha):
        if not hasattr(self, "texto_log"):
            return

        if linha == "":
            self.texto_log.configure(state="normal")
            self.texto_log.insert("end", "\n")
            self.texto_log.configure(state="disabled")
            return

        marca_tempo = datetime.now().strftime("%H:%M:%S")
        tag = self._classificar_linha(linha)

        self.texto_log.configure(state="normal")
        self.texto_log.insert("end", f"{marca_tempo}  ", "hora")
        self.texto_log.insert("end", f"{linha}\n", tag)
        self.texto_log.see("end")
        self.texto_log.configure(state="disabled")

    # =====================================================
    # ARQUIVOS / PASTAS
    # =====================================================
    def _selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha de RAs",
            filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.caminho_cpfs.set(caminho)
            if hasattr(self, "label_arquivo_selecionado"):
                self.label_arquivo_selecionado.configure(text=os.path.basename(caminho))

    def _abrir_manual(self):
        try:
            origem = caminho_recurso(os.path.join("assets", NOME_ARQUIVO_MANUAL))
            if not os.path.exists(origem):
                messagebox.showerror(
                    "Manual não encontrado",
                    "O arquivo do manual não foi encontrado junto com o programa.",
                )
                return

            destino = os.path.join(PASTA_DADOS, NOME_ARQUIVO_MANUAL)
            if not os.path.exists(destino) or os.path.getmtime(origem) > os.path.getmtime(destino):
                shutil.copyfile(origem, destino)

            if sys.platform.startswith("win"):
                os.startfile(destino)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", destino])
            else:
                subprocess.Popen(["xdg-open", destino])
        except Exception as erro:
            messagebox.showerror("Erro ao abrir manual", f"Não foi possível abrir o manual:\n{erro}")

    def _abrir_pasta(self, pasta):
        os.makedirs(pasta, exist_ok=True)
        caminho_absoluto = os.path.abspath(pasta)
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho_absoluto)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", caminho_absoluto])
            else:
                subprocess.Popen(["xdg-open", caminho_absoluto])
        except Exception as erro:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{erro}")

    def _limpar_imagens(self):
        if not os.path.exists(PASTA_SCREENSHOTS) or not glob.glob(os.path.join(PASTA_SCREENSHOTS, "*.png")):
            messagebox.showinfo("Limpeza de imagens", "Não há imagens de log para limpar.")
            return

        confirmar = messagebox.askyesno(
            "Limpar imagens de log",
            f"Isso vai apagar todas as screenshots de erro salvas em '{PASTA_SCREENSHOTS}'.\n\nDeseja continuar?",
        )
        if not confirmar:
            return

        apagados = limpar_imagens_log()
        messagebox.showinfo("Limpeza concluída", f"{apagados} imagem(ns) removida(s) com sucesso.")
        print(f"🧹 {apagados} imagem(ns) de log removida(s).")

    def _zerar_estatisticas(self):
        confirmar = messagebox.askyesno(
            "Zerar estatísticas do painel",
            "Isso vai apagar:\n\n"
            "• O log de controle (Logs\\log_processamento.xlsx / .csv)\n"
            "• O histórico de execuções (Logs\\historico_execucoes.json)\n\n"
            "Os resultados já capturados na pasta Saída NÃO são apagados.\n\n"
            "Atenção: como o log de controle também é usado para não reprocessar "
            "CPFs que já deram certo, depois de zerar, a próxima execução pode "
            "processar de novo CPFs que já tinham sido concluídos antes.\n\n"
            "Deseja continuar?",
        )
        if not confirmar:
            return

        apagados = []
        for caminho in (ARQUIVO_LOG_EXCEL, ARQUIVO_LOG_CSV, ARQUIVO_HISTORICO):
            try:
                if os.path.exists(caminho):
                    os.remove(caminho)
                    apagados.append(os.path.basename(caminho))
            except Exception as erro:
                messagebox.showerror("Erro", f"Não foi possível apagar {caminho}:\n{erro}")
                return

        self._historico_cache = []
        self._atualizar_cards_topo()
        self._atualizar_tabelas_execucoes()

        if apagados:
            messagebox.showinfo("Estatísticas zeradas", "Arquivos removidos:\n" + "\n".join(apagados))
        else:
            messagebox.showinfo("Nada para zerar", "Não havia estatísticas registradas ainda.")

    def _ao_fechar(self):
        if any(t.is_alive() for t in self.threads_execucao):
            if not messagebox.askyesno(
                "Automação em andamento",
                "A automação ainda está em execução. Deseja realmente fechar?",
            ):
                return
            self.stop_event.set()

        sys.stdout = sys.__stdout__
        self.destroy()


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
