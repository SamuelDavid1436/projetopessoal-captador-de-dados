# Captador de Dados CRM

Automação para consulta e extração de dados de inscrições no CRM Dynamics, a partir de uma lista de CPFs. Suporta múltiplos perfis do Chrome rodando em paralelo, tem um dashboard com histórico de execuções e interface gráfica com 3 temas.

![Status](https://img.shields.io/badge/status-em%20uso-brightgreen)
![Plataforma](https://img.shields.io/badge/plataforma-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-yellow)

---

## Índice

- [Visão geral](#visão-geral)
- [Funcionalidades](#funcionalidades)
- [Requisitos mínimos](#requisitos-mínimos)
- [Instalação e execução](#instalação-e-execução)
- [Como usar](#como-usar)
- [Estrutura de pastas](#estrutura-de-pastas)
- [Gerando o executável (.exe)](#gerando-o-executável-exe)
- [Estrutura do projeto (código)](#estrutura-do-projeto-código)
- [Solução de problemas](#solução-de-problemas)
- [Suporte](#suporte)

---

## Visão geral

O programa lê uma planilha Excel com uma coluna `CPF`, busca cada CPF no CRM (Dynamics 365), captura os dados pessoais e de inscrição encontrados, e consolida tudo em uma planilha de saída — sem precisar fazer a busca manualmente uma a uma.

Toda a navegação no CRM é feita via Selenium, controlando o Google Chrome instalado na máquina do usuário.

## Funcionalidades

- **Busca automática por CPF** no CRM Dynamics, com nova tentativa automática em caso de falha (até 3 tentativas por CPF).
- **Múltiplos perfis do Chrome em paralelo** (1 a 3), cada um com sua própria sessão de login, dividindo a lista de CPFs entre eles.
- **Captura completa**: nome, data de nascimento, celular (já formatado com DDI 55), e-mail, CEP, bairro, cidade, estado, e todos os dados da(s) inscrição(ões) encontrada(s).
- **Dashboard** com estatísticas gerais (dados capturados no total, perfis configurados, execuções do dia, tempo total do dia), execução em andamento com progresso e previsão de término, e histórico das execuções recentes.
- **Log de controle** que evita reprocessar CPFs que já tiveram sucesso em execuções anteriores.
- **3 temas visuais** (Roxo Escuro, Meia-noite Azul e Claro), com troca instantânea.
- **Log em tempo real** de tudo o que a automação está fazendo, com captura de screenshot automática quando um CPF falha.
- Resultado sempre salvo de forma incremental (`Resultado_parcial.xlsx`), então uma queda de energia ou fechamento inesperado não perde o que já foi processado.

## Requisitos mínimos

| Item | Mínimo | Recomendado |
|---|---|---|
| Sistema operacional | Windows 10, 64 bits | Windows 11, 64 bits |
| Memória RAM | 4 GB (1 perfil) | 8–16 GB (2–3 perfis simultâneos) |
| Processador | Dual-core | Quad-core |
| Navegador | Google Chrome instalado | — |
| Tela | 1366×768 | 1920×1080 |
| Disco | Algumas centenas de MB livres | — |

Também é necessário acesso à internet, à URL do CRM da empresa (incluindo o provedor de login/SSO), e credenciais válidas com permissão de acesso ao aplicativo de busca de inscrições.

> Requisitos detalhados, com explicação de cada item, estão no **Manual do Usuário (PDF)**.

## Instalação e execução

### Opção 1 — Usar o executável (usuário final)

1. Baixe `Captador de Dados CRM.exe`.
2. Dê duplo clique para abrir — não precisa instalar nada.
3. Na primeira execução, o programa cria automaticamente a pasta `Documentos\Captador de Dados CRM`, com as subpastas `Entrada`, `Saida`, `Logs` e `ChromeProfile`.

### Opção 2 — Rodar a partir do código-fonte (desenvolvimento)

```bash
git clone https://github.com/SamuelDavid1436/projetopessoal-captador-de-dados.git
cd projetopessoal-captador-de-dados
pip install -r requisitos.txt
python Captador_de_Dados_CRM.py
```

## Como usar

1. Abra o programa e vá na aba **Perfis**.
2. Selecione a planilha Excel (.xlsx) com a coluna `CPF` preenchida.
3. Escolha quantos perfis do Chrome (1 a 3) devem rodar em paralelo.
4. Clique em **➕ Nova Execução**.
5. Na primeira vez que usar 2 ou 3 perfis, faça login manualmente com a conta Google desejada em cada janela do Chrome que abrir — o login fica salvo para as próximas execuções.
6. Acompanhe o progresso na aba **Início** (dashboard) ou linha a linha na aba **Logs**.
7. Ao final, os resultados ficam em `Documentos\Captador de Dados CRM\Saida\Resultado_inscricoes.xlsx`.

## Estrutura de pastas

Toda a automação guarda seus dados em `Documentos\Captador de Dados CRM\`:

```
Captador de Dados CRM/
├── Entrada/
│   └── cpfs.xlsx                       # planilha de entrada (coluna CPF)
├── Saida/
│   ├── Resultado_inscricoes.xlsx       # resultado final consolidado
│   ├── Resultado_inscricoes_backup.csv
│   ├── Resultado_parcial.xlsx          # salvo incrementalmente durante a execução
│   └── Resultado_backup.csv
├── Logs/
│   ├── log_processamento.xlsx          # controle de CPFs já processados
│   ├── historico_execucoes.json        # histórico usado pelo dashboard
│   └── Screenshots/                    # prints automáticos de CPFs com erro
└── ChromeProfile/
    ├── perfil_1/
    ├── perfil_2/
    └── perfil_3/                       # pastas de login independentes por perfil
```

## Gerando o executável (.exe)

O `.exe` precisa ser gerado em uma máquina Windows (não há compilação cruzada de outro sistema operacional para Windows).

**Forma automática:**

1. Coloque na mesma pasta: `Captador_de_Dados_CRM.py`, `requisitos.txt`, `Gerar_Executavel.bat` e a pasta `assets/` (com `icone_captador.ico` e `icone_captador.png`).
2. Dê duplo clique em `Gerar_Executavel.bat`.
3. O executável final fica em `dist\Captador de Dados CRM.exe`.

O script cuida de tudo: cria um ambiente virtual, instala as dependências e roda o PyInstaller com as flags corretas (incluindo `--collect-submodules selenium`, necessário porque o Selenium usa import dinâmico para os módulos do Chrome, e `--collect-all customtkinter`, necessário para os temas/fontes da interface).

**Forma manual**, se preferir rodar o comando você mesmo:

```bash
pip install -r requisitos.txt
pyinstaller --noconsole --onefile ^
  --name "Captador de Dados CRM" ^
  --icon "assets\icone_captador.ico" ^
  --add-data "assets;assets" ^
  --collect-all customtkinter ^
  --collect-submodules selenium ^
  --collect-submodules webdriver_manager ^
  Captador_de_Dados_CRM.py
```

## Estrutura do projeto (código)

```
Captador_de_Dados_CRM.py
├── Navegação no CRM           (preencher CPF, buscar, ler resultado)
├── Captura de dados            (inscrições + dados pessoais)
├── Limpeza/formatação          (telefone, datas, e-mail)
├── Log de controle             (evita reprocessar CPFs já concluídos)
├── Resultado parcial/final     (Excel + CSV)
├── Histórico de execuções      (alimenta o dashboard)
├── Execução principal          (orquestra 1–3 perfis em paralelo)
└── Interface gráfica (customtkinter)
    ├── Sidebar + 7 páginas (Início, Perfis, Execuções, Configurações, Logs, Suporte, Sobre)
    └── 3 temas (Roxo Escuro, Meia-noite Azul, Claro)
```

**Dependências principais:** `customtkinter`, `selenium`, `webdriver-manager`, `pandas`, `openpyxl`, `pillow`.

## Solução de problemas

| Sintoma | Causa provável | Solução |
|---|---|---|
| `No module named 'selenium.webdriver.chrome.options'` no `.exe` | Faltou `--collect-submodules selenium` no build | Gere o `.exe` de novo com `Gerar_Executavel.bat` |
| Erro ao criar a pasta de dados (`WinError 2`) | Pasta "Documentos" redirecionada pelo OneDrive de forma quebrada | O programa já tenta locais alternativos automaticamente; nada a fazer |
| "Dados Capturados" não bate com a soma da tabela de Execuções | São métricas diferentes (total histórico x execução específica) | Ver explicação na aba Configurações, ou zerar as estatísticas do painel |
| Perfil pedindo login toda vez | Sessão do Google expirada naquele perfil | Fazer login manualmente de novo na janela daquele perfil |

Para mais detalhes, consulte o **Manual do Usuário (PDF)**, que traz o passo a passo de cada tela com capturas de tela.

## Suporte

- **Telefone:** (11) 9472-8128
- **E-mail:** samueldayvid5@icloud.com

Ao entrar em contato, inclua sempre que possível o arquivo `Logs\log_processamento.xlsx` e uma descrição do que aconteceu — isso acelera o diagnóstico.

---

© 2026 Captador de Dados CRM. Uso interno.
