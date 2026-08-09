# Manual de Uso — Robô de Busca de Inscrições (v2)

Este manual cobre apenas o que mudou em relação à versão anterior do robô.
A forma de instalar e rodar continua a mesma:

```
pip install customtkinter selenium webdriver-manager pandas openpyxl
python robo_crm_v2.py
```

---

## 1. Novo campo: E-mail

O robô agora captura o campo **E-mail** do cadastro no CRM (mesmo padrão de
captura usado para Nome, Celular, CEP etc.) e grava numa coluna nova chamada
**Email**, posicionada logo depois de **Celular** na planilha final.

Se o cadastro não tiver e-mail preenchido, a coluna fica em branco — isso
não impede o restante do processamento.

## 2. Telefone limpo e padronizado

O campo **Celular** agora passa por uma limpeza automática:

- Remove espaços, parênteses, traços e qualquer caractere que não seja número.
- Garante o prefixo do Brasil `55` na frente.

Exemplo: `(11) 94567-9896` vira `5511945679896`.

Isso é aplicado tanto para os CPFs novos quanto, de forma automática, para os
registros que já estavam no `Resultado_parcial.xlsx` — ou seja, ao rodar de
novo o robô também corrige o formato dos registros antigos.

## 3. Colunas de data sem hora

As colunas **Data primeira insc** e **Data segunda insc** agora trazem
apenas a data (`18/07/2026`), sem o horário (`15:50`). Isso vale tanto para
os novos registros quanto para os que já existiam no arquivo parcial.

## 4. Seletor de tema

No topo da tela agora existe um seletor **🎨 Tema** com três opções:

| Tema             | Estilo        |
|------------------|---------------|
| Roxo Escuro      | Escuro (padrão original) |
| Meia-noite Azul  | Escuro, com tons de azul |
| Claro            | Fundo claro, ideal para ambientes bem iluminados |

Basta clicar na opção desejada — a troca é instantânea e não precisa
reiniciar o programa.

## 5. Múltiplos perfis do Google/Chrome (processamento em paralelo)

Agora é possível rodar a automação com **1, 2 ou 3 janelas do Chrome ao
mesmo tempo**, cada uma logada com uma conta Google diferente. Isso divide a
lista de CPFs entre os perfis e acelera o processamento total.

### Como configurar

1. No card **👥 Perfis do Google (paralelismo)**, escolha no seletor quantos
   perfis quer usar (1, 2 ou 3).
2. Clique em **▶ Iniciar Automação**. O robô vai abrir uma janela do Chrome
   para cada perfil escolhido (elas aparecem levemente deslocadas na tela
   para não ficarem uma em cima da outra).
3. **Na primeira vez que usar 2 ou 3 perfis**, faça login manualmente com a
   conta Google desejada em cada uma das janelas abertas. Esse login fica
   salvo numa pasta própria de cada perfil (`ChromeProfile/perfil_1`,
   `ChromeProfile/perfil_2`, `ChromeProfile/perfil_3`), então nas próximas
   execuções o login já estará feito.
4. A lista de CPFs é dividida automaticamente entre os perfis ativos. Os
   resultados de todos os perfis são gravados no **mesmo** arquivo final
   (`Resultado_inscricoes.xlsx`), então você não precisa juntar nada
   manualmente.
5. O log de execução mostra de qual perfil é cada linha, por exemplo:
   `[Perfil 2] 🚀 Processando: 15`.
6. Os contadores de Sucesso / Sem inscrição / Erros e a barra de progresso
   já são a soma de todos os perfis rodando.
7. O botão **⏹ Parar** interrompe todos os perfis ao mesmo tempo (cada um
   termina o CPF que estiver processando naquele momento antes de parar).

### Finalização correta com múltiplos perfis

O robô só marca **"Arquivo final atualizado"** depois que **todos** os perfis
selecionados terminarem de processar o que era deles. Quando um perfil
termina antes dos outros:

- A janela do Chrome daquele perfil é fechada automaticamente (para
  economizar recursos da máquina).
- O log mostra `[Perfil X] ✅ Perfil concluído. Aguardando os demais
  perfis...`.
- O arquivo de resultado parcial já está atualizado com o que aquele perfil
  processou, mas o **arquivo final** (`Resultado_inscricoes.xlsx`) só é
  gerado depois que o último perfil ativo também terminar.

Quando o último perfil termina, aparece no log `🏁 TODOS OS PERFIS
CONCLUÍDOS` seguido de `📁 Arquivo final atualizado`, e só aí a mensagem
"Processamento finalizado" e os botões voltam a ficar disponíveis.

### Se aparecer erro ao criar a pasta (WinError 2 / FileNotFoundError)

Em algumas máquinas, a pasta "Documentos" do Windows fica **redirecionada
pelo OneDrive** para um link que não existe mais fisicamente. Nesses casos
o Python "acha" que a pasta existe, mas qualquer tentativa de criar algo
dentro dela falha com `WinError 2`.

O programa agora tenta, em ordem, até achar um lugar onde realmente
consiga gravar:

1. `Documentos\Captador de Dados CRM`
2. `Documentos\Captador de Dados CRM` (variação "Documentos" — Windows em
   português)
3. Direto na pasta do usuário: `C:\Users\<usuario>\Captador de Dados CRM`
4. Do lado do próprio `.py`/`.exe`
5. Como último recurso, na pasta temporária do Windows

Ou seja: se o `Documentos` de alguém estiver quebrado, o programa não trava
mais — ele simplesmente usa a próxima opção da lista automaticamente. Se
quiser saber onde os dados de um usuário específico acabaram indo parar,
peça pra ele abrir o programa e clicar no botão **📁 Saída** — ele abre a
pasta certa, seja qual for.

**Atenção:** teste sempre rodando o arquivo `.py` normalmente (duplo clique
ou `python Captador_de_Dados_CRM.py` no terminal), e não colando o código
dentro de células de Jupyter/IPython — rodar em notebook pode alterar como
o Python resolve caminhos e mascarar esse tipo de erro.

### "Dados Capturados" x "Registros" — por que os números podem não bater

O card **"Dados Capturados (histórico total)"** soma **todos** os CPFs com
sucesso já registrados em `Logs\log_processamento.xlsx`, desde a primeira
vez que o programa rodou nessa máquina — inclusive antes do painel de
histórico existir.

Já a coluna **"Registros"** da tabela de execuções mostra só o total
processado (sucessos + sem inscrição + erros) **daquela execução
específica**, e a tabela só lista execuções que rodaram depois que o
histórico (`Logs\historico_execucoes.json`) passou a existir.

Ou seja: se você já tinha testado o programa antes dessa versão, o card do
topo carrega esses sucessos antigos, mas eles não aparecem como uma "linha"
na tabela de Execuções — só o CPF ficou salvo no log, não a "ficha" da
execução. **Isso é esperado, não é erro de contagem.**

Se quiser zerar e começar a contagem do zero, vá em **Configurações → 📊
Estatísticas do painel → ♻️ Zerar estatísticas do painel**. Isso apaga só o
log de controle e o histórico de execuções — os arquivos já capturados na
pasta Saída não são apagados. Atenção: como esse log também evita
reprocessar CPFs que já tiveram sucesso, depois de zerar a próxima execução
pode processar de novo CPFs que já tinham sido concluídos antes.

### Observações importantes

- Use no máximo 3 perfis simultâneos — mais que isso pode sobrecarregar a
  máquina e também aumenta a chance de o CRM bloquear por excesso de acessos
  simultâneos com a mesma conta corporativa (se todas as contas forem do
  mesmo domínio, veja com o time de TI se há algum limite de sessões).
- Se um perfil não tiver CPFs suficientes para preencher sua parte da
  divisão (por exemplo, poucos CPFs para 3 perfis), o robô ajusta
  automaticamente e não abre janelas desnecessárias.
- O arquivo de log de controle (`Logs/log_processamento.xlsx`) e o de
  resultado parcial continuam sendo únicos e compartilhados — o robô já
  cuida de evitar que duas janelas gravem por cima uma da outra ao mesmo
  tempo.

---

## 6. Rebranding: "Captador de Dados CRM"

O projeto foi renomeado. Principais mudanças:

- **Arquivo principal**: `Captador_de_Dados_CRM.py` (era `robo_crm_v2.py`).
- **Título da janela e cabeçalho**: agora mostram "Captador de Dados CRM",
  com o logo (o mesmo ícone do `.exe`) ao lado do título.
- **Tema Claro**: repaginado com a paleta laranja/creme do mockup de
  referência (fundo `#faf7f3`, cards brancos, acento `#fd6b0e`). Os temas
  "Roxo Escuro" e "Meia-noite Azul" continuam iguais.
- **Ícone**: os arquivos `assets/icone_captador.ico` (para o Windows/EXE) e
  `assets/icone_captador.png` (para o logo dentro da interface) foram
  gerados a partir da imagem enviada. **Essa pasta `assets/` precisa ficar
  ao lado do `.py` (ou ser embutida no `.exe`, veja a seção de build
  abaixo) — sem ela, o app funciona normalmente, só não mostra o ícone.**

### Pasta geral de dados do usuário

Antes, as pastas `Entrada/`, `Saida/`, `Logs/` e `ChromeProfile/` eram
criadas soltas na pasta onde o programa era executado. Isso é problemático
num `.exe` distribuído para vários usuários (pode cair em `Program Files`,
sem permissão de escrita).

Agora, na primeira execução, o programa cria automaticamente uma pasta
geral em:

```
C:\Users\<usuario>\Documents\Captador de Dados CRM\
    ├── Entrada\
    ├── Saida\
    ├── Logs\
    │   └── Screenshots\
    └── ChromeProfile\
        ├── perfil_1\
        ├── perfil_2\
        └── perfil_3\
```

Isso funciona para qualquer usuário que rodar o `.exe`, independente de
onde ele foi instalado, e sem precisar de permissão de administrador.

### Gerando o .exe (forma automática — recomendada)

Agora tem um jeito de duas etapas, sem precisar decorar comando nenhum.

**1. Organize a pasta assim, no seu computador Windows:**

```
Captador de Dados CRM\
├── Captador_de_Dados_CRM.py
├── requisitos.txt
├── Gerar_Executavel.bat
└── assets\
    ├── icone_captador.ico
    └── icone_captador.png
```

**2. Dê duplo clique em `Gerar_Executavel.bat`.**

O script sozinho vai:
1. Conferir se o Python está instalado (e avisar com um link se não estiver).
2. Criar um ambiente virtual isolado (pasta `venv`), pra não bagunçar nada
   que você já tenha instalado.
3. Instalar todas as dependências do `requisitos.txt` (customtkinter,
   selenium, webdriver-manager, pandas, openpyxl, pillow, pyinstaller).
4. Rodar o PyInstaller com o nome, ícone e recursos (`assets/`) corretos —
   já incluindo `--collect-all customtkinter` e `--collect-submodules
   selenium` (veja o porquê logo abaixo), que evitam os dois erros mais
   comuns de tema/fonte e de módulo do Selenium faltando no `.exe` final.
5. Avisar no final onde o `.exe` ficou: `dist\Captador de Dados CRM.exe`.

Isso pode levar alguns minutos na primeira vez (baixando as dependências).
Nas próximas vezes que rodar, já vai ser mais rápido.

**Testei a análise de dependências deste exato script com o PyInstaller e
ela passou sem nenhum erro de módulo faltando** — o único ajuste necessário
foi rodar no Windows de fato, já que um executável gerado em outro sistema
operacional não funciona lá (por isso o `.bat` roda o PyInstaller
localmente, na sua própria máquina).

#### Por que o `--collect-submodules selenium` é obrigatório

O Selenium 4.x carrega `ChromeOptions`, `ChromeService` etc. de forma
"preguiçosa" (só importa o módulo de verdade na hora do uso, via
`importlib`). O PyInstaller não consegue enxergar isso analisando o código
estaticamente, então sem essa flag ele **não embute** os arquivos de
`selenium/webdriver/chrome/...` no `.exe` — e o programa quebra na hora de
abrir o Chrome, com um erro como:

```
Erro fatal: No module named 'selenium.webdriver.chrome.options'
```

Reproduzi esse erro exato aqui, confirmei a causa e testei que
`--collect-submodules selenium` resolve. Se você gerou o `.exe` antes dessa
correção, é só rodar o `Gerar_Executavel.bat` de novo (ele já vem com a
flag certa) e gerar um novo `.exe`.

### Gerando o .exe manualmente (alternativa)

Se preferir rodar o comando você mesmo, sem o `.bat`:

Com o PyInstaller instalado (`pip install pyinstaller`), a partir da pasta
onde estão o `.py` e a pasta `assets/`:

```
pyinstaller --noconsole --onefile ^
  --name "Captador de Dados CRM" ^
  --icon "assets\icone_captador.ico" ^
  --add-data "assets;assets" ^
  --collect-all customtkinter ^
  --collect-submodules selenium ^
  --collect-submodules webdriver_manager ^
  Captador_de_Dados_CRM.py
```

(`^` é quebra de linha no `cmd.exe` do Windows; se for rodar tudo em uma
linha só, é só juntar sem as `^`.)

- `--name "Captador de Dados CRM"` garante que o executável final se chame
  `Captador de Dados CRM.exe`.
- `--icon` define o ícone do próprio `.exe` (o que aparece no Explorer, na
  barra de tarefas etc.).
- `--add-data "assets;assets"` embute a pasta `assets/` dentro do `.exe`,
  para o logo aparecer também dentro da interface (o código já sabe
  procurar os arquivos tanto rodando como `.py` quanto dentro do `.exe`
  compilado, via `sys._MEIPASS`).
- `--collect-all customtkinter` evita erro de tema/fonte faltando.
- `--collect-submodules selenium` e `--collect-submodules webdriver_manager`
  evitam o erro de `ModuleNotFoundError` explicado acima.

O executável final fica em `dist/Captador de Dados CRM.exe`.

## 7. Novo layout: menu lateral + dashboard

A interface foi reorganizada num layout de **menu lateral + páginas**, igual
ao mockup de referência, e vale para **os 3 temas** (Roxo Escuro,
Meia-noite Azul e Claro) — só as cores mudam entre eles.

### Menu lateral

| Página | O que tem |
|---|---|
| 🏠 Início | Dashboard: cards de estatísticas gerais, execução em andamento (com progresso e previsão de término) e as últimas 5 execuções |
| 👤 Perfis | Seleção do arquivo de CPFs e de quantos perfis do Chrome rodar em paralelo |
| ✅ Execuções | Histórico completo (até 50 execuções) |
| ⚙️ Configurações | Seletor de tema e atalhos para as pastas Saída/Logs |
| 🖥️ Logs | O log de execução em tempo real (o que antes ficava na tela única) |
| ❓ Suporte | Dicas rápidas de onde procurar quando algo dá errado |
| ℹ️ Sobre | Nome, versão e descrição do programa |

### Painel "Início" (dashboard)

- **4 cards do topo**: Dados Capturados (total histórico de sucessos, somando
  todas as execuções já feitas), Perfis Configurados (o que está selecionado
  na página Perfis), Execuções Hoje e Tempo Total Hoje — esses dois últimos
  são calculados a partir de um novo histórico de execuções salvo em
  `Logs\historico_execucoes.json`.
- **Execução em andamento**: só aparece com dados quando uma execução está
  rodando — mostra perfil(is) ativos, horário de início, status, **previsão
  de término** (estimada a partir do progresso atual) e a barra de progresso
  geral com percentual. Abaixo, 4 sub-cards: Processados, Sucessos,
  Pendentes e Erros.
- **Execuções recentes**: tabela com Data/Hora, Perfis, Status, Registros e
  Duração das últimas 5 execuções, com um botão 📁 em cada linha pra abrir a
  pasta de resultados. O botão "Ver todas as execuções" leva pro histórico
  completo, na página Execuções.

O botão **"➕ Nova Execução"** (fica no topo da página Início e também na
página Execuções) inicia o processamento com o arquivo/perfis configurados
na página Perfis. O botão **"⏹ Parar"** funciona do mesmo jeito de antes.

### Paleta oficial do tema Claro

| Cor | Uso |
|---|---|
| `#FE8310` (laranja) | Destaque, botões principais, menu ativo |
| `#FEF8F1` (creme) | Fundo geral das páginas |
| `#FCFAF9` (gelo) | Fundo dos cards e do menu lateral |
| `#525252` (cinza escuro) | Texto principal — usado em **todo** fundo claro, garantindo contraste |

Os temas escuros (Roxo Escuro e Meia-noite Azul) mantêm exatamente as
mesmas cores de antes — só a estrutura de páginas/menu lateral é nova para
eles também.

**Validação feita:** abri a interface de verdade (display virtual), naveguei
por todas as 7 páginas, troquei entre os 3 temas, simulei uma execução em
andamento com dados de exemplo, e também conferi pixel a pixel que as cores
do tema Claro batem exatamente com os códigos hexadecimais acima. Também
regerei o `.exe` de teste com as mesmas flags do `Gerar_Executavel.bat` e o
empacotamento passou sem erros.

## Resumo das colunas da planilha final

```
CPF, Nome_Completo, Data_Nascimento, Celular, Email, CEP, Bairro,
Estado, Cidade, Quantidade_Inscricoes, Possui_Multiplas, Status,
Curso, Unidade, Formato_Oferta, Data primeira insc, Canal primeira ins,
Data segunda insc, Canal segunda ins, Pagamento, Vestibular, Contrato,
Matricula, Data_Processamento
```
