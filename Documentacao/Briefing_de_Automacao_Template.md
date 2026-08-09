# Briefing de Automação — Template

Preencha isso antes de pedir o código. Não precisa ser longo — só completo.

## 1. Contexto
- **O que a automação faz, em uma frase:**
- **Quem vai usar:** (você mesmo? equipe leiga? quantas pessoas?)
- **Onde vai rodar:** (seu PC só / .exe distribuído pra outros / servidor)
  - Se for `.exe` distribuído: já avisa aqui. Isso muda decisões de pasta, ícone, nome, tratamento de erro desde o início.

## 2. Dados e regras de negócio
- **Entrada:** formato do arquivo, colunas obrigatórias
- **Cada campo que precisa ser capturado**, já com a regra de formatação de cada um
  (ex: "telefone sem espaços, com 55 na frente"; "data sem hora")
- **Saída:** formato do arquivo, nome das colunas, ordem

## 3. Volume e performance
- Quantos registros por execução, em média e no pico?
- Precisa rodar em paralelo (múltiplas janelas/instâncias)? Quantas no máximo?

## 4. Identidade visual (se tiver interface)
- **Nome do produto**
- **Print de referência ou paleta de cores em hex**, se já tiver algo em mente
- **Ícone**, se já tiver
- Temas (claro/escuro) são necessários?

## 5. Empacotamento e distribuição
- Formato final: script `.py`, `.exe`, ambos?
- Nome do executável
- Onde salvar dados na máquina do usuário
- Precisa rodar em máquina sem o Python instalado?

## 6. Documentação
- Formato: Word, PDF, Markdown?
- Público: usuário final leigo ou técnico?
- Precisa de capturas de tela?
- Dados de suporte (telefone, e-mail) a incluir

## 7. Definição de "pronto"
- Como eu vou saber que terminou? (ex: "roda o .exe numa máquina limpa, sem Python, e funciona")
