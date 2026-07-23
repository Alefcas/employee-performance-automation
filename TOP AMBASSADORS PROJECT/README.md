# TOP AMBASSADORS — FC locale (Demo de Portfólio)

Sistema de automação em **Python** que ranqueia "Embaixadores" de um armazém de
fulfillment com base em **Qualidade** (DPMO) e **Produtividade** (Rate na LC).
O programa extrai dados de sistemas internos, aplica regras de negócio, gera um
**ranking**, um **e-mail HTML** de reconhecimento e **cartas de reconhecimento** em PDF.

> ⚠️ **Esta é uma versão de demonstração para portfólio.**
> Todos os dados são **100% fictícios**. Os sistemas internos reais foram
> substituídos por sites HTML locais de simulação, reunidos na pasta
> **`SITE OF ALEFCAS`**. Nenhuma informação real da empresa é usada.
> O identificador do galpão é simulado como **`FC locale`** e os turnos como
> **`Turno Dia`** e **`Turno Noite`**.

## Composição dos dados da demo

Os dados fictícios são montados para exibir todos os cenários de classificação:

- **7 TOP**: 5 em Qualidade + 2 em Produtividade
- **4 BOTTOM**: 1 em Qualidade + 3 em Produtividade
- **1 INVÁLIDO**: TOP em Produtividade **e** BOTTOM em Qualidade (não recebe
  reconhecimento; a qualidade ruim invalida a boa produtividade)
- **7 MEIO**: 4 em Qualidade + 3 em Produtividade (sem destaque)

Cada embaixador pertence a um dos dois turnos (`Turno Dia` / `Turno Noite`).

> O e-mail é **apenas gerado e mostrado no terminal** (não é enviado pelo Outlook)
> e sai **sem fotos** — os cards usam as iniciais do embaixador. Ao final, o
> terminal oferece: `ENTER = abrir HTML  |  N = apenas fechar`.

---

## O que o projeto demonstra

- **Automação de navegador** com Selenium/Firefox (abre o site, aguarda e detecta o download).
- **ETL** com pandas/openpyxl: leitura de planilhas com layouts irregulares, detecção
  automática de cabeçalho, cruzamento entre fontes, deduplicação por LC Level.
- **Regras de negócio** encadeadas (filtros de opportunities, horas mínimas, número de
  associados por embaixador, prioridade de Qualidade sobre Produtividade, casos inválidos).
- **Geração de artefatos**: ranking em Excel formatado, e-mail HTML e cartas em PDF (Pillow).
- **Fluxo de terminal** guiado: ao abrir, o programa já inicia o processo completo
  desde o começo (sem tela inicial de escolha).

## Fluxo (modo completo)

1. **Employee Roster** → nomes dos funcionários e turno (`roster.html`)
2. **Onboarding (SharePoint)** → planilha `associados.xlsx` (`onboarding.html`)
3. **QuickSight** → planilha `dados_funcionarios.xlsx` (`quicksight.html`)
4. **ATLAS** → defeitos por processo, para o cálculo de DPMO (`atlas.html`).
   Na aba **Raw Reports** o usuário clica manualmente na aba e em cada processo,
   dando ENTER no programa após cada download.
5. Classificação **TOP / BOTTOM / INVÁLIDO** e geração de ranking, e-mail e cartas.

## Como rodar a demo

Pré-requisitos: **Python 3.10+**, **Firefox** e as dependências abaixo.

```bash
pip install -r requirements.txt
python "SITE OF ALEFCAS/gerar_dados_ficticios.py"   # (opcional) regenera os dados fictícios
```

Depois é só dar dois cliques em **`▶ ABRIR TOP AMBASSADORS.bat`** (Windows) ou rodar:

```bash
python "2. RECURSOS/TOP AMBASSADORS.py"
```

Ao abrir, aparece um **painel de boas-vindas** (título e descrição do projeto). Ao
clicar em **"Iniciar processo completo"**, o fluxo começa no terminal. Quando o programa
abrir um dos sites "Site of Alefcas" no Firefox, clique no botão de **download** da
página e volte ao programa — ele detecta o arquivo baixado e segue o fluxo exatamente
como faria com os sites reais.

> Dica: o `ExtracaoATLAS.xlsx` fictício já vem incluído, então a etapa do ATLAS
> pode ser reaproveitada sem depender do navegador.

## Estrutura

```
TOP AMBASSADORS PROJECT/
├── ▶ ABRIR TOP AMBASSADORS.bat      # launcher (dois cliques)
├── associados.xlsx                  # planilha de Onboarding (fictícia)
├── dados_funcionarios.xlsx          # planilha de Produtividade (fictícia)
├── 1. HISTÓRICO/                    # histórico acumulado + template das cartas
├── 2. RECURSOS/                     # código-fonte + config.ini
│   ├── AMBASSADORS.py               # motor (ETL + regras + geração)
│   ├── email_ambassador.py          # geração/preview do e-mail
│   ├── PAINEL.py                    # painel de boas-vindas (tkinter)
│   ├── TOP AMBASSADORS.py           # launcher Python (painel + processo completo)
│   └── config.ini                   # parâmetros e URLs (apontam para os sites locais)
└── SITE OF ALEFCAS/                 # sites HTML falsos + dados fictícios
    ├── roster.html / onboarding.html / quicksight.html / atlas.html
    ├── downloads/                   # arquivos que os sites "baixam"
    └── gerar_dados_ficticios.py     # gerador dos dados fictícios
```

## Observações técnicas

- As URLs no `config.ini` usam o prefixo `local:` (ex.: `local:atlas.html`), resolvido
  em tempo de execução para um caminho `file:///` absoluto. Assim o projeto é portável:
  funciona em qualquer máquina após clonar, sem editar caminhos.
- As regras de classificação (limiares de Qualidade/Produtividade, mínimos etc.) ficam
  no `config.ini` e podem ser ajustadas sem tocar no código.

---

*Projeto de portfólio. Marca, dados e sistemas reais foram anonimizados/simulados.*
