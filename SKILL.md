---
name: pdf-to-memorial-rsc
description: >
  Gera o memorial RSC-PCCTAE completo (formatado UFV/ABNT OBRIGATÓRIO) a partir do PDF do
  Relatório Detalhado RSC emitido pelo sistema oficial da UFV. Lê o PDF, extrai dados
  estruturados (nome, matrícula, lotação, data de admissão, pontuação por grupo, TODOS os
  17 critérios com itens detalhados, nível RSC pleiteado), detecta automaticamente o ano de
  ingresso na UFV pela data de admissão, detecta automaticamente se a equivalência é com
  Mestrado ou Doutorado, e produz autonomamente os arquivos .md (fonte de verdade) e .docx
   (formatado UFV/ABNT). v3.3: --example gera memorial de exemplo com dados anônimos
   (placeholders) — exemplos .pdf/.docx removidos do pacote; run.py --example é a fonte
   única dos exemplos. v3.2: ano de ingresso automático, lotação na narrativa, epígrafe e
   agradecimentos personalizados, geração PDF removida do fluxo principal.
   v3.1: ESTRUTURA E TÓPICOS IDÊNTICOS ao memorial de referência aprovado pela CRSC-PCCTAE,
  incluindo as seções "A essência do meu fazer profissional" (3 dimensões), "Fundamentos
  legais" (requisitos do Nível VI), narrativa por critério com texto fluido em cada anexo,
  "Verificação dos requisitos legais" (8.2), "Reflexão Final" com subseções (9.1-9.3),
  e formatação UFV/ABNT obrigatória (Arial 12pt, margens 3cm/2cm, espaçamento 1.5,
  paginação, capa/folha de rosto/dedicatória/epígrafe conforme manual UFV).
  Use SEMPRE que o usuário fornecer um PDF de relatório RSC da UFV e pedir para gerar o
  memorial — ex.: "gere o memorial a partir deste PDF", "pdf para memorial", "RSC Detalhado",
  "relatório RSC", ou ao mencionar o arquivo "RSC Detalhado_*.pdf".
---

# PDF → Memorial RSC-PCCTAE — Gerador Autônomo v3.2

Gera o memorial completo de Reconhecimento de Saberes e Competências (RSC-PCCTAE)
autonomamente a partir do PDF oficial do Relatório Detalhado RSC emitido pelo sistema
da UFV (Pró-Reitoria de Gestão de Pessoas), em conformidade com o **Decreto nº 13.048,
de 3 de julho de 2026** (Art. 13).

v3.2: ano de ingresso extraído automaticamente da data de admissão; lotação incluída
na narrativa; epígrafe e agradecimentos personalizados dinâmicos; geração PDF removida
do fluxo principal (use o .docx para gerar PDF no Word/LibreOffice).

> **Decreto nº 13.048/2026:** [https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm](https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm)

## Novidades da v3.2

| Funcionalidade | Descrição |
|---|---|
| **Ano de ingresso automático** | Extraído automaticamente da data de admissão no PDF — sem pergunta interativa |
| **Lotação na narrativa** | Lotação do servidor extraída do PDF e inserida na introdução e agradecimentos |
| **Epígrafe personalizada** | Seleção inteligente de citação baseada na lotação e perfil do servidor |
| **Agradecimentos dinâmicos** | Menção personalizada à unidade de lotação |
| **Capa com brasão** | Placeholder do brasão UFV na capa |
| **PDF removido do fluxo** | Geração .pdf desativada — use o .docx para exportar PDF no Word/LibreOffice |
| **DOCX sem markdown residual** | `*italic*` e `**bold**` convertidos corretamente em tabelas, listas e parágrafos |

## Novidades da v3.1

| Funcionalidade | Descrição |
|---|---|---|
| **Estrutura idêntica ao referencial** | Seções e tópicos conforme memorial aprovado pela CRSC-PCCTAE |
| **Introdução completa** | 1.1 Quem sou e o que apresento / 1.2 A essência do meu fazer profissional (3 dimensões) / 1.3 Fundamentos legais (requisitos do Nível VI) |
| **Anexos com narrativa fluida** | Cada critério descrito em prosa, não em listas: "Memorialista: um construtor de comissões", "Uma trajetória de liderança institucional", "A face acadêmica de minha trajetória" |
| **Verificação de requisitos legais (8.2)** | Tabela comparativa: pontuação mínima, critérios, Anexo VI, titulação |
| **Reflexão Final (9.1-9.3)** | "Que saberes construí?" (5 saberes), "Qual minha contribuição singular?", "Pedido" |
| **UFV/ABNT OBRIGATÓRIA** | Arial 12pt, margens 3L/3T/2R/2B, espaçamento 1.5, paginação, capa institucional, folha de rosto com natureza, dedicatória e epígrafe sem título, agradecimentos com CAPES |
| **Extração COMPLETA** | Todos os 17 critérios com seus itens numéricos e pontos — sem truncamento |
| **Modo automático (`--auto`)** | Usa 2009 como ano de ingresso padrão, sem interação |
| **Ano automático** | Ano de ingresso extraído automaticamente da data de admissão do PDF |
| **Lotação na narrativa** | Lotação extraída do PDF e inserida na introdução |
| **Parser robusto** | Lida com artefatos de OCR e layout tabular complexo do PDF |

## Fluxo

```
[PDF] RSC Detalhado_*.pdf
    │
    ▼
[1] PARSER    → Extrai TODOS os dados estruturados:
    │            • Cabeçalho: nome, matrícula, cargo, titulação, lotação
    │            • Grupos I–VI: quantidade de critérios e pontuação
    │            • 17 critérios com descrição, itens e pontos
    │            • Nível RSC e equivalência acadêmica
    ▼
[2] ANO       → Se não informado via --ano-ingresso ou --auto,
    │            pergunta interativamente o ano de ingresso na UFV
    ▼
[3] GERADOR   → Monta .md completo conforme Art. 13 do Decreto 13.048/2026:
    │            ESTRUTURA IDÊNTICA ao memorial de referência aprovado:
    │            CAPA → FOLHA DE ROSTO → DEDICATÓRIA → AGRADECIMENTOS →
    │            EPÍGRAFE → LISTA DE SIGLAS → SUMÁRIO →
    │            1 INTRODUÇÃO — TRAJETÓRIA E FUNDAMENTOS (Art. 13, II)
    │              1.1 Quem sou e o que apresento
    │              1.2 A essência do meu fazer profissional (3 dimensões)
    │              1.3 Fundamentos legais (requisitos do Nível VI)
    │            2 ANEXO I — Comissões (narrativa por critério)
    │            3 ANEXO II — Projetos
    │            4 ANEXO III — Premiações
    │            5 ANEXO IV — Responsabilidades
    │            6 ANEXO V — Direção (narrativa de liderança)
    │            7 ANEXO VI — Produção Científica (narrativa acadêmica)
    │            8 SÍNTESE DE PONTUAÇÃO (8.1 Quadro geral + 8.2 Verificação legal)
    │            9 REFLEXÃO FINAL — SABERES E COMPETÊNCIAS (Art. 15)
    │              9.1 Que saberes construí? / 9.2 Contribuição singular / 9.3 Pedido
    │            REFERÊNCIAS
    │            Escrita fluida em prosa narrativa, não em tópicos
    ▼
[4] FORMATADOR → .md → .docx (pandoc + python-docx):
    │             A4, Arial 12pt, margens 3L/3T/2R/2B,
    │             espaçamento 1.5, paginação UFV
    │
    │   *Para gerar PDF, use "Salvar como PDF" no Word/LibreOffice*
```

## Detecção de Nível e Equivalência

O sistema lê automaticamente no PDF o campo **"RSC Requerido"** e analisa o
texto para determinar a equivalência acadêmica:

| Nível RSC | Base Legal (Art. 5º, §1º) | Equivalência |
|-----------|---------------------------|--------------|
| **Nível VI** | Servidor com diploma de **mestrado** (75%) | **Doutorado** |
| **Nível V** | Servidor com certificado de **pós-graduação lato sensu** (52%) | **Mestrado** |
| **Nível IV** | Servidor com diploma de **graduação** (30%) | **Graduação** |

A equivalência é automaticamente inserida no texto do memorial (título, folha de rosto,
reflexão final).

## Uso

```bash
python3 /caminho/para/run.py <caminho_do_pdf> [opções]
```

### Parâmetros

| Parâmetro | Descrição |
|-----------|-----------|
| `pdf` (obrigatório) | Caminho para o PDF do Relatório Detalhado RSC |
| `--output-dir, -o` | Diretório de saída (padrão: mesmo diretório do PDF) |
| `--nome, -n` | Nome do autor (padrão: extraído do PDF) |
| `--ano-ingresso, -a` | Ano de ingresso na UFV (opcional) |
| `--auto` | Modo automático: usa 2009 como ano de ingresso se não informado |

### Modo interativo

Se `--ano-ingresso` não for fornecido e `--auto` não estiver ativo,
o sistema perguntará:

```
📅 Em que ano você ingressou na UFV? 1992
```

### Exemplos

```bash
# Básico (pergunta o ano de ingresso interativamente)
python3 run.py "RSC Detalhado_03jun.pdf"

# Com diretório de saída e ano de ingresso
python3 run.py "RSC Detalhado_03jun.pdf" -o "/content/drive/My Drive/MeuMemorial" -a 1992

# Modo automático (usa 2009 como ano de ingresso)
python3 run.py "RSC Detalhado_03jun.pdf" --auto

# Forçando nome do autor
python3 run.py "RSC Detalhado_03jun.pdf" -n "MARIA DA SILVA" -a 2005 --auto
```

### Saída

| Arquivo | Formato | Descrição |
|---------|---------|-----------|
| `*_MEMORIAL.md` | Markdown UTF-8 | Fonte de verdade — texto completo em prosa |
| `*_MEMORIAL.docx` | Word — UFV/ABNT | Formatado Arial 12pt, margens UFV |

## Estrutura do memorial gerado (conforme Decreto nº 13.048/2026, Art. 13)

```
┌─ CAPA                  UNIVERSIDADE FEDERAL DE VIÇOSA → Autor → Título → Cidade → Ano
├─ FOLHA DE ROSTO        Autor → Título → Natureza (recuada 4cm) → Cidade → Ano
├─ DEDICATÓRIA           Texto à direita (sem título)
├─ AGRADECIMENTOS        Obrigatório UFV-PPG (inclui CAPES)
├─ EPÍGRAFE              (sem título)
├─ LISTA DE SIGLAS
├─ SUMÁRIO               Último pré-textual
│
│  ESTRUTURA IDÊNTICA AO MEMORIAL DE REFERÊNCIA APROVADO PELA CRSC-PCCTAE
│
├─ 1 INTRODUÇÃO — TRAJETÓRIA E FUNDAMENTOS
│   ├─ 1.1 Quem sou e o que apresento
│   ├─ 1.2 A essência do meu fazer profissional (3 dimensões)
│   └─ 1.3 Fundamentos legais (requisitos do Nível VI)
│
├─ 2 ANEXO I   — Comissões (narrativa: "Memorialista: um construtor de comissões")
│   ├─ 2.1 Memorialista: um construtor de comissões
│   ├─ 2.2 Item I-1: Conselhos superiores
│   ├─ 2.3 Item I-2: Coordenação/presidência
│   ├─ 2.4 Item I-3: Membro de comissões
│   ├─ 2.5 Item I-5: Vestibulares/concursos
│   └─ 2.6 Item I-6: Elaboração de provas
│
├─ 3 ANEXO II  — Projetos
│   ├─ 3.1 Pesquisa acadêmica: REUNI
│   └─ 3.2 Avaliação de trabalhos
│
├─ 4 ANEXO III — Premiações
│
├─ 5 ANEXO IV  — Responsabilidades Técnico-Administrativas
│   ├─ 5.1 Item IV-1: Sistemas estruturantes
│   └─ 5.2 Item IV-7: Sistemas institucionais
│
├─ 6 ANEXO V   — Direção / Assessoramento
│   ├─ 6.1 Uma trajetória de liderança institucional
│   ├─ 6.2 Item V-1: CD-02 (Pró-Reitor substituto)
│   ├─ 6.3 Item V-2: CD-03/04 (Assessor Especial)
│   ├─ 6.4 Item V-3: FG-01/02 (Chefe de Divisão)
│   └─ 6.5 Item V-4: FG-03+ (Chefia de Setor)
│
├─ 7 ANEXO VI  — Produção Científica
│   ├─ 7.1 A face acadêmica de minha trajetória
│   ├─ 7.2 Item VI-9: Livro com ISBN
│   ├─ 7.3 Item VI-10: Artigos publicados
│   ├─ 7.4 Item VI-15: Instrutor
│   └─ 7.5 Item VI-16: Coordenação de eventos
│
├─ 8 SÍNTESE DE PONTUAÇÃO
│   ├─ 8.1 Quadro geral (tabela)
│   └─ 8.2 Verificação dos requisitos legais (tabela de conformidade)
│
├─ 9 REFLEXÃO FINAL — SABERES E COMPETÊNCIAS
│   ├─ 9.1 Que saberes construí? (5 saberes)
│   ├─ 9.2 Qual minha contribuição singular?
│   └─ 9.3 Pedido
│
└─ REFERÊNCIAS
```

## Base legal

- **Lei nº 11.091/2005** — Estruturação do PCCTAE
- **Lei nº 15.367/2026** — Atualização do PCCTAE
- **Decreto nº 13.048/2026** — Critérios e procedimentos para o RSC-PCCTAE
  ([link](https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm))

  > **Art. 13, II:** O memorial deve conter "a descrição da trajetória profissional
  > e individual do servidor desenvolvida ao longo da carreira, resultante da
  > atuação profissional na dinâmica de ensino, de pesquisa e de extensão, e que
  > demonstre os saberes, as competências e as experiências relacionados ao nível
  > de RSC-PCCTAE pleiteado"

  > **Art. 13, §1º:** O memorial deverá apresentar, de forma clara e objetiva:
  > I - descrição das atividades e das experiências profissionais e individuais
  > vinculadas aos requisitos previstos no art. 3º, caput, incisos I a VI;
  > II - demonstração de que o conjunto da trajetória profissional se alinha ao
  > padrão de conhecimentos e competências que justificam o reconhecimento naquele nível.

## Requisitos

```bash
pip install pdfplumber python-docx weasyprint pyspellchecker
```

Além disso, **pandoc** deve estar instalado no sistema para conversão .md → .pdf
(via weasyprint) e .md → .docx (via referência do python-docx).

## Estrutura do skill

```
pdf-to-memorial-rsc/
├── run.py               # Gerador autônomo v3.2
├── SKILL.md             # Esta documentação
├── template/
│   ├── memorial_blueprint.py   # Blueprint estrutural
│   └── memorial_structure.md   # Template markdown
└── examples/
    ├── memorial_rsc_example.md # Exemplo de memorial gerado
    └── README.md               # Instruções do exemplo
```

## Localização

- **Script:** `vault/skills/pdf-to-memorial-rsc/run.py`
- **Skill:** `vault/skills/pdf-to-memorial-rsc/SKILL.md`
- **Template:** `vault/skills/pdf-to-memorial-rsc/template/`
- **Zip:** `pdf-to-memorial-rsc-skill-v3.0.zip`
