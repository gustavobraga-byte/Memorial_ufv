---
name: pdf-to-memorial-rsc
description: >
  Gera o memorial RSC-PCCTAE completo (formatado UFV/ABNT) a partir do PDF do Relatório
  Detalhado RSC emitido pelo sistema oficial da UFV. Lê o PDF, extrai dados estruturados
  (nome, matrícula, pontuação por grupo, itens detalhados, nível RSC pleiteado), pergunta
  interativamente o ano de ingresso na UFV, detecta automaticamente se a equivalência é
  com Mestrado ou Doutorado, e produz autonomamente os arquivos .md, .docx (UFV) e .pdf.
  Ao final, chama a skill ufv-abnt para normalizar todos os arquivos gerados.
  Use SEMPRE que o usuário fornecer um PDF de relatório RSC da UFV e pedir para gerar o
  memorial — ex.: "gere o memorial a partir deste PDF", "pdf para memorial", "RSC Detalhado",
  "relatório RSC", ou ao mencionar o arquivo "RSC Detalhado_*.pdf".
---

# PDF → Memorial RSC-PCCTAE — Gerador Autônomo v2.0

Gera o memorial completo de Reconhecimento de Saberes e Competências (RSC-PCCTAE)
autonomamente a partir do PDF oficial do Relatório Detalhado RSC emitido pelo sistema
da UFV (Pró-Reitoria de Gestão de Pessoas), em conformidade com o **Decreto nº 13.048,
de 3 de julho de 2026**.

> **Decreto nº 13.048/2026:** [https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm](https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm)

## Fluxo

```
[PDF] RSC Detalhado_*.pdf
    │
    ▼
[1] PARSER    → Extrai dados estruturados (nome, matrícula, total pts,
    │            critérios por grupo, itens detalhados, nível RSC)
    │          → Detecta "elegível" e "equivalente" para determinar
    │            se o nível pleiteado equivale a Mestrado ou Doutorado
    ▼
[2] INTERATIVO → Pergunta ao usuário o ano de ingresso na UFV
    │             (obrigatório para o memorial)
    ▼
[3] GERADOR   → Monta .md completo conforme Art. 13 do Decreto 13.048/2026:
    │            CAPA → FOLHA DE ROSTO → DEDICATÓRIA → AGRADECIMENTOS →
    │            EPÍGRAFE → LISTA DE SIGLAS → SUMÁRIO →
    │            TRAJETÓRIA PROFISSIONAL E INDIVIDUAL (Art. 13, II) →
    │            DESCRIÇÃO DAS ATIVIDADES (Art. 13, §1º, I) →
    │            DEMONSTRAÇÃO DE ALINHAMENTO (Art. 13, §1º, II) →
    │            6 ANEXOS → SÍNTESE DE PONTUAÇÃO → REFLEXÃO FINAL → REFERÊNCIAS
    ▼
[4] FORMATADOR → .md → .docx (pandoc + python-docx): A4, Arial 12pt,
    │             margens 3L/3T/2R/2B, espaçamento 1.5, paginação UFV
    ▼
[5] UFV-ABNT  → Chama a skill ufv-abnt para normalizar e validar
    │            a formatação de todos os arquivos gerados (.md, .docx, .pdf)
    │            conforme normas ABNT NBR 14724/2024, NBR 6023/2018,
    │            NBR 10520/2023 e Manual de Normalização UFV 2025
    ▼
[6] CONVERSOR → .docx → .pdf (LibreOffice)
```

## Detecção de Nível e Equivalência

O sistema lê automaticamente no PDF o campo **"RSC Requerido"** e busca pelos termos
**"elegível"** e **"equivalente"** para determinar a equivalência acadêmica:

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
| `--ano-ingresso, -a` | Ano de ingresso na UFV (opcional; se não informado, será perguntado interativamente) |

### Fluxo interativo

Ao executar, o sistema **perguntará o ano de ingresso na UFV** se ele não for fornecido
via parâmetro `--ano-ingresso`:

```
📅 Em que ano você ingressou na UFV? 1992
```

### Exemplos

```bash
# Básico (pergunta o ano de ingresso interativamente)
python3 run.py "RSC Detalhado_03jun.pdf"

# Com diretório de saída e ano de ingresso
python3 run.py "RSC Detalhado_03jun.pdf" -o "/content/drive/My Drive/MeuMemorial" -a 1992

# Forçando nome do autor
python3 run.py "RSC Detalhado_03jun.pdf" -n "MARIA DA SILVA" -a 2005
```

### Saída

| Arquivo | Formato | Descrição |
|---------|---------|-----------|
| `*_MEMORIAL.md` | Markdown UTF-8 | Fonte de verdade — texto completo |
| `*_MEMORIAL.docx` | Word — UFV/ABNT | Formatado Arial 12pt, margens UFV |
| `*_MEMORIAL.pdf` | PDF | Pronto para impressão/entrega |

Após a geração, a **skill ufv-abnt** é acionada para validar e normalizar
definitivamente todos os arquivos conforme as normas UFV/ABNT.

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
│  (Art. 13, II — descrição da trajetória profissional e individual)
├─ TRAJETÓRIA PROFISSIONAL E INDIVIDUAL
│   ├─ 1.1 Quem sou
│   ├─ 1.2 Trajetória profissional ao longo da carreira
│   └─ 1.3 Atuação na dinâmica de ensino, pesquisa e extensão
│
│  (Art. 13, §1º, I — descrição das atividades e experiências)
├─ DESCRIÇÃO DAS ATIVIDADES E EXPERIÊNCIAS
│   └─ 2.1 Vinculação aos requisitos do Art. 3º (I a VI)
│
│  (Art. 13, §1º, II — demonstração de alinhamento)
├─ DEMONSTRAÇÃO DE ALINHAMENTO
│   └─ 3.1 Saberes, competências e nível pleiteado
│
├─ ANEXOS (I a VI)
├─ SÍNTESE DE PONTUAÇÃO
├─ REFLEXÃO FINAL
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
pip install pdfplumber python-docx
```
Além disso, **pandoc** e **LibreOffice** devem estar instalados no sistema
para a conversão .docx → .pdf.

## Skill: Normalização UFV/ABNT

Ao final da geração, a **skill ufv-abnt** é automaticamente acionada para:
- Validar a formatação de todos os arquivos gerados
- Verificar conformidade com ABNT NBR 14724/2024, NBR 6023/2018 e NBR 10520/2023
- Aplicar o Manual de Normalização UFV 2025 (PIRES; SILVA, 2025)
- Garantir que o memorial segue o modelo UFV-PPG para trabalhos acadêmicos

## Localização

- **Script:** `vault/skills/pdf-to-memorial-rsc/run.py`
- **Skill:** `vault/skills/pdf-to-memorial-rsc/SKILL.md`
- **Template:** `vault/skills/pdf-to-memorial-rsc/template/memorial_structure.md`
- **Blueprint:** `vault/skills/pdf-to-memorial-rsc/template/memorial_blueprint.py`
- **Zip:** `pdf-to-memorial-rsc-skill.zip`
