---
name: pdf-to-memorial-rsc
description: >
  Gera o memorial RSC-PCCTAE completo (formatado UFV/ABNT) a partir do PDF do Relatório
  Detalhado RSC emitido pelo sistema oficial da UFV. Lê o PDF, extrai dados estruturados
  (nome, matrícula, pontuação por grupo, itens detalhados), e produz autonomamente os
  arquivos .md, .docx (UFV) e .pdf. Use SEMPRE que o usuário fornecer um PDF de relatório
  RSC da UFV e pedir para gerar o memorial — ex.: "gere o memorial a partir deste PDF",
  "pdf para memorial", "RSC Detalhado", "relatório RSC", ou ao mencionar o arquivo
  "RSC Detalhado_*.pdf".
---

# PDF → Memorial RSC-PCCTAE — Gerador Autônomo v1.0

Gera o memorial completo de Reconhecimento de Saberes e Competências (RSC-PCCTAE)
autonomamente a partir do PDF oficial do Relatório Detalhado RSC emitido pelo sistema
da UFV (Pró-Reitoria de Gestão de Pessoas).

## Fluxo

```
[PDF] RSC Detalhado_*.pdf
    │
    ▼
[1] PARSER    → Extrai dados estruturados (nome, matrícula, total pts,
    │            critérios por grupo, itens detalhados)
    ▼
[2] GERADOR   → Monta .md completo com: CAPA → FOLHA DE ROSTO → DEDICATÓRIA →
    │            AGRADECIMENTOS → EPÍGRAFE → LISTA DE SIGLAS → SUMÁRIO →
    │            6 ANEXOS → SÍNTESE DE PONTUAÇÃO → REFLEXÃO FINAL → REFERÊNCIAS
    ▼
[3] FORMATADOR → .md → .docx (pandoc + python-docx): A4, Arial 12pt,
    │             margens 3L/3T/2R/2B, espaçamento 1.5, paginação UFV
    ▼
[4] CONVERSOR → .docx → .pdf (LibreOffice)
```

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

### Exemplos

```bash
# Básico
python3 run.py "RSC Detalhado_03jun.pdf"

# Com diretório de saída explícito
python3 run.py "RSC Detalhado_03jun.pdf" -o "/content/drive/My Drive/MeuMemorial"

# Forçando nome do autor
python3 run.py "RSC Detalhado_03jun.pdf" -n "RICARDO GANDINI LUGÃO"
```

### Saída

| Arquivo | Formato | Descrição |
|---------|---------|-----------|
| `*_MEMORIAL.md` | Markdown UTF-8 | Fonte de verdade — texto completo |
| `*_MEMORIAL.docx` | Word — UFV/ABNT | Formatado Arial 12pt, margens UFV |
| `*_MEMORIAL.pdf` | PDF | Pronto para impressão/entrega |

## Estrutura do memorial gerado

```
┌─ CAPA          UNIVERSIDADE FEDERAL DE VIÇOSA → Autor → Título → Cidade → Ano
├─ FOLHA DE ROSTO Autor → Título → Natureza (recuada 4cm) → Cidade → Ano
├─ DEDICATÓRIA    Texto à direita (sem título)
├─ AGRADECIMENTOS Obrigatório UFV-PPG (inclui CAPES)
├─ EPÍGRAFE       (sem título)
├─ LISTA DE SIGLAS
├─ SUMÁRIO        Último pré-textual
├─ INTRODUÇÃO     → 6 Anexos → Síntese → Reflexão Final
└─ REFERÊNCIAS    Espaço simples, alinhadas à esquerda
```

## Requisitos

```bash
pip install pdfplumber python-docx
```
Além disso, **pandoc** e **LibreOffice** devem estar instalados no sistema
para a conversão .docx → .pdf.

## Skill: Normalização UFV/ABNT

Este script implementa a formatação conforme:
- Manual de Normalização UFV 2025 (PIRES; SILVA, 2025)
- ABNT NBR 14724/2024, NBR 6023/2018, NBR 10520/2023
- Modelo UFV-PPG para teses e dissertações

## Localização

- **Script:** `vault/skills/pdf-to-memorial-rsc/run.py`
- **Skill:** `vault/skills/pdf-to-memorial-rsc/SKILL.md`
