---
name: pdf-to-memorial-rsc
description: >
  Gera o memorial RSC-PCCTAE completo (formatado UFV/ABNT OBRIGATÓRIO) a partir do PDF do
  Relatório Detalhado RSC emitido pelo sistema oficial da UFV. Lê o PDF, extrai dados
  estruturados, e permite que o agente PesquisAI (LLM) gere NARRATIVAS LONGAS em linguagem
  CIENTÍFICA — como o memorial de um professor titular. O run.py então formata e coloca
  nos tópicos pré-estabelecidos, gerando .md e .docx (formatado UFV/ABNT).
  v5.0: FLUXO LLM NATIVO — o agente gera textos ricos em linguagem acadêmico-científica
  para cada seção, e o run.py integra esses textos na estrutura do memorial.
  Use SEMPRE que o usuário fornecer um PDF de relatório RSC da UFV e pedir para gerar o
  memorial — ex.: "gere o memorial a partir deste PDF", "pdf para memorial", "RSC Detalhado",
  "relatório RSC", ou ao mencionar o arquivo "RSC Detalhado_*.pdf".
---

# PDF → Memorial RSC-PCCTAE — Gerador Autônomo v5.0 (NARRATIVAS LLM)

> **Gere textos LONGOS, em LINGUAGEM CIENTÍFICA, como o memorial de um professor titular.**
> O `run.py` depois formata e coloca nos tópicos pré-estabelecidos.

Gera o memorial completo de Reconhecimento de Saberes e Competências (RSC-PCCTAE)
em conformidade com o **Decreto nº 13.048, de 3 de julho de 2026** (Art. 13).

> **Decreto nº 13.048/2026:** [https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm](https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm)

---

## 🧠 O PROBLEMA (v4.0)

O gerador antigo usava **textos genéricos e hardcoded** no `run.py`:

```python
# Exemplo do que era gerado (genérico):
"Ao longo da carreira, participei ativamente de comissões e grupos de trabalho..."
"Este critério refere-se a: atividades de comissão. Foram registrados N itens..."
```

Isso resultava em narrativas curtas, repetitivas e sem profundidade —
**muito aquém do que se espera de um memorial de professor titular.**

---

## ✅ A SOLUÇÃO (v5.0)

**Arquitetura em 3 etapas:**

```
[PDF]                  Dados extraídos           Narrativas do LLM
  │                           │                         │
  ▼                           ▼                         ▼
┌─────────┐    --dump-json   ┌──────────┐              ┌─────────────┐
│ run.py  │ ───────────────> │ dados.json│  (passo 1)  │  AGENTE LLM  │
│ parser  │                  └──────────┘              │ (PesquisAI)  │
└─────────┘                                            └──────┬──────┘
       │                                                       │
       │                                              Gera narrativas
       │                                              longas/científicas
       │                                                       │
       │                                                       ▼
       │                                              ┌──────────────┐
       │                                              │narrativas.json│
       │                                              └──────┬───────┘
       │                                                     │
       │   --narrativas narrativas.json                      │
       └─────────────────────────────────────────────────────┘
                              │
                              ▼
                      ┌─────────────┐
                      │  run.py     │  Formata + insere nos
                      │  formatador │  tópicos pré-estabelecidos
                      └──────┬──────┘
                             │
                    ┌────────┴────────┐
                    ▼                 ▼
               *_MEMORIAL.md    *_MEMORIAL.docx
```

### Passo a passo para o agente PesquisAI:

---

## PASSO 1: Extrair dados do PDF

```bash
python3 run.py "RSC Detalhado_*.pdf" --dump-json dados_extraidos.json
```

Isso gera um `dados_extraidos.json` com todos os campos extraídos do PDF:
- nome, matrícula, cargo, titulação, lotação
- ano de ingresso, anos de carreira
- grupos I-VI com pontuação
- TODOS os critérios (I-01 a VI-17) com descrição, itens e pontos

---

## PASSO 2: Gerar narrativas com o LLM (VOCÊ, PesquisAI)

A partir do `dados_extraidos.json`, você deve gerar um arquivo `narrativas.json`
com textos LONGOS e em LINGUAGEM CIENTÍFICA para cada seção.

### 📋 INSTRUÇÕES PARA GERAÇÃO DE NARRATIVAS:

**REGRAS DE ESTILO — MEMORIAL DE PROFESSOR TITULAR:**

1. **Tom:** Solene, acadêmico, em primeira pessoa ("participei", "contribuí", "construí")
2. **Extensão:** Cada texto deve ter entre **200 e 600 palavras** por seção
3. **Linguagem:** Científica, formal, com vocabulário técnico da área de gestão universitária
4. **Estrutura:** Use parágrafos completos e coesos — NUNCA bullet points ou listas
5. **Profundidade:** Conecte as atividades pontuadas com o desenvolvimento institucional da UFV
6. **Contexto:** Cada narrativa deve explicar o significado da atividade, o contexto, os resultados e o aprendizado — não apenas descrever
7. **Evite:** Clareza excessiva ("participei de uma comissão"). Prefira: "Minha atuação no âmbito dos conselhos superiores permitiu-me contribuir para a formulação de políticas institucionais que impactaram diretamente o desenvolvimento acadêmico e administrativo da UFV, especialmente no que tange à..."
8. **Dados:** Incorpore os números do relatório (pontos, itens, anos) naturalmente no texto — não apenas liste

### ESTRUTURA DO JSON DE SAÍDA:

```json
{
  "introducao_quem_sou": "Texto completo da seção 1.1 (300-500 palavras)...",
  "introducao_essencia": "Texto completo da seção 1.2 (300-500 palavras)...",
  "introducao_dimensoes": [
    "**Dimensão XXX:** descrição detalhada...",
    "**Dimensão YYY:** descrição detalhada...",
    "**Dimensão ZZZ:** descrição detalhada..."
  ],
  "introducao_fundamentos": "Texto da seção 1.3 (200-300 palavras)...",

  "anexo_I_intro": "Texto introdutório do Anexo I (300-500 palavras)...",
  "anexo_I_criterios": {
    "I-01": "Texto detalhado para I-1: Conselhos superiores (200-400 palavras)...",
    "I-02": "Texto detalhado para I-2: Coordenação/presidência (200-400 palavras)...",
    "I-03": "Texto detalhado para I-3: Membro de comissões (200-400 palavras)...",
    "I-05": "Texto detalhado para I-5: Vestibulares/concursos (200-400 palavras)...",
    "I-06": "Texto detalhado para I-6: Elaboração de provas (200-400 palavras)..."
  },

  "anexo_II_intro": "Texto introdutório do Anexo II (200-400 palavras)...",
  "anexo_II_criterios": {
    "II-01": "Texto detalhado...",
    "II-02": "Texto detalhado..."
  },

  "anexo_III": "Texto completo do Anexo III (200-400 palavras)...",

  "anexo_IV_intro": "Texto introdutório do Anexo IV (200-400 palavras)...",
  "anexo_IV_criterios": {
    "IV-01": "Texto detalhado...",
    "IV-07": "Texto detalhado..."
  },

  "anexo_V_intro": "Texto introdutório do Anexo V (300-500 palavras)...",
  "anexo_V_criterios": {
    "V-01": "Texto detalhado...",
    "V-02": "Texto detalhado...",
    "V-03": "Texto detalhado..."
  },

  "anexo_VI_intro": "Texto introdutório do Anexo VI (300-500 palavras)...",
  "anexo_VI_criterios": {
    "VI-09": "Texto detalhado para livros publicados...",
    "VI-10": "Texto detalhado para artigos publicados...",
    "VI-15": "Texto detalhado para atuação como instrutor...",
    "VI-16": "Texto detalhado para coordenação de eventos..."
  },

  "reflexao_saberes": "Texto completo 9.1 (300-500 palavras)...",
  "reflexao_contribuicao": "Texto completo 9.2 (200-400 palavras)...",
  "reflexao_pedido": "Texto completo 9.3 (150-300 palavras)..."
}
```

**IMPORTANTE:** Inclua APENAS as chaves dos critérios que realmente existem nos dados extraídos. Critérios sem pontuação podem ser omitidos.

### EXEMPLO DE NARRATIVA BEM ESCRITA (para referência):

> **introducao_quem_sou** (deve soar assim — longo, científico, solene):
>
> "Meu nome é [NOME], matrícula SIAPE [MATRÍCULA], servidor público federal ocupante do cargo de [CARGO] na Universidade Federal de Viçosa (UFV), instituição à qual dedico minha trajetória profissional desde [ANO_INGRESSO]. Ao longo de [N] anos de serviço público federal, construí uma carreira pautada pelo compromisso inabalável com a excelência da gestão universitária, pela formulação e implementação de políticas institucionais que transcendem o desempenho ordinário das atribuições do cargo e se inserem no âmbito estratégico da administração pública do ensino superior. O presente memorial, elaborado em conformidade com o Art. 13 do Decreto nº 13.048/2026 e com as diretrizes estabelecidas pela Comissão para Reconhecimento de Saberes e Competências do PCCTAE (CRSC-PCCTAE) da UFV, constitui o registro reflexivo e sistematizado de uma trajetória profissional que resultou na construção de saberes e competências diferenciados — nos termos do Art. 15 do referido decreto — os quais fundamentam o pleito de Reconhecimento de Saberes e Competências no Nível VI, equivalente ao título de Doutor."

---

## PASSO 3: Executar o formatador com as narrativas

```bash
python3 run.py "RSC Detalhado_*.pdf" --narrativas narrativas.json
```

O `run.py` vai:
1. Extrair os dados do PDF novamente (ou usar os mesmos)
2. Carregar as narrativas do JSON
3. Inserir cada narrativa no tópico correspondente
4. Gerar o `.md` e o `.docx` formatados UFV/ABNT

### Saída

| Arquivo | Descrição |
|---------|-----------|
| `*_MEMORIAL.md` | Markdown UTF-8 — fonte de verdade |
| `*_MEMORIAL.docx` | Word formatado UFV/ABNT (Arial 12pt, margens 3cm/2cm) |

---

## Comandos rápidos (resumo)

```bash
# 1. Extrair dados
python3 run.py "RSC Detalhado_*.pdf" --dump-json dados_extraidos.json

# 2. (VOCÊ, LLM) Gera narrativas.json a partir de dados_extraidos.json

# 3. Gerar memorial com narrativas do LLM
python3 run.py "RSC Detalhado_*.pdf" --narrativas narrativas.json

# Fallback (sem narrativas do LLM — usa textos genéricos)
python3 run.py "RSC Detalhado_*.pdf"
```

---

## Novidades da v5.0

| Funcionalidade | Descrição |
|---|---|
| **Narrativas LLM** | O agente PesquisAI gera textos longos em linguagem científica |
| **`--narrativas`** | Novo parâmetro que aceita JSON com narrativas geradas pelo LLM |
| **`--dump-json`** | Exporta dados extraídos do PDF para alimentar o LLM |
| **Fallback genérico** | Sem `--narrativas`, o sistema gera textos dinâmicos (v4.0) |
| **Estrutura flexível** | Critérios específicos podem ter ou não narrativa; fallback por critério |
| **Reflexão completa** | Suporte a `reflexao_completa` que substitui as 3 subseções (9.1-9.3) |

---

## Parâmetros completos

| Parâmetro | Descrição |
|-----------|-----------|
| `pdf` (obrigatório) | Caminho para o PDF do Relatório Detalhado RSC |
| `--output-dir, -o` | Diretório de saída (padrão: mesmo diretório do PDF) |
| `--nome, -n` | Nome do autor (padrão: extraído do PDF) |
| `--ano-ingresso, -a` | Ano de ingresso na UFV (opcional) |
| `--auto` | Modo automático: extrai ano da data de admissão |
| `--narrativas, -N` | JSON com narrativas geradas pelo LLM |
| `--dump-json, -J` | Exporta dados extraídos para JSON (sem gerar memorial) |

---

## Estrutura do memorial gerado

```
PRÉ-TEXTUAIS: CAPA → FOLHA DE ROSTO → DEDICATÓRIA → AGRADECIMENTOS →
              EPÍGRAFE → LISTA DE SIGLAS → SUMÁRIO

1 INTRODUÇÃO — TRAJETÓRIA E FUNDAMENTOS
  1.1 Quem sou e o que apresento     ← narrativa LLM
  1.2 A essência do meu fazer profissional  ← narrativa LLM + dimensões
  1.3 Fundamentos legais              ← narrativa LLM

2 ANEXO I — Comissões                 ← narrativa LLM + narrativas por critério
3 ANEXO II — Projetos                 ← narrativa LLM + narrativas por critério
4 ANEXO III — Premiações              ← narrativa LLM
5 ANEXO IV — Responsabilidades        ← narrativa LLM + narrativas por critério
6 ANEXO V — Direção                   ← narrativa LLM + narrativas por critério
7 ANEXO VI — Produção Científica      ← narrativa LLM + narrativas por critério

8 SÍNTESE DE PONTUAÇÃO
  8.1 Quadro geral (tabela extraída do PDF)
  8.2 Verificação dos requisitos legais

9 REFLEXÃO FINAL — SABERES E COMPETÊNCIAS
  9.1 Que saberes construí?           ← narrativa LLM
  9.2 Qual minha contribuição singular? ← narrativa LLM
  9.3 Pedido                          ← narrativa LLM

REFERÊNCIAS
```

---

## Requisitos

```bash
pip install pdfplumber python-docx pyspellchecker
```

---

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

## Localização

- **Script:** `/root/.agents/skills/memorial/run.py`
- **Skill:** `/root/.agents/skills/memorial/SKILL.md`
- **Template:** `/root/.agents/skills/memorial/template/`
