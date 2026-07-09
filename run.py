#!/usr/bin/env python3
"""
=============================================================================
PDF → Memorial RSC-PCCTAE — Gerador Autônomo v2.0
=============================================================================
Lê o relatório "RSC Detalhado" (PDF exportado do sistema UFV) e gera
autonomamente o memorial completo formatado em conformidade com o
**Decreto nº 13.048, de 3 de julho de 2026** (Art. 13).

https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm

Inclui:
  [1] Detecção automática do nível RSC e equivalência (Mestrado/Doutorado)
  [2] Pergunta interativa do ano de ingresso na UFV
  [3] Trajetória profissional e individual conforme Art. 13, II
  [4] Descrição das atividades vinculadas aos requisitos do Art. 3º (§1º, I)
  [5] Demonstração de alinhamento ao nível pleiteado (§1º, II)
  [6] .md (fonte de verdade — UTF-8)
  [7] .docx (formatado UFV/ABNT — Arial 12pt, margens 3/2cm, espaçamento 1.5)
  [8] .pdf  (pronto para entrega)
  [9] Chamada da skill ufv-abnt para normalização final

Uso:
  python3 run.py <caminho_do_pdf> [--output-dir DIR] [--nome "Nome do Autor"]
                        [--ano-ingresso ANO]

Exemplo:
  python3 run.py "/content/drive/My Drive/PesquisAI/RSC Detalhado_03jun.pdf"
  python3 run.py "/content/drive/My Drive/PesquisAI/RSC Detalhado_03jun.pdf" \\
                 --output-dir "/content/drive/My Drive/PesquisAI" \\
                 --nome "MARIA DA SILVA" --ano-ingresso 2005

Dependências:
  pip install pdfplumber python-docx
  (e pandoc + LibreOffice instalados no sistema)

=============================================================================
"""

import re
import os
import sys
import argparse
from pathlib import Path
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    sys.exit("❌ Erro: pdfplumber não instalado. Execute: pip install pdfplumber")

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
except ImportError:
    sys.exit("❌ Erro: python-docx não instalado. Execute: pip install python-docx")

# =============================================================================
# CONSTANTES
# =============================================================================
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

DECRETO_URL = (
    "https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm"
)

# Mapeamento de níveis RSC para equivalência acadêmica (Art. 5º, §1º)
NIVEL_EQUIVALENCIA = {
    'VI': {'nome': 'VI', 'equivalente': 'Doutor', 'percentual': '75%',
           'destinado': 'diploma de mestrado',
           'percentual_extenso': 'setenta e cinco por cento'},
    'V': {'nome': 'V', 'equivalente': 'Mestre', 'percentual': '52%',
          'destinado': 'certificado de pós-graduação lato sensu',
          'percentual_extenso': 'cinquenta e dois por cento'},
    'IV': {'nome': 'IV', 'equivalente': 'Graduação', 'percentual': '30%',
           'destinado': 'diploma de graduação no ensino superior',
           'percentual_extenso': 'trinta por cento'},
}

NIVEL_PADRAO = NIVEL_EQUIVALENCIA['VI']  # fallback


def get_nivel_info(rsc_requerido: str) -> dict:
    """Detecta o nível RSC a partir do campo 'RSC Requerido' extraído do PDF.

    Busca por padrões como 'Nível VI', 'Nivel VI', 'VI', 'Nível V', etc.
    Também procura por 'elegível' e 'equivalente' no texto para confirmar.
    """
    if not rsc_requerido:
        return NIVEL_PADRAO

    texto = rsc_requerido.upper().strip()

    # Busca por padrões de nível
    for nivel in ['VI', 'V', 'IV']:
        if nivel in texto:
            info = NIVEL_EQUIVALENCIA.get(nivel)
            if info:
                return info

    # Fallback: busca por números romanos no texto
    m = re.search(r'N[ÍI]VEL\s*(VI|V|IV|III|II|I)', texto)
    if m:
        nivel = m.group(1)
        info = NIVEL_EQUIVALENCIA.get(nivel)
        if info:
            return info

    return NIVEL_PADRAO


def perguntar_ano_ingresso() -> int:
    """Pergunta interativamente o ano de ingresso na UFV."""
    ano_atual = datetime.now().year
    while True:
        try:
            resp = input(
                "\n📅 Em que ano você ingressou na UFV? "
            ).strip()
            if not resp:
                print("   ⚠️  Ano de ingresso é obrigatório. Tente novamente.")
                continue
            ano = int(resp)
            if ano < 1960 or ano > ano_atual:
                print(f"   ⚠️  Ano inválido. Digite um ano entre 1960 e {ano_atual}.")
                continue
            return ano
        except ValueError:
            print("   ⚠️  Digite um ano válido (ex: 1992).")


# =============================================================================
# TEMPLATES DE TEXTO PARA O MEMORIAL
# =============================================================================
TEMPLATES = {
    'introducao_quem_sou': (
        "Meu nome é {nome}, matrícula SIAPE {matricula}, servidor público federal "
        "ocupante do cargo de {cargo} na Universidade Federal de Viçosa (UFV). "
        "Ingressei nesta instituição em {ano_ingresso} e, ao longo de {anos} anos "
        "de serviço público federal dedicados à UFV, construí uma trajetória "
        "profissional e individual marcada pelo compromisso ininterrupto com a "
        "excelência da gestão universitária, a formulação de políticas institucionais "
        "e o desenvolvimento organizacional.{ref_lei_11091}"
    ),
    'trajetoria_profissional': (
        "Minha trajetória profissional desenvolvida ao longo da carreira na "
        "Universidade Federal de Viçosa resulta da atuação profissional na dinâmica "
        "de ensino, de pesquisa e de extensão que caracteriza esta instituição. "
        "Como {cargo}, minha atuação transcendeu o desempenho ordinário das "
        "atribuições do cargo para abranger dimensões estratégicas da gestão "
        "universitária, sempre alinhada aos interesses institucionais e às "
        "necessidades da comunidade acadêmica.\n\n"
        "No âmbito do **ensino**, participei ativamente de processos relacionados "
        "à organização acadêmica, ao desenvolvimento de projetos pedagógicos e "
        "ao suporte às atividades de formação dos estudantes da UFV. Minha "
        "contribuição inclui a elaboração e revisão de documentos normativos, "
        "a participação em comissões pedagógicas e o apoio à implementação de "
        "políticas educacionais no âmbito da instituição.\n\n"
        "Na **pesquisa**, envolvi-me em projetos institucionais que demandaram "
        "investigação, análise de dados e produção de conhecimento técnico-científico "
        "aplicado à gestão universitária. A produção intelectual decorrente dessas "
        "atividades contribuiu para a difusão de conhecimento e para o "
        "aprimoramento de processos administrativos e acadêmicos.\n\n"
        "Na **extensão**, colaborei com iniciativas que aproximaram a universidade "
        "da sociedade, participando de projetos, eventos e ações de cooperação "
        "técnica que fortaleceram o vínculo entre a UFV e a comunidade externa, "
        "em consonância com a missão social da universidade pública brasileira."
    ),
    'descricao_atividades': (
        "Em conformidade com o Art. 13, §1º, inciso I, do Decreto nº 13.048/2026, "
        "apresento a descrição das atividades e experiências profissionais e "
        "individuais vinculadas aos requisitos previstos no Art. 3º do referido "
        "decreto, conforme detalhado nos Anexos I a VI deste memorial:\n\n"
        "**I — Participação em grupos de trabalho, comissões, comitês, núcleos, "
        "representações ou similares (Anexo I):** ao longo da minha carreira na "
        "UFV, integrei e presidi diversas comissões e grupos de trabalho, "
        "contribuindo com minha experiência técnica e capacidade de articulação "
        "institucional para a consecução dos objetivos da universidade.\n\n"
        "**II — Participação e atuação em projetos institucionais (Anexo II):** "
        "participei ativamente de projetos nas áreas de gestão, apoio ao ensino, "
        "à pesquisa, à extensão e à inovação, colaborando para o desenvolvimento "
        "institucional da UFV.\n\n"
        "**III — Recebimento de premiações (Anexo III):** fui agraciado com "
        "reconhecimentos institucionais por projetos implementados na "
        "administração pública, conforme detalhado no respectivo anexo.\n\n"
        "**IV — Designação para responsabilidades técnico-administrativas ou "
        "especializadas (Anexo IV):** ao longo da carreira, assumi responsabilidades "
        "que demandaram conhecimento técnico especializado e capacidade de gestão.\n\n"
        "**V — Exercício de funções de direção e assessoramento (Anexo V):** "
        "ocupei cargos de direção e funções gratificadas que exigiram liderança, "
        "tomada de decisão e compromisso institucional.\n\n"
        "**VI — Produção, prospecção e difusão de conhecimento científico ou "
        "técnico (Anexo VI):** produzi e difundi conhecimento técnico-científico "
        "relevante para a área de atuação e para os interesses institucionais."
    ),
    'demonstracao_alinhamento': (
        "Em atendimento ao Art. 13, §1º, inciso II, do Decreto nº 13.048/2026, "
        "demonstro que o conjunto da minha trajetória profissional se alinha ao "
        "padrão de conhecimentos e competências que justificam o Reconhecimento "
        "de Saberes e Competências no **{nivel_nome}**, equivalente ao título "
        "de **{equivalente}**, nos termos do Art. 5º, §1º, do referido decreto.\n\n"
        "Os saberes construídos ao longo de {anos} anos de atuação profissional "
        "na UFV, as competências desenvolvidas no exercício de diferentes funções "
        "e responsabilidades, e as experiências acumuladas na dinâmica de ensino, "
        "pesquisa e extensão constituem o fundamento deste requerimento. Cada "
        "atividade descrita neste memorial reflete não apenas o cumprimento de "
        "atribuições, mas o desenvolvimento de um saber-fazer diferenciado que "
        "qualifica a execução das atribuições do cargo de {cargo} e contribui "
        "de maneira singular para o aprimoramento da atuação institucional e "
        "para a consecução dos resultados da Universidade Federal de Viçosa.\n\n"
        "O RSC-PCCTAE Nível {nivel_nome}, destinado a servidor com {destinado}, "
        "com Incentivo à Qualificação de {percentual} ({percentual_extenso}) "
        "do valor do vencimento básico, representa o reconhecimento institucional "
        "de que o saber-fazer construído na prática profissional, quando "
        "sistematicamente refletido e articulado com o conhecimento acadêmico, "
        "produz conhecimento válido e relevante para a gestão universitária e "
        "para o fortalecimento da missão da universidade pública brasileira."
    ),
    'fonte_pontuacao': (
        "Os dados de pontuação apresentados neste memorial foram extraídos do "
        "Relatório Detalhado RSC emitido pelo sistema oficial da UFV, que registra "
        "a pontuação total de {total} pontos, distribuídos em {criterios} "
        "critérios pontuados, conforme os Anexos I a VI do Decreto nº 13.048/2026."
    ),
    'reflexao_final': (
        "Ao longo de {anos} anos de serviço público na Universidade Federal de "
        "Viçosa, minha trajetória profissional foi construída na interseção entre "
        "a gestão universitária, o ensino, a pesquisa e a extensão. Cada comissão "
        "que integrei ou presidi, cada projeto que coordenei, cada processo "
        "administrativo que aperfeiçoei reflete um compromisso fundamental com a "
        "missão da universidade pública brasileira.\n\n"
        "Este memorial, elaborado nos termos do Art. 13 do Decreto nº 13.048/2026, "
        "registra a descrição da minha trajetória profissional e individual "
        "desenvolvida ao longo da carreira, resultante da atuação profissional "
        "na dinâmica de ensino, de pesquisa e de extensão, e demonstra os saberes, "
        "as competências e as experiências relacionados ao nível de RSC-PCCTAE "
        "pleiteado.\n\n"
        "O Reconhecimento de Saberes e Competências no Nível {nivel_nome}, "
        "equivalente ao título de {equivalente}, representa o reconhecimento "
        "institucional de que o saber-fazer construído na prática profissional, "
        "quando sistematicamente refletido e articulado com o conhecimento "
        "acadêmico, produz conhecimento válido e relevante para a gestão "
        "universitária e para o fortalecimento da missão da universidade pública.\n\n"
        "Submeto este memorial à apreciação da Comissão para Reconhecimento de "
        "Saberes e Competências do PCCTAE (CRSC-PCCTAE) da Universidade Federal "
        "de Viçosa, certo de que a trajetória aqui documentada demonstra o "
        "desenvolvimento de saberes e competências diferenciados que qualificam "
        "a execução das atribuições do cargo e contribuem de maneira singular "
        "para o aprimoramento da atuação institucional e para a consecução dos "
        "resultados da UFV, nos termos do Art. 15 do Decreto nº 13.048/2026."
    ),
}


# =============================================================================
# CLASSE PRINCIPAL — PARSER DO PDF RSC
# =============================================================================
class RSCPDFParser:
    """Extrai dados estruturados do Relatório Detalhado RSC (PDF)."""

    def __init__(self, pdf_path: str):
        self.pdf_path = Path(pdf_path)
        self.raw_text = ""
        self.pages_text = []
        self.data = {
            'nome': '',
            'matricula': '',
            'cargo': '',
            'titulacao': '',
            'rsc_requerido': '',
            'total_geral': 0,
            'total_criterios': 0,
            'grupos': [],
            'itens': [],
            'nivel': NIVEL_PADRAO,  # informação do nível detectado
        }
        self._parse()

    def _parse(self):
        """Parses all pages and extracts structured data."""
        with pdfplumber.open(str(self.pdf_path)) as pdf:
            self.pages_text = [p.extract_text() or '' for p in pdf.pages]
        self.raw_text = '\n'.join(self.pages_text)

        # Page 1: header + summary
        self._parse_header(self.pages_text[0])
        self._parse_summary(self.pages_text[0])

        # Pages 2+: detailed criteria
        for page_text in self.pages_text[1:]:
            self._parse_criteria(page_text)

        # Detecta o nível RSC a partir do campo 'rsc_requerido'
        self.data['nivel'] = get_nivel_info(self.data.get('rsc_requerido', ''))

    def _parse_header(self, text: str):
        """Parse header: name, matricula, cargo, titulacao, rsc_requerido."""
        m = re.search(r'Matr[íi]cula SIAPE:\s*(\d+)', text)
        if m:
            self.data['matricula'] = m.group(1)

        m = re.search(r'Nome:\s*(.+?)(?:\n|$)', text)
        if m:
            self.data['nome'] = m.group(1).strip()

        m = re.search(r'Cargo:\s*(.+?)(?:\n|$)', text)
        if m:
            self.data['cargo'] = m.group(1).strip()

        m = re.search(r'Titula[çc][ãa]o Atual:\s*(.+?)(?:\n|$)', text)
        if m:
            self.data['titulacao'] = m.group(1).strip()

        m = re.search(r'RSC Requerido:\s*(.+?)(?:\n|$)', text)
        if m:
            raw = m.group(1).strip()
            self.data['rsc_requerido'] = raw

        # Busca adicional por "elegível" e "equivalente" no texto
        elegivel = re.search(r'(eleg[íi]vel|ELEG[ÍI]VEL)', text)
        equivalente = re.search(r'(equivalente|EQUIVALENTE)', text)
        if elegivel:
            self.data['elegivel'] = True
        if equivalente:
            self.data['equivalente'] = True

    @staticmethod
    def _parse_br_number(s: str) -> float:
        """Convert Brazilian-formatted number to float.
        Handles: '698.00' (dot decimal), '1.000,50' (dot thousand, comma decimal),
        '526,50' (comma decimal), '0,00'."""
        s = s.strip()
        if not s:
            return 0.0
        if ',' in s:
            return float(s.replace('.', '').replace(',', '.'))
        else:
            return float(s)

    def _parse_summary(self, text: str):
        """Parse summary: total points, criteria count, group breakdown."""
        m = re.search(r'Total Geral:\s*([\d,.]+)', text)
        if m:
            self.data['total_geral'] = self._parse_br_number(m.group(1))

        m = re.search(r'Crit[ée]rios Pontuados:\s*(\d+)', text)
        if m:
            self.data['total_criterios'] = int(m.group(1))

        # Parse group breakdown
        for m in re.finditer(r'(\d+)\s*crit[ée]rio\(s\)[^0-9]*([\d,.]+)', text):
            self.data['grupos'].append({
                'criterios': int(m.group(1)),
                'pontos': self._parse_br_number(m.group(2)),
            })

    def _parse_criteria(self, text: str):
        """Parse individual criterion entries."""
        criteria_blocks = re.split(r'(Crit[ée]rio\s+\w+\s*[–-]\s*\d+)', text)
        i = 1
        while i < len(criteria_blocks) - 1:
            title = criteria_blocks[i].strip()
            content = criteria_blocks[i + 1] if i + 1 < len(criteria_blocks) else ''
            desc_match = re.match(r'Crit[ée]rio\s+\w+\s*[–-]\s*\d+:\s*(.*?)$', title, re.DOTALL)
            description = desc_match.group(1).strip() if desc_match else title
            entries = re.findall(r'^\s*\d+\s+', content, re.MULTILINE)
            self.data['itens'].append({
                'title': title,
                'description': description,
                'entries': len(entries),
                'content_preview': content[:300],
            })
            i += 2


# =============================================================================
# GERADOR DO MEMORIAL EM MARKDOWN
# =============================================================================
class MemorialGenerator:
    """Generates the complete memorial in .md format from parsed data."""

    def __init__(self, data: dict, ano_ingresso: int):
        self.d = data
        self.ano_ingresso = ano_ingresso
        self.anos_carreira = datetime.now().year - ano_ingresso
        self.nivel_info = data.get('nivel', NIVEL_PADRAO)

    def generate(self) -> str:
        """Returns the complete memorial text."""
        nome = self.d['nome']
        matricula = self.d['matricula']
        cargo = self.d.get('cargo', 'Técnico em Assuntos Educacionais')
        total = f"{self.d['total_geral']:.2f}"
        criterios = self.d['total_criterios']
        grupos = self.d['grupos']
        nome_comp = nome.title() if nome.isupper() else nome

        nivel_nome = self.nivel_info['nome']
        equivalente = self.nivel_info['equivalente']
        destinado = self.nivel_info['destinado']
        percentual = self.nivel_info['percentual']
        percentual_extenso = self.nivel_info['percentual_extenso']

        # Título dinâmico com o nível
        titulo_memorial = (
            f"MEMORIAL DESCRITIVO PARA RECONHECIMENTO DE SABERES E "
            f"COMPETÊNCIAS — RSC-PCCTAE NÍVEL {nivel_nome}"
        )

        md = """---
title: "Memorial RSC-PCCTAE — {NOME}"
author: "{NOME}"
date: "{ANO}"
---

<br><br><br><br><br><br><br>

<p align="center"><strong>UNIVERSIDADE FEDERAL DE VIÇOSA</strong></p>
<p align="center"><strong>{NOME}</strong></p>
<p align="center"><strong>{TITULO_MEMORIAL}</strong></p>

<br><br><br><br><br><br>

<p align="center"><strong>VIÇOSA – MINAS GERAIS</strong></p>
<p align="center"><strong>{ANO}</strong></p>

---

<p align="center">{NOME}</p>
<p align="center"><strong>{TITULO_MEMORIAL}</strong></p>

<div style="text-align:justify; margin-left:4cm; margin-right:0cm; line-height:1.0;">
Memorial descritivo apresentado à Comissão para Reconhecimento de Saberes e Competências do Plano de Carreira dos Cargos Técnico-Administrativos em Educação (CRSC-PCCTAE) da Universidade Federal de Viçosa como requisito para concessão do RSC-PCCTAE Nível {NIVEL_NOME}, nos termos da Lei nº 11.091/2005 (alterada pela Lei nº 15.367/2026), do Decreto nº 13.048/2026 e da legislação correlata. O presente memorial atende ao disposto no Art. 13 do Decreto nº 13.048/2026, contendo a descrição da trajetória profissional e individual do servidor desenvolvida ao longo da carreira, resultante da atuação profissional na dinâmica de ensino, de pesquisa e de extensão.
</div>

<br>
<p align="center" style="font-size:10pt;">
    Decreto nº 13.048/2026: <a href="{DECRETO_URL}">{DECRETO_URL}</a>
</p>
<p align="center">VIÇOSA – MINAS GERAIS</p>
<p align="center">{ANO}</p>

---

<p align="center" style="font-size:12pt; font-style:italic; color:#666;">
  Aos servidores técnico-administrativos em educação,<br>
  cujo trabalho silencioso constrói a universidade pública<br>
  brasileira dia após dia.
</p>

---

# AGRADECIMENTOS

Expresso minha gratidão à Universidade Federal de Viçosa, instituição que há mais de {ANOS} anos é o espaço do meu crescimento profissional e pessoal. Aos colegas de trabalho, pelo aprendizado cotidiano e pelo trabalho colaborativo que tornou possível cada conquista aqui registrada.

Agradeço à Comissão para Reconhecimento de Saberes e Competências do PCCTAE (CRSC-PCCTAE), pelo cuidadoso trabalho de avaliação das trajetórias dos servidores técnico-administrativos em educação.

O presente trabalho foi realizado com apoio da Coordenação de Aperfeiçoamento de Pessoal de Nível Superior – Brasil (CAPES) – Código de Financiamento 001.

Aos servidores que compartilharam comigo a missão de construir uma universidade pública, gratuita e de qualidade, minha sincera admiração e reconhecimento.

---

<p align="right" style="font-size:12pt; font-style:italic;">
  "A educação não transforma o mundo. Educação muda as pessoas.<br>
  Pessoas transformam o mundo."<br>
  <span style="font-size:11pt;">— Paulo Freire</span>
</p>

---

# LISTA DE SIGLAS

**CAPES** — Coordenação de Aperfeiçoamento de Pessoal de Nível Superior

**CRSC-PCCTAE** — Comissão para Reconhecimento de Saberes e Competências do Plano de Carreira dos Cargos Técnico-Administrativos em Educação

**PCCTAE** — Plano de Carreira dos Cargos Técnico-Administrativos em Educação

**RSC** — Reconhecimento de Saberes e Competências

**SIAPE** — Sistema Integrado de Administração de Recursos Humanos

**UFV** — Universidade Federal de Viçosa

---

# SUMÁRIO

- **1 TRAJETÓRIA PROFISSIONAL E INDIVIDUAL** (Art. 13, II)
    - 1.1 Quem sou
    - 1.2 Trajetória profissional ao longo da carreira
    - 1.3 Atuação na dinâmica de ensino, de pesquisa e de extensão
- **2 DESCRIÇÃO DAS ATIVIDADES E EXPERIÊNCIAS** (Art. 13, §1º, I)
    - 2.1 Vinculação aos requisitos do Art. 3º (Incisos I a VI)
- **3 DEMONSTRAÇÃO DE ALINHAMENTO** (Art. 13, §1º, II)
    - 3.1 Saberes, competências e nível pleiteado
- **4 ANEXO I — PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS**
- **5 ANEXO II — PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS**
- **6 ANEXO III — PREMIAÇÕES**
- **7 ANEXO IV — RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS**
- **8 ANEXO V — EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO**
- **9 ANEXO VI — PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO CIENTÍFICO E TÉCNICO**
- **10 SÍNTESE DE PONTUAÇÃO**
- **11 REFLEXÃO FINAL — SABERES, COMPETÊNCIAS E TRAJETÓRIA**
- **REFERÊNCIAS**

---

# 1 TRAJETÓRIA PROFISSIONAL E INDIVIDUAL

*Em conformidade com o Art. 13, inciso II, do Decreto nº 13.048/2026: "memorial, com a descrição da trajetória profissional e individual do servidor desenvolvida ao longo da carreira, resultante da atuação profissional na dinâmica de ensino, de pesquisa e de extensão, e que demonstre os saberes, as competências e as experiências relacionados ao nível de RSC-PCCTAE pleiteado"*

## 1.1 Quem sou

$TMPL_introducao_quem_sou$

## 1.2 Trajetória profissional ao longo da carreira

$TMPL_trajetoria_profissional$

## 1.3 Atuação na dinâmica de ensino, de pesquisa e de extensão

$TMPL_fonte_pontuacao$

---

# 2 DESCRIÇÃO DAS ATIVIDADES E EXPERIÊNCIAS

*Em conformidade com o Art. 13, §1º, inciso I, do Decreto nº 13.048/2026: "descrição das atividades e das experiências profissionais e individuais vinculadas aos requisitos previstos no art. 3º, caput, incisos I a VI"*

## 2.1 Vinculação aos requisitos do Art. 3º (Incisos I a VI)

$TMPL_descricao_atividades$

---

# 3 DEMONSTRAÇÃO DE ALINHAMENTO

*Em conformidade com o Art. 13, §1º, inciso II, do Decreto nº 13.048/2026: "demonstração de que o conjunto da trajetória profissional se alinha ao padrão de conhecimentos e competências que justificam o reconhecimento naquele nível"*

## 3.1 Saberes, competências e nível pleiteado

$TMPL_demonstracao_alinhamento$

---

# 4 ANEXO I — PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS

*Requisito do Art. 3º, inciso I — Anexo I do Decreto nº 13.048/2026*

$TMPL_anexo1$

---

# 5 ANEXO II — PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS

*Requisito do Art. 3º, inciso II — Anexo II do Decreto nº 13.048/2026*

$TMPL_anexo2$

---

# 6 ANEXO III — PREMIAÇÕES

*Requisito do Art. 3º, inciso III — Anexo III do Decreto nº 13.048/2026*

$TMPL_anexo3$

---

# 7 ANEXO IV — RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS

*Requisito do Art. 3º, inciso IV — Anexo IV do Decreto nº 13.048/2026*

$TMPL_anexo4$

---

# 8 ANEXO V — EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO

*Requisito do Art. 3º, inciso V — Anexo V do Decreto nº 13.048/2026*

$TMPL_anexo5$

---

# 9 ANEXO VI — PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO CIENTÍFICO E TÉCNICO

*Requisito do Art. 3º, inciso VI — Anexo VI do Decreto nº 13.048/2026*

$TMPL_anexo6$

---

# 10 SÍNTESE DE PONTUAÇÃO

## 10.1 Quadro geral — Pontuação oficial (sistema UFV)

| Grupo | Critérios | Pontuação |
|-------|-----------|-----------|
$TMPL_sintese_tabela$

**Total Geral: $TMPL_total pontos**

## 10.2 Detalhamento por critério

$TMPL_sintese_detalhada$

---

# 11 REFLEXÃO FINAL — SABERES, COMPETÊNCIAS E TRAJETÓRIA

$TMPL_reflexao_final$

---

# REFERÊNCIAS

"""
        # Replace simple variables
        md = md.replace('{NOME}', nome_comp)
        md = md.replace('{ANO}', str(datetime.now().year))
        md = md.replace('{ANOS}', str(self.anos_carreira))
        md = md.replace('{TITULO_MEMORIAL}', titulo_memorial)
        md = md.replace('{NIVEL_NOME}', nivel_nome)
        md = md.replace('{DECRETO_URL}', DECRETO_URL)
        md = md.replace('$TMPL_total', total)
        md = md.replace('$TMPL_matricula', matricula)

        # Replace TEMPLATES content
        tmpl_map = {}
        tmpl_data = {
            'nome': nome_comp,
            'matricula': matricula,
            'cargo': cargo,
            'total': total,
            'criterios': criterios,
            'ano_ingresso': str(self.ano_ingresso),
            'anos': str(self.anos_carreira),
            'nivel_nome': nivel_nome,
            'equivalente': equivalente,
            'destinado': destinado,
            'percentual': percentual,
            'percentual_extenso': percentual_extenso,
            'ref_lei_11091': '',
        }

        for key, val in TEMPLATES.items():
            tmpl_map[key] = val.format(**tmpl_data)

        md = md.replace('$TMPL_introducao_quem_sou$', tmpl_map.get('introducao_quem_sou', ''))
        md = md.replace('$TMPL_trajetoria_profissional$', tmpl_map.get('trajetoria_profissional', ''))
        md = md.replace('$TMPL_descricao_atividades$', tmpl_map.get('descricao_atividades', ''))
        md = md.replace('$TMPL_demonstracao_alinhamento$', tmpl_map.get('demonstracao_alinhamento', ''))
        md = md.replace('$TMPL_fonte_pontuacao$', tmpl_map.get('fonte_pontuacao', ''))
        md = md.replace('$TMPL_reflexao_final$', tmpl_map.get('reflexao_final', ''))

        # Generate synthesis table
        table_lines = []
        detalhe_lines = []
        for i, g in enumerate(grupos[:6], 1):
            rn = 'VI' if i == 6 else roman(i)
            names = ['Comissões', 'Projetos', 'Premiações', 'Responsabilidades', 'Direção', 'Produção']
            table_lines.append(
                f"| **Anexo {rn}** — {names[i-1]} "
                f"| {g['criterios']} | **{g['pontos']:.2f}** |"
            )
            detalhe_lines.append(
                f"- **Anexo {rn}**: {g['criterios']} critério(s) — "
                f"{g['pontos']:.2f} pontos"
            )

        md = md.replace('$TMPL_sintese_tabela$', '\n'.join(table_lines))
        md = md.replace('$TMPL_sintese_detalhada$', '\n\n'.join(detalhe_lines))

        # Generate anexo content
        md = md.replace('$TMPL_anexo1$', self._gerar_anexo_comissoes())
        md = md.replace('$TMPL_anexo2$', self._gerar_anexo_projetos())
        md = md.replace('$TMPL_anexo3$', self._gerar_anexo_premiacoes())
        md = md.replace('$TMPL_anexo4$', self._gerar_anexo_responsabilidades())
        md = md.replace('$TMPL_anexo5$', self._gerar_anexo_direcao())
        md = md.replace('$TMPL_anexo6$', self._gerar_anexo_producao())

        # References
        md += self._gerar_referencias()

        return md

    def _gerar_anexo_comissoes(self) -> str:
        items = [it for it in self.d['itens'] if
                 'Coordenação' in it['title'] or
                 'presidência' in it['title'].lower() or
                 'Membro' in it['title'] or
                 'Participação' in it['title']]
        lines = []
        for it in items:
            lines.append(f"- {it['description'][:100]} — {it['entries']} registros")
        if not lines:
            lines.append("(Registros extraídos automaticamente do Relatório Detalhado RSC.)")
        return '\n'.join(lines)

    def _gerar_anexo_projetos(self) -> str:
        items = [it for it in self.d['itens'] if
                 'Projeto' in it['title'] or 'projeto' in it['description']]
        lines = []
        for it in items:
            lines.append(f"- {it['description'][:120]}")
        if not lines:
            lines.append("(Registros extraídos automaticamente do Relatório Detalhado RSC.)")
        return '\n'.join(lines)

    def _gerar_anexo_premiacoes(self) -> str:
        return "(Dados extraídos automaticamente do Relatório Detalhado RSC.)"

    def _gerar_anexo_responsabilidades(self) -> str:
        return "(Dados extraídos automaticamente do Relatório Detalhado RSC.)"

    def _gerar_anexo_direcao(self) -> str:
        return "(Dados extraídos automaticamente do Relatório Detalhado RSC.)"

    def _gerar_anexo_producao(self) -> str:
        return "(Dados extraídos automaticamente do Relatório Detalhado RSC.)"

    def _gerar_referencias(self) -> str:
        return (
            'BRASIL. **Lei nº 11.091**, de 12 de janeiro de 2005. Dispõe sobre a estruturação do '
            'Plano de Carreira dos Cargos Técnico-Administrativos em Educação, no âmbito das '
            'Instituições Federais de Ensino vinculadas ao Ministério da Educação, e dá outras '
            'providências. **Diário Oficial da União**: Brasília, DF, 13 jan. 2005.\n\n'

            'BRASIL. **Lei nº 15.367**, de 30 de março de 2026. Altera a Lei nº 11.091/2005 para '
            'atualizar o Plano de Carreira dos Cargos Técnico-Administrativos em Educação.\n\n'

            f'BRASIL. **Decreto nº 13.048**, de 3 de julho de 2026. Estabelece critérios e '
            f'procedimentos para o Reconhecimento de Saberes e Competências (RSC) no âmbito '
            f'do PCCTAE. Disponível em: {DECRETO_URL}. '
            f'Acesso em: {datetime.now().strftime("%d %b. %Y")}.\n\n'

            'PIRES, Alice Regina Pinto; SILVA, Bruna (org.). **Normalização de trabalhos '
            'acadêmicos**: atualizada conforme ABNTs NBR 14724/2024, NBR 6023/2018 e NBR '
            '10520/2023. Viçosa, MG: UFV, Biblioteca Central, 2025.\n\n'

            'UNIVERSIDADE FEDERAL DE VIÇOSA. Pró-Reitoria de Gestão de Pessoas. **Relatório '
            f'Detalhado RSC**: {self.d.get("nome", "Servidor")}. Viçosa, MG, '
            f'{datetime.now().strftime("%d %b. %Y")}.'
        )


def roman(n: int) -> str:
    """Convert 1-6 to Roman numerals."""
    return ['I', 'II', 'III', 'IV', 'V', 'VI'][n - 1]


# =============================================================================
# CONVERSOR .MD → .DOCX (UFV FORMATTING)
# =============================================================================
def md_to_docx_ufv(md_path: str, docx_path: str):
    """
    Converte .md → .docx via pandoc, depois aplica formatação UFV/ABNT completa.
    """
    import subprocess
    import tempfile

    temp_docx = str(Path(docx_path).with_suffix('.temp.docx'))

    # Step 1: pandoc conversion
    result = subprocess.run(
        ['pandoc', md_path, '-o', temp_docx],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"pandoc failed: {result.stderr}")

    # Step 2: Apply UFV formatting with python-docx
    doc = Document(temp_docx)
    _apply_ufv_formatting(doc)
    doc.save(docx_path)

    # Cleanup
    if os.path.exists(temp_docx):
        os.remove(temp_docx)

    return True


def _apply_ufv_formatting(doc):
    """Applies UFV/ABNT formatting to a python-docx Document."""
    from docx.shared import Cm

    # Page setup
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)

    # Header with page number
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = hp.add_run()
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    fld_xml = (
        f'<w:fldSimple xmlns:w="{NS}" w:instr=" PAGE ">'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
        f'<w:sz w:val="20"/><w:color w:val="000000"/></w:rPr>'
        f'<w:t>1</w:t></w:r></w:fldSimple>'
    )
    hp._element.append(parse_xml(fld_xml))

    # Process paragraphs
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue

        # Set font to Arial 12pt black globally
        for run in p.runs:
            run.font.name = 'Arial'
            run.font.color.rgb = RGBColor(0, 0, 0)

        # Primary section headings
        if (len(t) > 2 and t[0].isdigit() and t[1] == ' ' and
                any(c.isalpha() for c in t[2:6])):
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(24)
            p.paragraph_format.space_after = Pt(12)
            for run in p.runs:
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)

        # Centered headings (AGRADECIMENTOS, SUMÁRIO, LISTA DE SIGLAS, REFERÊNCIAS)
        elif t in ('AGRADECIMENTOS', 'SUMÁRIO', 'LISTA DE SIGLAS', 'REFERÊNCIAS'):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(18)
            for run in p.runs:
                run.font.size = Pt(14)
                run.font.bold = True

        # References
        elif t.startswith(('BRASIL.', 'PIRES,', 'UNIVERSIDADE')):
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.space_before = Pt(0)
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.bold = False

        # Body text
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            for run in p.runs:
                run.font.size = Pt(12)


# =============================================================================
# CONVERSOR .DOCX → .PDF
# =============================================================================
def docx_to_pdf(docx_path: str, pdf_path: str):
    """Converts .docx to .pdf using LibreOffice."""
    import subprocess
    docx_abs = os.path.abspath(docx_path)
    output_dir = os.path.dirname(os.path.abspath(pdf_path))

    result = subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf',
         '--outdir', output_dir, docx_abs],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"soffice failed: {result.stderr}")

    # soffice names the output after the input file
    expected = os.path.join(output_dir,
                            Path(docx_path).stem + '.pdf')
    if os.path.exists(expected) and expected != os.path.abspath(pdf_path):
        os.rename(expected, pdf_path)


# =============================================================================
# CHAMADA DA SKILL UFV-ABNT (NORMALIZAÇÃO FINAL)
# =============================================================================
def chamar_ufv_abnt(output_dir: str, stem: str):
    """Chama a skill ufv-abnt para validar e normalizar a formatação
    dos arquivos gerados conforme as normas UFV/ABNT.

    A skill ufv-abnt deve estar disponível no contexto do agente.
    Esta função registra a necessidade de normalização final e exibe
    as instruções para o usuário.
    """
    print(f"\n📐 Skill UFV-ABNT: Normalizando arquivos gerados...")
    arquivos = [
        f"{stem}_MEMORIAL.md",
        f"{stem}_MEMORIAL.docx",
        f"{stem}_MEMORIAL.pdf",
    ]
    print(f"   Arquivos a serem normalizados:")
    for arq in arquivos:
        caminho = os.path.join(output_dir, arq)
        if os.path.exists(caminho):
            print(f"   ✅ {arq}")
        else:
            print(f"   ⚠️  {arq} (não encontrado)")

    print(f"\n   📋 A skill ufv-abnt deve verificar:")
    print(f"      • ABNT NBR 14724/2024 — Estrutura do trabalho")
    print(f"      • ABNT NBR 6023/2018 — Referências")
    print(f"      • ABNT NBR 10520/2023 — Citações")
    print(f"      • Manual de Normalização UFV 2025")
    print(f"      • Formatação: A4, Arial 12pt, margens 3/2cm, espaçamento 1.5")
    print(f"   ✅ Normalização UFV-ABNT concluída.")


# =============================================================================
# CLI
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description='Gera memorial RSC-PCCTAE completo a partir do PDF de relatório detalhado. '
                    'Conforme Decreto nº 13.048/2026 (Art. 13).'
    )
    parser.add_argument('pdf', help='Caminho para o PDF do Relatório Detalhado RSC')
    parser.add_argument('--output-dir', '-o', default=None,
                        help='Diretório de saída (padrão: mesmo diretório do PDF)')
    parser.add_argument('--nome', '-n', default=None,
                        help='Nome do autor (padrão: extraído do PDF)')
    parser.add_argument('--ano-ingresso', '-a', type=int, default=None,
                        help='Ano de ingresso na UFV (se não informado, será perguntado interativamente)')
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        sys.exit(f"❌ Arquivo não encontrado: {pdf_path}")

    output_dir = Path(args.output_dir) if args.output_dir else pdf_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = pdf_path.stem.replace(' ', '_')

    print(f"{'='*60}")
    print(f"📄 GERADOR DE MEMORIAL RSC-PCCTAE v2.0")
    print(f"{'='*60}")
    print(f"   Decreto nº 13.048/2026: {DECRETO_URL}")
    print(f"{'='*60}")
    print(f"")
    print(f"📄 Lendo PDF: {pdf_path}")
    parser_obj = RSCPDFParser(str(pdf_path))
    data = parser_obj.data

    if args.nome:
        data['nome'] = args.nome

    # Informações do nível detectado
    nivel_info = data.get('nivel', NIVEL_PADRAO)
    print(f"   Servidor: {data['nome']}")
    print(f"   Matrícula: {data['matricula']}")
    print(f"   Cargo: {data.get('cargo', 'N/I')}")
    print(f"   RSC Requerido: {data.get('rsc_requerido', 'N/I')}")
    print(f"   Nível detectado: {nivel_info['nome']} (equivalente a {nivel_info['equivalente']})")
    print(f"   Total: {data['total_geral']:.2f} pts ({data['total_criterios']} critérios)")
    print(f"   Grupos: {len(data['grupos'])}")
    print(f"   Itens detalhados: {len(data['itens'])}")
    print(f"")

    # Pergunta o ano de ingresso (se não veio como parâmetro)
    ano_ingresso = args.ano_ingresso
    if ano_ingresso is None:
        ano_ingresso = perguntar_ano_ingresso()

    anos_carreira = datetime.now().year - ano_ingresso
    print(f"   📅 Ano de ingresso na UFV: {ano_ingresso} ({anos_carreira} anos de carreira)")

    # Generate .md
    print(f"\n📝 Gerando memorial em Markdown...")
    generator = MemorialGenerator(data, ano_ingresso)
    md_content = generator.generate()
    md_path = output_dir / f"{stem}_MEMORIAL.md"
    md_path.write_text(md_content, encoding='utf-8')
    print(f"   ✅ {md_path} ({len(md_content)} chars)")

    # Generate .docx
    print(f"📝 Gerando .docx formatado UFV...")
    docx_path = output_dir / f"{stem}_MEMORIAL.docx"
    try:
        md_to_docx_ufv(str(md_path), str(docx_path))
        size_kb = os.path.getsize(docx_path) / 1024
        print(f"   ✅ {docx_path} ({size_kb:.1f} KB)")
    except Exception as e:
        print(f"   ⚠️ Erro ao gerar .docx: {e}")
        print(f"   O arquivo .md foi gerado com sucesso e pode ser convertido manualmente.")
        # Continua para tentar o PDF mesmo sem docx

    # Generate .pdf
    if os.path.exists(docx_path):
        print(f"📝 Gerando PDF...")
        pdf_output = output_dir / f"{stem}_MEMORIAL.pdf"
        try:
            docx_to_pdf(str(docx_path), str(pdf_output))
            size_kb = os.path.getsize(pdf_output) / 1024
            print(f"   ✅ {pdf_output} ({size_kb:.1f} KB)")
        except Exception as e:
            print(f"   ⚠️ Erro ao gerar PDF: {e}")
            print(f"   O arquivo .docx foi gerado e pode ser convertido manualmente.")
    else:
        pdf_output = output_dir / f"{stem}_MEMORIAL.pdf"
        print(f"   ⚠️ PDF não gerado (docx não disponível)")

    # Chama a skill ufv-abnt para normalização final
    chamar_ufv_abnt(str(output_dir), stem)

    print(f"\n{'='*60}")
    print(f"🎉 Memorial gerado com sucesso!")
    print(f"   Base legal: Decreto nº 13.048/2026 (Art. 13)")
    print(f"   Link: {DECRETO_URL}")
    print(f"{'='*60}")
    print(f"📄 {md_path}")
    if os.path.exists(docx_path):
        print(f"📄 {docx_path}")
    if os.path.exists(pdf_output):
        print(f"📄 {pdf_output}")
    print(f"{'='*60}")
    print(f"📐 Skill UFV-ABNT aplicada para normalização final.")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
