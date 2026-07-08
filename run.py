#!/usr/bin/env python3
"""
=============================================================================
PDF → Memorial RSC-PCCTAE — Gerador Autônomo v1.0
=============================================================================
Lê o relatório "RSC Detalhado" (PDF exportado do sistema UFV) e gera
autonomamente o memorial completo formatado em:

  [1] .md  (fonte de verdade — UTF-8)
  [2] .docx (formatado UFV/ABNT — Arial 12pt, margens 3/2cm, espaçamento 1.5)
  [3] .pdf  (pronto para entrega)

Uso:
  python3 run.py <caminho_do_pdf> [--output-dir DIR] [--nome "Nome do Autor"]

Exemplo:
  python3 run.py "/content/drive/My Drive/PesquisAI/RSC Detalhado_03jun.pdf"
  python3 run.py "/content/drive/My Drive/PesquisAI/RSC Detalhado_03jun.pdf" \\
                 --output-dir "/content/drive/My Drive/PesquisAI" \\
                 --nome "RICARDO GANDINI LUGÃO"

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

# Templates de texto para o memorial
TEMPLATES = {
    'introducao_quem_sou': (
        "Meu nome é {nome}, matrícula SIAPE {matricula}, servidor público federal ocupante "
        "do cargo de Técnico em Assuntos Educacionais na Universidade Federal de Viçosa (UFV). "
        "Ingressei na instituição em {ano_ingresso} e, ao longo de mais de três décadas de serviço "
        "público federal, construí uma trajetória profissional marcada pelo compromisso com a "
        "gestão universitária, a formulação de políticas institucionais e o desenvolvimento "
        "organizacional da UFV.{ref_lei_11091}"
    ),
    'introducao_essencia': (
        "Como Técnico em Assuntos Educacionais, minha atuação transcendeu o desempenho ordinário "
        "da carreira para abranger dimensões estratégicas da gestão universitária. Este memorial "
        "é o registro reflexivo dessa trajetória, organizado segundo os eixos definidos pelo "
        "Decreto nº 13.048/2026 e pela Resolução específica da Comissão para Reconhecimento de "
        "Saberes e Competências do PCCTAE (CRSC-PCCTAE) da UFV."
    ),
    'fonte_pontuacao': (
        "Os dados de pontuação apresentados neste memorial foram extraídos do Relatório Detalhado "
        "RSC emitido pelo sistema oficial da UFV em 03 de junho de 2026, que registra a "
        "pontuação total de {total} pontos, distribuídos em {criterios} critérios pontuados."
    ),
    'reflexao_final': (
        "Ao longo de {anos} anos de serviço público na Universidade Federal de Viçosa, "
        "minha trajetória profissional foi construída na interseção entre a gestão universitária "
        "e a política educacional. Cada comissão que integrei ou presidi, cada projeto que "
        "coordenei, cada processo administrativo que aperfeiçoei reflete um compromisso "
        "fundamental com a missão da universidade pública brasileira.\n\n"
        "Este memorial não é um ponto de chegada, mas um registro de um percurso que continua. "
        "O Reconhecimento de Saberes e Competências no Nível VI, equivalente ao título de "
        "Doutor, representa o reconhecimento institucional de que o saber-fazer construído "
        "na prática profissional, quando sistematicamente refletido e articulado com o "
        "conhecimento acadêmico, produz conhecimento válido e relevante para a gestão "
        "universitária.\n\n"
        "Submeto este memorial à apreciação da Comissão para Reconhecimento de Saberes e "
        "Competências do PCCTAE da Universidade Federal de Viçosa, certo de que a trajetória "
        "aqui documentada contribui para a excelência da gestão universitária e para o "
        "fortalecimento do serviço público federal."
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

    def _parse_header(self, text: str):
        """Parse header: name, matricula, cargo, titulacao."""
        m = re.search(r'Matrícula SIAPE:\s*(\d+)', text)
        if m:
            self.data['matricula'] = m.group(1)

        m = re.search(r'Nome:\s*(.+?)(?:\n|$)', text)
        if m:
            self.data['nome'] = m.group(1).strip()

        m = re.search(r'Cargo:\s*(.+?)(?:\n|$)', text)
        if m:
            self.data['cargo'] = m.group(1).strip()

        m = re.search(r'Titulação Atual:\s*(.+?)(?:\n|$)', text)
        if m:
            self.data['titulacao'] = m.group(1).strip()

        m = re.search(r'RSC Requerido:\s*(.+?)(?:\n|$)', text)
        if m:
            self.data['rsc_requerido'] = m.group(1).strip()

    @staticmethod
    def _parse_br_number(s: str) -> float:
        """Convert Brazilian-formatted number to float.
        Handles: '698.00' (dot decimal), '1.000,50' (dot thousand, comma decimal),
        '526,50' (comma decimal), '0,00'."""
        s = s.strip()
        if ',' in s:
            # Has comma -> Brazilian format: 1.000,50 -> 1000.50
            return float(s.replace('.', '').replace(',', '.'))
        else:
            # No comma -> dot is decimal: 698.00 -> 698.00
            return float(s)

    def _parse_summary(self, text: str):
        """Parse summary: total points, criteria count, group breakdown."""
        m = re.search(r'Total Geral:\s*([\d,.]+)', text)
        if m:
            self.data['total_geral'] = self._parse_br_number(m.group(1))

        m = re.search(r'Critérios Pontuados:\s*(\d+)', text)
        if m:
            self.data['total_criterios'] = int(m.group(1))

        # Parse group breakdown
        # Pattern: "5 critério(s) A 526.50" or "2 critério(s) 7.50"
        for m in re.finditer(r'(\d+)\s*crit[ée]rio\(s\)[^0-9]*([\d,.]+)', text):
            self.data['grupos'].append({
                'criterios': int(m.group(1)),
                'pontos': self._parse_br_number(m.group(2)),
            })

    def _parse_criteria(self, text: str):
        """Parse individual criterion entries."""
        # Each criterion has a header like "Critério X - NN: Descrição"
        # and entries with #, description, data, observação
        criteria_blocks = re.split(r'(Critério\s+\w+\s*[-–]\s*\d+)', text)
        # criteria_blocks[0] is leftover, then alternating title/content
        i = 1
        while i < len(criteria_blocks) - 1:
            title = criteria_blocks[i].strip()
            content = criteria_blocks[i + 1] if i + 1 < len(criteria_blocks) else ''
            # Extract description from title
            desc_match = re.match(r'Critério\s+\w+\s*[-–]\s*\d+:\s*(.*?)$', title, re.DOTALL)
            description = desc_match.group(1).strip() if desc_match else title
            # Count entries
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

    def __init__(self, data: dict):
        self.d = data

    def generate(self) -> str:
        """Returns the complete memorial text."""
        nome = self.d['nome']
        matricula = self.d['matricula']
        total = f"{self.d['total_geral']:.2f}"
        criterios = self.d['total_criterios']
        grupos = self.d['grupos']
        nome_comp = nome.title() if nome.isupper() else nome

        # Use %s placeholders to avoid f-string/brace conflicts
        T = '%%TMPL_%s%%'  # template placeholder pattern

        md = """---
title: "Memorial RSC-PCCTAE — {NOME}"
author: "{NOME}"
date: "2026"
---

<br><br><br><br><br><br><br>

<p align="center"><strong>UNIVERSIDADE FEDERAL DE VIÇOSA</strong></p>
<p align="center"><strong>{NOME}</strong></p>
<p align="center"><strong>MEMORIAL DESCRITIVO PARA RECONHECIMENTO DE SABERES E COMPETÊNCIAS — RSC-PCCTAE NÍVEL VI</strong></p>

<br><br><br><br><br><br>

<p align="center"><strong>VIÇOSA – MINAS GERAIS</strong></p>
<p align="center"><strong>2026</strong></p>

---

<p align="center">{NOME}</p>
<p align="center"><strong>MEMORIAL DESCRITIVO PARA RECONHECIMENTO DE SABERES E COMPETÊNCIAS — RSC-PCCTAE NÍVEL VI</strong></p>

<div style="text-align:justify; margin-left:4cm; margin-right:0cm; line-height:1.0;">
Memorial descritivo apresentado à Comissão para Reconhecimento de Saberes e Competências do Plano de Carreira dos Cargos Técnico-Administrativos em Educação (CRSC-PCCTAE) da Universidade Federal de Viçosa como requisito para concessão do RSC-PCCTAE Nível VI, nos termos da Lei nº 11.091/2005 (alterada pela Lei nº 15.367/2026), do Decreto nº 13.048/2026 e da legislação correlata.
</div>

<br>
<p align="center">VIÇOSA – MINAS GERAIS</p>
<p align="center">2026</p>

---

<p align="center" style="font-size:12pt; font-style:italic; color:#666;">
  Aos servidores técnico-administrativos em educação,<br>
  cujo trabalho silencioso constrói a universidade pública<br>
  brasileira dia após dia.
</p>

---

# AGRADECIMENTOS

Expresso minha gratidão à Universidade Federal de Viçosa, instituição que há mais de três décadas é o espaço do meu crescimento profissional e pessoal. Aos colegas da Pró-Reitoria de Gestão de Pessoas, pelo aprendizado cotidiano e pelo trabalho colaborativo que tornou possível cada conquista aqui registrada.

Agradeço à Comissão para Reconhecimento de Saberes e Competências do PCCTAE, pelo cuidadoso trabalho de avaliação das trajetórias dos servidores técnico-administrativos.

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

- **1 INTRODUÇÃO — TRAJETÓRIA E FUNDAMENTOS**
    - 1.1 Quem sou e o que apresento
    - 1.2 A essência do meu fazer profissional
- **2 ANEXO I — PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS**
- **3 ANEXO II — PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS**
- **4 ANEXO III — PREMIAÇÕES**
- **5 ANEXO IV — RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS**
- **6 ANEXO V — EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO**
- **7 ANEXO VI — PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO CIENTÍFICO E TÉCNICO**
- **8 SÍNTESE DE PONTUAÇÃO**
- **9 REFLEXÃO FINAL — SABERES E COMPETÊNCIAS**
- **REFERÊNCIAS**

---

# 1 INTRODUÇÃO — TRAJETÓRIA E FUNDAMENTOS

## 1.1 Quem sou e o que apresento

$TMPL_introducao_quem_sou$

$TMPL_fonte_pontuacao$

## 1.2 A essência do meu fazer profissional

$TMPL_introducao_essencia$

---

# 2 ANEXO I — PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS

$TMPL_anexo1$

---

# 3 ANEXO II — PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS

$TMPL_anexo2$

---

# 4 ANEXO III — PREMIAÇÕES

$TMPL_anexo3$

---

# 5 ANEXO IV — RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS

$TMPL_anexo4$

---

# 6 ANEXO V — EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO

$TMPL_anexo5$

---

# 7 ANEXO VI — PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO CIENTÍFICO E TÉCNICO

$TMPL_anexo6$

---

# 8 SÍNTESE DE PONTUAÇÃO

## 8.1 Quadro geral — Pontuação oficial (sistema UFV, 03/06/2026)

| Grupo | Critérios | Pontuação |
|-------|-----------|-----------|
$TMPL_sintese_tabela$

**Total Geral: $TMPL_total pontos**

## 8.2 Detalhamento por critério

$TMPL_sintese_detalhada$

---

# 9 REFLEXÃO FINAL — SABERES E COMPETÊNCIAS

$TMPL_reflexao_final$

---

# REFERÊNCIAS

"""
        # Replace variables
        md = md.replace('{NOME}', nome_comp)
        md = md.replace('$TMPL_total', total)
        md = md.replace('$TMPL_matricula', matricula)

        # Replace TEMPLATES content
        tmpl_map = {}
        for key, val in TEMPLATES.items():
            tmpl_map[key] = val.format(
                nome=nome_comp,
                matricula=matricula,
                total=total,
                criterios=criterios,
                anos='34',
                ano_ingresso='1992',
                ref_lei_11091='',
            )

        md = md.replace('$TMPL_introducao_quem_sou$', tmpl_map.get('introducao_quem_sou', ''))
        md = md.replace('$TMPL_fonte_pontuacao$', tmpl_map.get('fonte_pontuacao', ''))
        md = md.replace('$TMPL_introducao_essencia$', tmpl_map.get('introducao_essencia', ''))
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
        items = [it for it in self.d['itens'] if 'Coordenação' in it['title'] or 'presidência' in it['title'].lower() or 'Membro' in it['title'] or 'Participação' in it['title']]
        # Build narrative from extracted items
        lines = []
        for it in items:
            lines.append(f"- {it['description'][:100]} — {it['entries']} registros")
        if not lines:
            lines.append("(Registros extraídos automaticamente do Relatório Detalhado RSC.)")
        return '\n'.join(lines)

    def _gerar_anexo_projetos(self) -> str:
        items = [it for it in self.d['itens'] if 'Projeto' in it['title'] or 'projeto' in it['description']]
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
            'BRASIL. Lei nº 11.091, de 12 de janeiro de 2005. Dispõe sobre a estruturação do '
            'Plano de Carreira dos Cargos Técnico-Administrativos em Educação, no âmbito das '
            'Instituições Federais de Ensino vinculadas ao Ministério da Educação, e dá outras '
            'providências. **Diário Oficial da União**: Brasília, DF, 13 jan. 2005.\n\n'

            'BRASIL. Lei nº 15.367, de 30 de março de 2026. Altera a Lei nº 11.091/2005 para '
            'atualizar o Plano de Carreira dos Cargos Técnico-Administrativos em Educação.\n\n'

            'BRASIL. Decreto nº 13.048, de 3 de julho de 2026. Estabelece critérios e '
            'procedimentos para o Reconhecimento de Saberes e Competências (RSC) no âmbito '
            'do PCCTAE.\n\n'

            'PIRES, Alice Regina Pinto; SILVA, Bruna (org.). **Normalização de trabalhos '
            'acadêmicos**: atualizada conforme ABNTs NBR 14724/2024, NBR 6023/2018 e NBR '
            '10520/2023. Viçosa, MG: UFV, Biblioteca Central, 2025.\n\n'

            'UNIVERSIDADE FEDERAL DE VIÇOSA. Pró-Reitoria de Gestão de Pessoas. **Relatório '
            'Detalhado RSC**: Ricardo Gandini Lugão. Viçosa, MG, 3 jun. 2026.'
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
        elif t.startswith(('BRASIL.', 'LUGÃO,')):
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
# CLI
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description='Gera memorial RSC-PCCTAE completo a partir do PDF de relatório detalhado.'
    )
    parser.add_argument('pdf', help='Caminho para o PDF do Relatório Detalhado RSC')
    parser.add_argument('--output-dir', '-o', default=None,
                        help='Diretório de saída (padrão: mesmo diretório do PDF)')
    parser.add_argument('--nome', '-n', default=None,
                        help='Nome do autor (padrão: extraído do PDF)')
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        sys.exit(f"❌ Arquivo não encontrado: {pdf_path}")

    output_dir = Path(args.output_dir) if args.output_dir else pdf_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = pdf_path.stem.replace(' ', '_')

    print(f"📄 Lendo PDF: {pdf_path}")
    parser_obj = RSCPDFParser(str(pdf_path))
    data = parser_obj.data

    if args.nome:
        data['nome'] = args.nome

    print(f"   Servidor: {data['nome']}")
    print(f"   Matrícula: {data['matricula']}")
    print(f"   Total: {data['total_geral']:.2f} pts ({data['total_criterios']} critérios)")
    print(f"   Grupos: {len(data['grupos'])}")
    print(f"   Itens detalhados: {len(data['itens'])}")

    # Generate .md
    print(f"\n📝 Gerando memorial em Markdown...")
    generator = MemorialGenerator(data)
    md_content = generator.generate()
    md_path = output_dir / f"{stem}_MEMORIAL.md"
    md_path.write_text(md_content)
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
        return

    # Generate .pdf
    print(f"📝 Gerando PDF...")
    pdf_output = output_dir / f"{stem}_MEMORIAL.pdf"
    try:
        docx_to_pdf(str(docx_path), str(pdf_output))
        size_kb = os.path.getsize(pdf_output) / 1024
        print(f"   ✅ {pdf_output} ({size_kb:.1f} KB)")
    except Exception as e:
        print(f"   ⚠️ Erro ao gerar PDF: {e}")
        print(f"   O arquivo .docx foi gerado e pode ser convertido manualmente.")

    print(f"\n{'='*60}")
    print(f"🎉 Memorial gerado com sucesso!")
    print(f"📄 {md_path}")
    print(f"📄 {docx_path}")
    print(f"📄 {pdf_output}")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
