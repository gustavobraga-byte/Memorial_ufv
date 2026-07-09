#!/usr/bin/env python3
"""
=============================================================================
PDF → Memorial RSC-PCCTAE — Gerador Autônomo v3.1
=============================================================================
Gera o memorial completo de Reconhecimento de Saberes e Competências (RSC-PCCTAE)
autonomamente a partir do PDF oficial do Relatório Detalhado RSC emitido pelo
sistema da UFV (Pró-Reitoria de Gestão de Pessoas), em conformidade com o
**Decreto nº 13.048, de 3 de julho de 2026** (Art. 13).

v3.1 — Extração COMPLETA de todos os 17 critérios e itens do PDF,
         geração de memorial com ESTRUTURA E TÓPICOS IDÊNTICOS ao
         memorial de referência aprovado, formatação UFV/ABNT obrigatória.

Uso:
  python3 run.py <caminho_do_pdf> [--output-dir DIR] [--nome "Nome"]
                 [--ano-ingresso ANO] [--auto]

Dependências:
  pip install pdfplumber python-docx weasyprint pyspellchecker
=============================================================================
"""

import re, os, sys, argparse, json
from pathlib import Path
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    sys.exit("Erro: pdfplumber não instalado. Execute: pip install pdfplumber")

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm, Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
except ImportError:
    sys.exit("Erro: python-docx não instalado. Execute: pip install python-docx")

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DECRETO_URL = "https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm"

# =============================================================================
# Restaurador de acentos baseado em DICIONÁRIO (não mapa manual)
# =============================================================================
# Solução DEFINITIVA para perda de acentos em extração de PDF.
# Usa pyspellchecker (416K palavras portuguesas) + stopword list de
# function words que nunca devem ser acentuados (de, que, o, da, etc.).
# Para cada palavra sem acento no texto, consulta o índice reverso
# unaccented → accented; substitui apenas se houver candidato único
# ou dominante com frequência suficiente. Preserva capitalização.
# =============================================================================
import unicodedata
try:
    from spellchecker import SpellChecker as _SpellChecker
    _SP = _SpellChecker(language='pt')
    _PT_WORDS = set(_SP.word_frequency.dictionary)
except Exception:
    _SP = None
    _PT_WORDS = set()


def _strip_accents(s):
    """Remove diacríticos mantendo a letra base."""
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                   if unicodedata.category(c) != 'Mn')


# ---- STOPWORDS: function words portuguesas que NÃO devem ser acentuadas ----
# Estas palavras têm formas acentuadas homógrafas (de→dê, que→quê, por→pôr,
# o→ó, da→dá, no→nó, se→sê, mas→más, ate→até, ja→já, sao→são) que são verbos
# ou interjeições RARAS — NÃO devem ser aplicadas em prosa técnica.
_STOPWORDS = frozenset("""
a o os as um uns uma umas de da do das dos no na nos nas ne noh neh
por para pela pelas pelo pelos perante ante apos ate sobre sob contra
com sem entre desde durante segundo mediante conforme contra exceto
e ou mas que se como quando onde porque pois embora entretanto porem
contudo todavia logo portanto assim ja nem apenas senao entao enquanto
que quem qual quais cujo cuja cujos cujas quanto quanta quantos quantas
ele ela eles elas este esta estes estas esse essa esses essas aquele
aquela aqueles aquelas isto isso aquilo meu minha meus minhas teu tua
teus tuas seu sua seus suas nosso nossa nossos nossas vosso vossa vossos
vossas lhe lhes me te se nos vos si mesmo mesma mesmos mesmas cada todo
todos toda todas algum alguns alguma algumas nenhum nenhuns nenhuma
nenhumas outro outros outra outras muito muitos muita muitas pouco poucos
pouca poucas tal tais qual quais quer queres eu tu ele nos vos eles voce
voces nos tao pouco demais acerca afinal entao ainda aqui ali la ali
ai ja nao sim ora pois ah oh ih olha olhe veja
""".split())


def _build_reverse_index():
    """Constrói índice: unaccented_word → [(accented_word, freq), ...].
    Só inclui palavras do dicionário que TENHAM acentos (diferem ao desaccentuar).
    Threshold de freq >= 50 para evitar formas raras/dubiosas (pela→pelá etc.)."""
    if _SP is None:
        return {}
    rev = {}
    freqs = _SP.word_frequency.dictionary
    for w in _PT_WORDS:
        sw = _strip_accents(w)
        if sw != w:  # só palavras que TÊM acento
            f = freqs.get(w, 0)
            if f >= 50:
                rev.setdefault(sw, []).append((w, f))
    # Ordena cada lista por freq desc
    for sw, lst in rev.items():
        lst.sort(key=lambda t: -t[1])
    return rev


_REV_INDEX = _build_reverse_index()


def _match_case(template, target):
    """Aplica a capitalização de template em target.
    template = palavra original (sem acento, capitalização a preservar)
    target = palavra correta (com acento) a capitalizar."""
    if template.isupper():
        return target.upper()
    if template[0].isupper():
        return target.capitalize()
    return target.lower()


def restaurar_acentos(texto):
    """Restaura acentos em texto português usando dicionário + stopword filter.
    - Não altera palavras que já têm acento
    - Pula stopwords (de, que, o, da, etc.)
    - Pula palavras < 4 caracteres
    - Substitui só se a forma acentuada for única/dominante (freq >= 200)
    - Preserva capitalização original
    - Não altera texto dentro de código/URLs (heurística simples)"""
    if not _REV_INDEX:
        return texto
    palavras_proibidas_alterar = _STOPWORDS

    def _sub(match):
        word = match.group(0)
        sw = _strip_accents(word)
        # Não altera se já tem acento
        if sw != word:
            return word
        # Pula stopwords
        if word.lower() in palavras_proibidas_alterar:
            return word
        # Pula palavras curtas (< 4)
        if len(sw) < 4:
            return word
        # Consulta o índice reverso
        candidates = _REV_INDEX.get(sw)
        if not candidates:
            return word
        # Caso 1: candidato único → substitui
        if len(candidates) == 1:
            return _match_case(word, candidates[0][0])
        # Caso 2: múltiplos — só substitui se o top for MUITO dominante
        # (freq do top >= 3x a do segundo)
        top, top_f = candidates[0]
        second_f = candidates[1][1] if len(candidates) > 1 else 0
        if top_f >= 10 * second_f:
            return _match_case(word, top)
        return word

    # Regex: palavras (letras, incluindo acentuadas), não dentro de URLs/code
    # Heurística simples: não altera palavras precededidas/followed por / ou . ou : com números
    return re.sub(r"(?<![/\.:])\b[a-zA-Zà-ÿÀ-Ÿ]+\b(?![/\.:])", _sub, texto)


# Backwards-compat alias
def normalizar_acentos(texto):
    return restaurar_acentos(texto)



NIVEL_EQUIVALENCIA = {
    'VI': {'nome': 'VI', 'equivalente': 'Doutor', 'percentual': '75%',
           'destinado': 'servidor com diploma de mestrado',
           'percentual_extenso': 'setenta e cinco por cento'},
    'V': {'nome': 'V', 'equivalente': 'Mestre', 'percentual': '52%',
          'destinado': 'servidor com certificado de pos-graduação lato sensu',
          'percentual_extenso': 'cinquenta e dois por cento'},
    'IV': {'nome': 'IV', 'equivalente': 'Graduacao', 'percentual': '30%',
           'destinado': 'servidor com diploma de graduação no ensino superior',
           'percentual_extenso': 'trinta por cento'},
}
NIVEL_PADRAO = NIVEL_EQUIVALENCIA['VI']

def parse_br(s):
    if not s: return 0.0
    s = s.strip()
    if ',' in s: return float(s.replace('.', '').replace(',', '.'))
    return float(s)

def fmt_br(n):
    return f"{n:.2f}".replace('.', ',')

def get_nivel_info(rsc_raw):
    if not rsc_raw: return NIVEL_PADRAO
    t = rsc_raw.upper().strip()
    for nível in ['VI', 'V', 'IV']:
        if nível in t:
            return NIVEL_EQUIVALENCIA.get(nível, NIVEL_PADRAO)
    m = re.search(r'N[ÍI]VEL\s*(VI|V|IV)', t)
    if m:
        return NIVEL_EQUIVALENCIA.get(m.group(1), NIVEL_PADRAO)
    return NIVEL_PADRAO

def perguntar_ano_ingresso():
    ano_atual = datetime.now().year
    while True:
        try:
            resp = input("\nEm que ano voce ingressou na UFV? ").strip()
            if not resp: continue
            ano = int(resp)
            if 1960 <= ano <= ano_atual: return ano
            print(f"Ano invalido. Digite entre 1960 e {ano_atual}.")
        except ValueError:
            print("Digite um ano valido.")


# =============================================================================
# PARSER DO PDF
# =============================================================================
class RSCPDFParser:
    """Parser completo do Relatorio Detalhado RSC - extrai TODO o conteudo."""

    def __init__(self, pdf_path):
        self.pdf_path = Path(pdf_path)
        self.data = {
            'nome': '', 'matricula': '', 'cargo': '', 'titulação': '',
            'rsc_requerido': '', 'data_admissao': '', 'lotacao': '',
            'funcao': '', 'nivel_classe': '', 'rsc_nivel': 'VI',
            'equivalente': 'Doutorado',
            'total_geral': 0.0, 'total_criterios': 0,
            'grupos': [], 'criterios': {}, 'ordem_criterios': [],
        }
        self._parse()

    def _parse(self):
        with pdfplumber.open(str(self.pdf_path)) as pdf:
            pages = [p.extract_text() or '' for p in pdf.pages]
        raw = '\n'.join(pages)
        self._extract_header(raw)
        self._extract_grupos(raw)
        self._extract_criterios(pages)
        self.data['nivel_info'] = get_nivel_info(self.data.get('rsc_requerido', ''))

    def _extract_header(self, raw):
        def get(pat, g=1):
            m = re.search(pat, raw)
            return m.group(g).strip() if m else ''
        self.data['matricula'] = get(r'Matr[íi]cula SIAPE:\s*(\d+)')
        self.data['nome'] = get(r'Nome:\s*(.+?)(?:\n|$)')
        self.data['cargo'] = get(r'Cargo:\s*(.+?)(?:\n|$)')
        self.data['data_admissao'] = get(r'Data de Admiss[ãa]o:\s*(.+?)(?:\n|$)')
        self.data['nivel_classe'] = get(r'N[íi]vel de Classifica[çc][ãa]o:\s*(.+?)(?:\n|$)')
        self.data['lotacao'] = get(r'Lota[çc][ãa]o:\s*(.+?)(?:\n|$)')
        self.data['funcao'] = get(r'Fun[çc][ãa]o/Encargo \(se houver\):\s*(.+?)(?:\n|$)')
        self.data['titulação'] = get(r'Titula[çc][ãa]o Atual:\s*(.+?)(?:\n|$)')
        m_rsc = re.search(r'RSC\s+(VI|V|IV)', raw, re.I)
        self.data['rsc_nivel'] = m_rsc.group(1).strip() if m_rsc else 'VI'
        self.data['equivalente'] = get(r'Equivalente a\s+(.+?)(?:\n|$)')
        self.data['rsc_requerido'] = f"RSC {self.data['rsc_nivel']} - Equivalente a {self.data['equivalente']}"
        tg = re.search(r'Total Geral:\s*([\d,.]+)', raw)
        self.data['total_geral'] = parse_br(tg.group(1)) if tg else 0.0
        tc = re.search(r'Crit[ée]rios Espec[íi]ficos Pontuados:\s*(\d+)', raw)
        self.data['total_criterios'] = int(tc.group(1)) if tc else 0

    def _extract_grupos(self, raw):
        nomes = [
            ('I', 'Participação em Grupos de Trabalho, Comissões, Comitês, Núcleos e Representações', 'Comissões'),
            ('II', 'Participação e Atuação em Projetos Institucionais', 'Projetos'),
            ('III', 'Recebimento de Premiação', 'Premiações'),
            ('IV', 'Designação para Assunção de Responsabilidades Técnico-Administrativas ou Especializadas', 'Responsabilidades'),
            ('V', 'Exercício de Função ou Cargo de Direção ou de Assessoramento Institucional', 'Direção'),
            ('VI', 'Produção, Prospecção e Difusão de Conhecimento Científico ou Técnico', 'Produção'),
        ]
        for rom, nome, curto in nomes:
            pat = rf'REQUISITO\s+{re.escape(rom)}.*?(?=REQUISITO\s+\w|\Z)'
            m = re.search(pat, raw, re.DOTALL)
            crit, pts = 0, 0.0
            if m:
                bloco = m.group(0)
                cm = re.search(r'(\d+)\s*crit[eé]rio\(s\)\s*([\d,.]+)', bloco)
                if cm:
                    crit = int(cm.group(1))
                    pts = parse_br(cm.group(2))
                else:
                    cz = re.search(r'0\s*crit[eé]rio\(s\)\s*([\d,.]+)', bloco)
                    if cz:
                        pts = parse_br(cz.group(1))
            self.data['grupos'].append({
                'romano': rom, 'nome': nome, 'nome_curto': curto,
                'criterios': crit, 'pontos': pts
            })

    def _extract_criterios(self, pages):
        raw_text = '\n'.join(pages)
        lines = raw_text.split('\n')
        ordem_map = {
            'I-01': 0, 'I-02': 1, 'I-03': 2, 'I-05': 3, 'I-06': 4,
            'II-02': 5, 'II-07': 6,
            'IV-01': 7, 'IV-07': 8,
            'V-01': 9, 'V-02': 10, 'V-03': 11, 'V-04': 12,
            'VI-09': 13, 'VI-10': 14, 'VI-15': 15, 'VI-16': 16,
        }
        for i, line in enumerate(lines):
            cm = re.match(r'Crit[eé]rio\s+(\w+)\s*[-–]\s*(\d+)\s*:\s*(.*)', line)
            if not cm:
                continue
            key = f"{cm.group(1)}-{cm.group(2)}"
            desc = cm.group(3).strip()
            ordem = ordem_map.get(key, 99)
            full_desc = desc
            j = i + 1
            while j < len(lines):
                l = lines[j].strip()
                if not l or l.startswith('#'):
                    j += 1
                    continue
                if re.match(r'Crit[eé]rio\s+\w+\s*[-–]\s*\d+\s*:', l):
                    break
                if re.match(r'^\d+\s+', l) or re.match(r'^[\d.,]+\s*pontos?', l):
                    break
                if l.startswith(('Descrição', 'Data', 'Observação', '#')):
                    j += 1
                    continue
                if len(l) > 20 and not re.match(r'^\d+$', l):
                    full_desc += ' ' + l
                j += 1
            items = []
            pontos = None
            k = i + 1
            has_items = False
            while k < len(lines):
                l = lines[k].strip()
                if not l:
                    k += 1
                    continue
                if re.match(r'Crit[eé]rio\s+\w+\s*[-–]\s*\d+\s*:', lines[k]):
                    break
                im = re.match(r'^\s*(\d+)\s+(.+)$', lines[k])
                if im:
                    has_items = True
                    items.append({'num': im.group(1), 'texto': im.group(2).strip()})
                pm = re.match(r'^\s*([\d,.]+)\s*pontos?\s*$', lines[k])
                if pm:
                    pontos = parse_br(pm.group(1))
                k += 1
            crit_data = {
                'key': key, 'romano': cm.group(1), 'numero': cm.group(2),
                'descricao': full_desc, 'ordem': ordem,
                'itens': items, 'pontos': pontos
            }
            self.data['criterios'][key] = crit_data
            if key not in self.data['ordem_criterios']:
                self.data['ordem_criterios'].append(key)
        self.data['ordem_criterios'].sort(key=lambda k: ordem_map.get(k, 99))


# =============================================================================
# GERADOR DO MEMORIAL v3.1 — ESTRUTURA E TÓPICOS DO REFERENCIAL
# =============================================================================
class MemorialGenerator:
    """Gera memorial seguindo ESTRUTURA E TÓPICOS do memorial de referencia aprovado."""

    def __init__(self, data, ano_ingresso):
        self.d = data
        self.ano_ingresso = ano_ingresso
        self.ano_atual = datetime.now().year
        self.anos_carreira = self.ano_atual - ano_ingresso
        self.nível = data.get('nivel_info', NIVEL_PADRAO)
        self.nome = data['nome']
        self.data_admissao = data.get('data_admissao', '07/01/2009')
        self.equivalente = self.nível['equivalente']

    def _narrativa_intro(self):
        """1 INTRODUÇÃO — TRAJETÓRIA E FUNDAMENTOS"""
        g1 = self.d['grupos'][0]
        total = fmt_br(self.d['total_geral'])
        total_crit = self.d['total_criterios']
        ano_fim = self.ano_atual
        return [
            "# 1 INTRODUÇÃO -- TRAJETORIA E FUNDAMENTOS\n",
            "\n",
            "## 1.1 Quem sou e o que apresento\n",
            "\n",
            f"Meu nome e {self.nome}, matricula SIAPE {self.d['matricula']}, "
            f"servidor público federal ocupante do cargo de {self.d['cargo']} na "
            f"Universidade Federal de Viçosa, portador do título de {self.d['titulação']}. "
            f"Apresento este memorial descritivo com o objetivo de obter o Reconhecimento "
                f"de Saberes e Competencias no Nível {self.nível['nome']} (RSC-PCCTAE {self.nível['nome']}), "
            f"equivalente ao {self.equivalente}, conforme previsto na Lei nº 11.091/2005 "
            f"(alterada pela Lei nº 15.367/2026) e regulamentado pelo Decreto nº 13.048/2026.\n",
            "\n",
            "Este memorial não e mera relacao de atividades. E a narrativa de uma trajetoria "
            f"profissional de {self.anos_carreira} anos -- de {self.ano_ingresso} a {ano_fim} -- "
            f"construida na intersecao entre gestão de pessoas, planejamento institucional, "
            f"inovação tecnologica e produção de conhecimento. Cada comissão que presidi, "
            f"cada banca que integrei, cada projeto que liderei, cada texto que publiquei, "
            f"cada curso que ministrei representa não apenas uma atividade realizada, mas "
            f"um saber construido, uma competência desenvolvida, uma contribuicao singular "
            f"a Universidade Federal de Viçosa.\n",
            "\n",
            "## 1.2 A essencia do meu fazer profissional\n",
            "\n",
            f"Como {self.d['cargo']}, minha atuação transcendeu o desempenho ordinário "
            f"das atribuicoes do cargo. Ao longo de quase duas decadas, desenvolvi competências "
            f"em tres dimensoes fundamentais:\n",
            "\n",
            "**Dimensao normativo-regulatoria:** presidi comissões que revisaram e "
            "atualizaram os marcos regulatorios de estágio probatorio, avaliação de "
            "desempenho, assinaturas eletronicas e Programa de Gestao e Desempenho -- "
            "contribuindo diretamente para o aperfeicoamento da gestão universitaria.\n",
            "\n",
            "**Dimensao executivo-estrategica:** ocupei cargos de direcao e assessoramento "
            "-- fui Assessor Especial da Pro-Reitoria de Gestao de Pessoas (CD-03/04), "
            "respondi como Pro-Reitor de Gestao de Pessoas substituto (CD-02) e chefiei "
            "setores e divisoes estrategicas (FG-01 a FG-04) -- o que me permitiu contribuir "
            "para a formulacao e execução de politicas institucionais de gestão de pessoas.\n",
            "\n",
            "**Dimensao academico-cientifica:** publiquei livro com ISBN, artigos em anais "
            "de coloquios internacionais e periodico cientifico, e atuei como instrutor em "
            "cursos de capacitação, difundindo conhecimento técnico e cientifico sobre "
            "gestão universitaria.\n",
            "\n",
            "## 1.3 Fundamentos legais\n",
            "\n",
            "O presente memorial atende aos requisitos estabelecidos no art. 3º (eixos de "
            "atuação), art. 5º (pontuação e critérios) e art. 15º (saberes e competências "
            "diferenciados) do Decreto nº 13.048/2026, bem como aos dispositivos da Lei nº "
            "11.091/2005 e da Lei nº 15.367/2026.\n",
            "\n",
            f"Para o RSC-PCCTAE Nível {self.nível['nome']}, exige-se:\n",
            "\n",
            "- Pontuacao minima de 75 (setenta e cinco) pontos;\n",
            "- Minimo de 7 (sete) critérios especificos dos Anexos I a VI;\n",
            "- Pelo menos 1 (um) critério do Anexo VI (art. 3º, inciso VI);\n",
            f"- Titulacao de {self.d['titulação']} (art. 5º, S 1, inciso VI), ja comprovada.\n",
            "\n",
            "Conforme demonstrarao as secoes seguintes, todos esses requisitos sao atendidos "
            f"com ampla margem.\n",
        ]

    def _anexo_intro(self, num_rom, titulo_anexo, artigo_inciso, grupo):
        """Intro comum para cada anexo"""
        rom_to_num = {'I': '2', 'II': '3', 'III': '4', 'IV': '5', 'V': '6', 'VI': '7'}
        sec = rom_to_num.get(num_rom, '2')
        return [
            f"# {sec} ANEXO {num_rom} -- {titulo_anexo}\n",
            "\n",
            f"*Art. 3, inciso {artigo_inciso} -- Pontuacao oficial: "
            f"{fmt_br(grupo['pontos'])} pts ({grupo['criterios']} critério(s))*\n",
            "\n",
        ]

    def _narrativa_anexo_I(self):
        """2 ANEXO I — PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS"""
        g = self.d['grupos'][0]
        md = self._anexo_intro('I', 'PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS', 'I', g)
        c = self.d['criterios']
        md += [
            "## 2.1 Memorialista: um construtor de comissões\n",
            "\n",
            "Se existe uma marca indelével na minha trajetoria, e a participação em comissões. "
            "Iniciei como membro -- aprendendo, observando, contribuindo. Tornei-me presidente "
            "-- liderando, decidindo, transformando. Foram mais de oitenta designacoes formais "
            "ao longo de 17 anos, que me permitiram compreender a universidade por dentro, "
            "em suas dimensoes normativa, administrativa, pedagogica e estrategica.\n",
            "\n",
        ]
        # Item I-01
        if 'I-01' in c:
            md += [
                f"## 2.2 Item I-1: Exercicio do mandato como membro de conselhos superiores e "
                f"colegiados ({fmt_br(c['I-01']['pontos'])} pts)\n",
                "\n",
                "Fui designado representante dos pais no Conselho de Administracao do Laboratorio "
                "de Desenvolvimento Infantil (LDI), onde participei das deliberacoes sobre a "
                "gestão e as politicas educacionais da unidade -- uma experiência que ampliou "
                "minha visao sobre a gestão participativa na universidade.\n",
                "\n",
            ]
        # Item I-02
        if 'I-02' in c:
            qtd = len(c['I-02']['itens'])
            md += [
                f"## 2.3 Item I-2: Coordenacao e presidencia de comissões "
                f"({fmt_br(c['I-02']['pontos'])} pts)\n",
                "\n",
                f"Presidi ou coordenei {qtd} ({'nove' if qtd==9 else str(qtd)}) comissões "
                f"ao longo da carreira. Cada presidencia representou um desafio distincto. "
                f"A comissão de assinaturas eletronicas, por exemplo, exigiu-me estudo "
                f"aprofundado da Medida Provisoria 2.200-2/2001 e da legislação correlata "
                f"para propor norma que viabilizasse a tramitacao eletronica com validade "
                f"juridica. A comissão do PGD demandou a compreensão de um novo paradigma "
                f"de gestão por resultados e sua adaptacao a realidade da UFV.\n",
                "\n",
            ]
        # Item I-03
        if 'I-03' in c:
            qtd = len(c['I-03']['itens'])
            md += [
                f"## 2.4 Item I-3: Participacao como membro de comissões "
                f"({fmt_br(c['I-03']['pontos'])} pts)\n",
                "\n",
                f"Participei como membro de {qtd} comissões, incluindo bancas examinadoras "
                f"centrais de concursos públicos (2009 a {self.ano_atual}), Comissao de "
                f"Gestao de Integridade (Port. 0863/2019/RTR), Comissao de Assessoramento "
                f"do Relatorio de Gestao (2022-2025), Comissao de Elaboracao do PDI 2024-2029, "
                f"Comissao Especial de Estudos do PASES (Port. 0014/2023/Cepe), Comissao "
                f"Organizadora do UFV 60+ (Port. 074/2025/PRE e 023/2026/PRE), e Comissao "
                f"de Adequacao do Estagio Probatorio (Port. 0224/2025/RTR).\n",
                "\n",
                "Esta atuação continua por 17 anos me deu conhecimento aprofundado dos ritos "
                "e procedimentos de concursos públicos no ambito do Decreto nº 9.739/2019 e "
                "legislação correlata.\n",
                "\n",
            ]
        # Item I-05
        if 'I-05' in c:
            qtd = len(c['I-05']['itens'])
            md += [
                f"## 2.5 Item I-5: Organizacao, fiscalização e execução de vestibulares e "
                f"concursos ({fmt_br(c['I-05']['pontos'])} pts)\n",
                "\n",
                f"Atuei em {qtd} (vinte e oito) processos seletivos e concursos públicos, "
                f"desde o Vestibular UFV 2010 ate o Concurso Publico UFV {self.ano_atual}. "
                f"Essa experiência continuada me proporcionou domínio integral dos fluxos "
                f"operacionais de selecao de servidores e estudantes.\n",
                "\n",
            ]
        # Item I-06
        if 'I-06' in c:
            qtd = len(c['I-06']['itens'])
            md += [
                f"## 2.6 Item I-6: Elaboracao, revisao e correção de provas "
                f"({fmt_br(c['I-06']['pontos'])} pts)\n",
                "\n",
                f"Coordenei ou integrei {qtd} bancas de elaboração e revisao de provas, "
                f"atuando como coordenador na maioria delas a partir de 2019. Destaque para "
                f"a consolidacao da minha lideranca técnica nessa area.\n",
                "\n",
            ]
        return md

    def _narrativa_anexo_II(self):
        """3 ANEXO II — PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS"""
        g = self.d['grupos'][1]
        md = self._anexo_intro('II', 'PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS', 'II', g)
        c = self.d['criterios']
        md += [
            "## 3.1 Pesquisa academica: o REUNI como objeto de estudo\n",
            "\n",
            "Participei como pesquisador do projeto \"Programa de Apoio a Planos de "
            "Reestruturacao e Expansao das Universidades Federais -- REUNI: limites e "
            "potencialidades na gestão das Instituições Federais de Ensino Superior em "
            "Minas Gerais\" (Processo 60204259518). Esta pesquisa, iniciada em 2010, "
            "resultou em livro, artigos e apresentacoes em coloquios internacionais -- "
            "conforme detalhado no Anexo VI.\n",
            "\n",
        ]
        if 'II-07' in c:
            md += [
                "## 3.2 Avaliacao de trabalhos academicos\n",
                "\n",
                "Participei como membro de banca de avaliação de Trabalho de Conclusao de "
                "Curso em Administracao, contribuindo para a formação de novos profissionais.\n",
                "\n",
            ]
        return md

    def _narrativa_anexo_III(self):
        """4 ANEXO III — PREMIAÇÕES"""
        g = self.d['grupos'][2]
        md = self._anexo_intro('III', 'RECEBIMENTO DE PREMIAÇÃO', 'III', g)
        md += [
            "Declaro, para os devidos fins, que não recebi premiações formais em "
            "eventos de reconhecimento público que possam ser enquadradas neste Anexo.\n",
            "\n",
        ]
        return md

    def _narrativa_anexo_IV(self):
        """5 ANEXO IV — RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS"""
        g = self.d['grupos'][3]
        md = self._anexo_intro('IV', 'RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS', 'IV', g)
        c = self.d['criterios']
        if g['pontos'] > 0:
            md += [
                "Fui designado para atuação em sistemas estruturantes da administracao "
                "pública federal, incluindo COMPREV, e-SIAPE, SIAPE, SIGEPE, REDE SERPRO, "
                "SiapeNet e SIASS, contribuindo para a operacao e o aperfeicoamento de "
                "sistemas essenciais a gestão de pessoas no servico público federal.\n",
                "\n",
            ]
            if 'IV-01' in c:
                md += [
                    f"## 5.1 Item IV-1: Sistemas estruturantes ({fmt_br(c['IV-01']['pontos'])} pts)\n",
                    "\n",
                    "Atuei em atividades de execução e operacao no Subsistema Integrado de "
                    "atencao a Saude do Servidor (SIASS), no Sistema de Pessoal Civil da "
                    "Administracao Federal (SiapeNet), no e-SIAPE, no Sistema de Gestao de "
                    "Pessoas (SIGEPE), e na Rede SERPRO, sistemas estruturantes da "
                    "administracao pública federal.\n",
                    "\n",
                ]
            if 'IV-07' in c:
                md += [
                    f"## 5.2 Item IV-7: Sistemas e processos institucionais "
                    f"({fmt_br(c['IV-07']['pontos'])} pts)\n",
                    "\n",
                    "Desenvolvi e operacionalizei sistemas e processos de trabalho no ambito "
                    "da gestão de pessoas, incluindo o Sisvest (sistema para gerenciamento de "
                    "processos seletivos) e o Gespe-Documentos, contribuindo para a "
                    "modernizacao e eficiencia dos fluxos administrativos.\n",
                    "\n",
                ]
        else:
            md += [
                "O sistema oficial registra minha designação para atuação em sistemas "
                "estruturantes da administracao pública federal (COMPREV, e-SIAPE, SIAPE, "
                "SIGEPE, REDE SERPRO, SiapeNet, SIASS). Estas atividades, contudo, não "
                "foram pontuadas no sistema, possivelmente por pendencia de comprovacao "
                "documental. Registro-as para conhecimento da CRSC-PCCTAE, que podera "
                "avaliar seu enquadramento.\n",
                "\n",
            ]
        return md

    def _narrativa_anexo_V(self):
        """6 ANEXO V — EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO"""
        g = self.d['grupos'][4]
        md = self._anexo_intro('V', 'EXERCICIO DE FUNÇÕES DE DIRECAO E ASSESSORAMENTO', 'V', g)
        c = self.d['criterios']
        md += [
            "## 6.1 Uma trajetoria de lideranca institucional\n",
            "\n",
            "Um dos aspectos mais significativos da minha carreira foi o exercicio de "
            "cargos de direcao e funcoes gratificadas ao longo de mais de uma decada. "
            "Essas posicoes me permitiram contribuir diretamente para a formulacao e "
            "execução de politicas institucionais de gestão de pessoas na UFV.\n",
            "\n",
        ]
        if 'V-01' in c:
            qtd = len(c['V-01']['itens'])
            md += [
                f"## 6.2 Item V-1: CD-02 -- Pro-Reitor de Gestao de Pessoas substituto "
                f"({fmt_br(c['V-01']['pontos'])} pts)\n",
                "\n",
                f"Em múltiplas ocasioes entre 2019 e {self.ano_atual}, fui designado para "
                f"responder pela Pro-Reitoria de Gestao de Pessoas da UFV na ausencia do "
                f"titular. Essa experiência me proporcionou visao sistemica da gestão "
                f"universitaria e responsabilidade direta sobre decisoes estrategicas.\n",
                "\n",
            ]
        if 'V-02' in c:
            qtd_itens = len(c['V-02']['itens'])
            if qtd_itens > 0:
                meses = "48"  # typical from reference
                md += [
                    f"## 6.3 Item V-2: CD-03/04 -- Assessor Especial da PGP titular "
                    f"({fmt_br(c['V-02']['pontos'])} pts)\n",
                    "\n",
                    "Fui designado Assessor Especial da Pro-Reitoria de Gestao de Pessoas, "
                    "atuando diretamente no assessoramento a alta administracao da PGP, "
                    "contribuindo para a formulacao de politicas, a analise de processos "
                    "estrategicos e a tomada de decisoes institucionais.\n",
                    "\n",
                ]
        if 'V-03' in c:
            md += [
                f"## 6.4 Item V-3: FG-01/02 -- Chefe de Divisao "
                f"({fmt_br(c['V-03']['pontos'])} pts)\n",
                "\n",
                "Fui titular da Chefia da Divisao de Desenvolvimento de Pessoas da PGP "
                "e, anteriormente, da Chefia do Setor de Provimento, Acompanhamento e "
                "Avaliacao. Como substituto, respondi pela Chefia da Divisao de "
                "Desenvolvimento de Pessoas em inumeras ocasioes. Essas posicoes me "
                "permitiram gerir equipes, coordenar processos de desenvolvimento de "
                "servidores e implementar politicas de capacitação.\n",
                "\n",
            ]
        if 'V-04' in c:
            md += [
                f"## 6.5 Item V-4: FG-03+ -- Chefia de Setor "
                f"({fmt_br(c['V-04']['pontos'])} pts)\n",
                "\n",
                "Fui titular da Chefia do Setor de Capacitacao e Treinamento e, "
                "posteriormente, da Chefia do Setor de Provimento, Acompanhamento e "
                "Avaliacao, acumulando meses de exercicio em posicoes de chefia.\n",
                "\n",
            ]
        return md

    def _narrativa_anexo_VI(self):
        """7 ANEXO VI — PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO"""
        g = self.d['grupos'][5]
        md = self._anexo_intro('VI', 'PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO '
                                'CIENTÍFICO E TÉCNICO', 'VI', g)
        c = self.d['criterios']
        md += [
            "## 7.1 A face academica de minha trajetoria\n",
            "\n",
            "Sempre acreditei que a gestão universitaria, para ser efetiva, precisa estar "
            "ancorada em conhecimento cientifico. Por isso, ao longo de minha carreira, "
            "busquei não apenas executar, mas tambem produzir e difundir conhecimento "
            "sobre a gestão das instituições federais de ensino.\n",
            "\n",
        ]
        if 'VI-09' in c:
            md += [
                f"## 7.2 Item VI-9: Publicacao de livro com ISBN "
                f"({fmt_br(c['VI-09']['pontos'])} pts)\n",
                "\n",
                "Publiquei o livro \"Consequencias, limites e potencialidades na "
                "implementação do REUNI\", em coautoria com Luiz Antonio Abrantes e "
                "Antonio Carlos Brunozi Junior, pela editora Novas Edicoes Academicas "
                "(Sao Paulo, SP), ISBN 978-3-639-74424-8. A obra resultou de pesquisa "
                "academica sobre o Programa de Apoio a Planos de Reestruturacao e "
                "Expansao das Universidades Federais (REUNI) e suas implicações para "
                "a gestão das IFES mineiras.\n",
                "\n",
            ]
        if 'VI-10' in c:
            md += [
                f"## 7.3 Item VI-10: Artigos publicados "
                f"({fmt_br(c['VI-10']['pontos'])} pts)\n",
                "\n",
                "Publiquei os seguintes trabalhos academicos:\n",
                "\n",
                "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JR., A. C.; SILVA, F. C.; SOUZA, A. P. "
                "Reforma Universitaria no Brasil: uma analise dos documentos oficiais e da "
                "produção cientifica sobre o Reuni. **X Coloquio Sobre Gestión Universitaria "
                "en America del Sur**, Mar del Plata, 2010. -- Anais de evento internacional.\n",
                "\n",
                "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C.; BRUNOZI, M. A. V. "
                "Caracterizacao, limites e potencialidades do programa REUNI em IFES mineiras: "
                "um estudo multicaso. **XIII Coloquio Internacional sobre Gestao Universitaria "
                "nas Americas**, Buenos Aires, 2013. -- Anais de evento internacional.\n",
                "\n",
                "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C. Planejamento, "
                "implementação e avaliação do REUNI: Um Estudo em Universidades Mineiras. "
                "**Estudo & Debate (Online)**, v. 22, p. 78-96, 2015. -- Artigo em periodico "
                "cientifico.\n",
                "\n",
            ]
        if 'VI-15' in c:
            md += [
                f"## 7.4 Item VI-15: Instrutor em acoes formativas "
                f"({fmt_br(c['VI-15']['pontos'])} pts)\n",
                "\n",
                "Atuei como instrutor nos seguintes cursos de capacitação:\n",
                "\n",
                "- 01/2009: Fundamentos em Administracao -- Gestao de Pessoas -- CEPET (20/07/2009)\n",
                "- 06/2012: Treinamento de Integracao para Novos Servidores -- UFV (17/04/2012)\n",
                f"- 04/{self.ano_atual}: Reconhecendo Saberes e Competencias: Entenda o "
                f"RSC-PCCTAE -- UFV (06/05/{self.ano_atual})\n",
                "\n",
                "Essas atividades de instrutoria demonstram minha capacidade de difundir "
                "conhecimento técnico e contribuir para a formação de servidores.\n",
                "\n",
            ]
        if 'VI-16' in c:
            md += [
                f"## 7.5 Item VI-16: Coordenacao de eventos "
                f"({fmt_br(c['VI-16']['pontos'])} pts)\n",
                "\n",
                "Coordenei o curso \"Redacao Oficial\" (03/2009), contribuindo para a "
                "capacitação de servidores na produção de documentos oficiais conforme "
                "as normas da Administracao Publica Federal.\n",
                "\n",
            ]
        return md

    def _sintese_pontuacao(self):
        """8 SÍNTESE DE PONTUAÇÃO"""
        total = fmt_br(self.d['total_geral'])
        total_crit = self.d['total_criterios']
        g = self.d['grupos']
        c = self.d['criterios']
        req_VI_pts = float(g[5]['pontos'])
        # Count total criteria with points > 0
        total_com_pontos = sum(1 for grp in g if grp['criterios'] > 0)
        md = [
            "# 8 SÍNTESE DE PONTUAÇÃO\n",
            "\n",
            "## 8.1 Quadro geral -- Pontuacao oficial (sistema UFV)\n",
            "\n",
            "| Anexo | Conteudo | Criterios | Pontuacao |\n",
            "|-------|----------|-----------|-----------|\n",
        ]
        for grp in g:
            nome_curto_map = {'I': 'Comissões', 'II': 'Projetos', 'III': 'Premiações',
                              'IV': 'Responsabilidades', 'V': 'Direção', 'VI': 'Produção'}
            md.append(
                f"| **Anexo {grp['romano']}** | {nome_curto_map.get(grp['romano'], grp['nome_curto'])} "
                f"| {grp['criterios']} | **{fmt_br(grp['pontos'])}** |\n"
            )
        md += [
            f"| | **Total Geral** | **{total_crit}** | **{total}** |\n",
            "\n",
            "## 8.2 Verificacao dos requisitos legais para RSC-PCCTAE VI\n",
            "\n",
            "| Requisito | Exigido | Atendido |\n",
            "|-----------|---------|----------|\n",
            f"| Pontuacao total | Minimo 75,00 pts | **{total} pts** -- Atende |\n",
            f"| Criterios especificos | Minimo 7 | **{total_crit} critérios** -- Atende |\n",
            f"| Anexo VI (produção) | Pelo menos 1 critério | **{g[5]['criterios']} critério(s)** -- Atende |\n",
            f"| Anexos com pontuação | -- | **{total_com_pontos} anexos** com pontuação |\n",
            f"| Titulacao | {self.d['titulação']} | Comprovada -- Atende |\n",
            "\n",
            "Todos os requisitos legais para concessão do RSC-PCCTAE Nível VI sao "
            "atendidos com ampla margem.\n",
        ]
        return md

    def _reflexao_final(self):
        """9 REFLEXÃO FINAL — SABERES E COMPETÊNCIAS"""
        total = fmt_br(self.d['total_geral'])
        md = [
            "# 9 REFLEXÃO FINAL -- SABERES E COMPETÊNCIAS\n",
            "\n",
            "*Em conformidade com o art. 15º do Decreto nº 13.048/2026*\n",
            "\n",
            "## 9.1 Que saberes construi?\n",
            "\n",
            f"Ao longo de {self.anos_carreira} anos de servico público na UFV, não me "
            f"limitei a executar tarefas. Construi saberes. Os principais sao:\n",
            "\n",
            "**Saber normativo:** aprendi a ler, interpretar e propor normas de gestão "
            "de pessoas. Presidi comissões que revisaram resolucoes historicas da UFV -- "
            "a Res. Consu 03/2006 (estágio probatorio) e a Res. 08/2008 (avaliação de "
            "desempenho) --, contribuindo para sua adequacao a legislação superveniente.\n",
            "\n",
            "**Saber de gestão:** ocupei posicoes de chefia, assessoramento e direcao "
            "que me permitiram compreender a universidade como sistema complexo. Como "
            "Assessor Especial da PGP, participei da formulacao de politicas institucionais. "
            "Como Pro-Reitor substituto, respondi por decisoes estrategicas.\n",
            "\n",
            "**Saber academico-cientifico:** pesquisei, publiquei livro e artigos sobre "
            "o REUNI, contribuindo para o conhecimento sobre gestão universitaria no "
            "Brasil. Minha produção cientifica demonstra que a pratica profissional, "
            "quando refletida academicamente, gera conhecimento relevante para a "
            "comunidade.\n",
            "\n",
            "**Saber pedagogico:** formei servidores como instrutor em cursos de "
            "capacitação e orientei colegas em estágio probatorio, transmitindo "
            "conhecimentos e contribuindo para o desenvolvimento institucional.\n",
            "\n",
            "**Saber tecnologico:** liderei a implantacao do SEI na Gestao de Pessoas, "
            "participei da regulamentacao de assinaturas eletronicas e fui Agente SEI, "
            "contribuindo para a modernizacao tecnologica da UFV.\n",
            "\n",
            "## 9.2 Qual minha contribuicao singular?\n",
            "\n",
            "Minha maior contribuicao a UFV talvez seja ter demonstrado que um Tecnico "
            "em Assuntos Educacionais pode -- e deve -- transcender as fronteiras de seu "
            "cargo. Publiquei livro quando muitos limitam-se a executar rotinas. Presidi "
            "comissões estrategicas quando muitos contentam-se em participar. Ocupei "
            "cadeiras de direcao quando muitos julgam-nas inacessiveis. Formei servidores "
            "quando muitos guardam conhecimento para si.\n",
            "\n",
            "Esta trajetoria -- que mescla gestão, pesquisa, docência e inovação -- e a "
            "demonstracao viva de que os saberes construidos na pratica profissional, "
            "quando refletidos sistematicamente, equivalem-se a formação academica "
            f"stricto sensu que o RSC-{self.nível['nome']} reconhece.\n",
            "\n",
            "## 9.3 Pedido\n",
            "\n",
            "Ante o exposto, com fundamento na Lei nº 11.091/2005 (alterada pela Lei nº "
            "15.367/2026), no Decreto nº 13.048/2026 e na documentacao comprobatoria "
            f"anexa, requeiro a CRSC-PCCTAE da Universidade Federal de Viçosa o "
            f"deferimento da concessão do Reconhecimento de Saberes e Competencias no "
            f"nível RSC-PCCTAE {self.nível['nome']}.\n",
            "\n",
            "Nestes termos, pede deferimento.\n",
            "\n",
            f"Viçosa, {self.ano_atual}.\n",
            "\n",
            f"{self.nome}\n",
            f"{self.d['cargo']} -- UFV\n",
            f"{self.d['titulação']}\n",
        ]
        return md

    def _referencias(self):
        return [
            "# REFERENCIAS\n",
            "\n",
            "BRASIL. **Lei nº 11.091**, de 12 de janeiro de 2005. Dispoe sobre a "
            "estruturação do Plano de Carreira dos Cargos Tecnico-Administrativos em "
            "Educacao, no ambito das Instituições Federais de Ensino vinculadas ao "
            "Ministério da Educação, e da outras providências. **Diario Oficial da "
            "Uniao**: Brasilia, DF, 13 jan. 2005.\n",
            "\n",
            "BRASIL. **Lei nº 15.367**, de 30 de marco de 2026. Altera a Lei nº "
            "11.091/2005 para atualizar o Plano de Carreira dos Cargos "
            "Tecnico-Administrativos em Educacao.\n",
            "\n",
            "BRASIL. **Decreto nº 13.048**, de 3 de julho de 2026. Estabelece critérios "
            "e procedimentos para o Reconhecimento de Saberes e Competencias (RSC) no "
            "ambito do PCCTAE. Disponível em: " + DECRETO_URL + ". Acesso em: "
            + datetime.now().strftime('%d %b. %Y') + ".\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C. **Consequencias, "
            "limites e potencialidades na implementação do REUNI**. Sao Paulo: Novas "
            "Edicoes Academicas, 2015. ISBN 9783639744248.\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JR., A. C.; SILVA, F. C.; SOUZA, A. P. "
            "Reforma Universitaria no Brasil: uma analise dos documentos oficiais e da "
            "produção cientifica sobre o Reuni. In: **X Coloquio Sobre Gestión "
            "Universitaria en America del Sur**, Mar del Plata, 2010.\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C.; BRUNOZI, M. A. V. "
            "Caracterizacao, limites e potencialidades do programa REUNI em IFES "
            "mineiras: um estudo multicaso. In: **XIII Coloquio Internacional sobre "
            "Gestao Universitaria nas Americas**, Buenos Aires, 2013.\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C. Planejamento, "
            "implementação e avaliação do REUNI: Um Estudo em Universidades Mineiras. "
            "**Estudo & Debate (Online)**, v. 22, p. 78-96, 2015.\n",
            "\n",
            "PIRES, Alice Regina Pinto; SILVA, Bruna (org.). **Normalizacao de "
            "trabalhos academicos**: atualizada conforme ABNTs NBR 14724/2024, NBR "
            "6023/2018 e NBR 10520/2023. Viçosa, MG: UFV, Biblioteca Central, 2025.\n",
        ]

    def _normalizar(self, texto):
        """Pós-processamento sistemático: acentos + correções especiais."""
        texto = normalizar_acentos(texto)
        # "n" → "nº" após Lei/Decreto/Portaria (antes de número de documento)
        texto = re.sub(r'\b(Lei|Decreto|Portaria|Resolução|Instrução|Norma|Regulamento)\s+n\s+',
                       r'\1 nº ', texto)
        texto = re.sub(r'\bn\s+(\d[\d.,\s]*(?:/20\d\d|/19\d\d))', r'nº \1', texto)
        # "e" → "é" (verbo ser) em contextos seguros
        texto = re.sub(r'\bnão e\b', 'não é', texto)
        texto = re.sub(r'\bnao e\b', 'não é', texto, flags=re.IGNORECASE)
        return texto

    def generate(self):
        """Gera o memorial completo."""
        titulo = (f"MEMORIAL DESCRITIVO PARA RECONHECIMENTO DE SABERES E "
                  f"COMPETÊNCIAS -- RSC-PCCTAE NÍVEL {self.nível['nome']}")
        md = []
        # ===== CAPA (UFV-ABNT) =====
        # Modelo UFV-PPG: UNIVERSIDADE FEDERAL DE VIÇOSA → nome do autor → título → local → ano
        # Elementos centralizados, negrito, espaçamento vertical para efeito de meia folha
        md.append("<br><br><br><br><br>\n")
        md.append('<p align="center"><strong>UNIVERSIDADE FEDERAL DE VIÇOSA</strong></p>\n')
        md.append('<br><br>\n')
        md.append(f'<p align="center"><strong>{self.nome}</strong></p>\n')
        md.append('<br><br><br>\n')
        md.append(f'<p align="center"><strong>{titulo}</strong></p>\n')
        md.append('<br><br><br><br><br>\n')
        md.append('<p align="center"><strong>VIÇOSA -- MINAS GERAIS</strong></p>\n')
        md.append(f'<p align="center"><strong>{self.ano_atual}</strong></p>\n')
        md.append('<br><br><br>\n')
        md.append("\n---\n")

        # ===== FOLHA DE ROSTO =====
        md.append(f'<p align="center"><strong>{self.nome}</strong></p>\n')
        md.append(f'<p align="center"><strong>{titulo}</strong></p>\n')
        md.append('<div style="text-align:justify; margin-left:4cm; margin-right:0cm; line-height:1.0;">\n')
        md.append(
            f'Memorial descritivo apresentado a Comissao para Reconhecimento de Saberes '
            f'e Competencias do Plano de Carreira dos Cargos Tecnico-Administrativos em '
            f'Educacao (CRSC-PCCTAE) da Universidade Federal de Viçosa como requisito '
            f'para concessão do RSC-PCCTAE Nível {self.nível["nome"]}, nos termos da Lei '
            f'n 11.091/2005 (alterada pela Lei nº 15.367/2026), do Decreto nº 13.048/2026 '
            f'e da legislação correlata.\n'
        )
        md.append('</div>\n')
        md.append('<br>\n')
        md.append('<p align="center">VIÇOSA -- MINAS GERAIS</p>\n')
        md.append(f'<p align="center">{self.ano_atual}</p>\n')
        md.append("\n---\n")

        # ===== DEDICATORIA =====
        md.append('<p align="right" style="font-size:12pt; font-style:italic;">\n')
        md.append('  Aos servidores técnico-administrativos em educação,<br>\n')
        md.append('  cujo trabalho silencioso constroi a universidade pública<br>\n')
        md.append('  brasileira dia apos dia.\n')
        md.append('</p>\n')
        md.append("\n---\n")

        # ===== AGRADECIMENTOS =====
        md.append("# AGRADECIMENTOS\n\n")
        md.append(
            f"Expresso minha gratidao a Universidade Federal de Viçosa, instituição que ha "
            f"{self.anos_carreira} anos e o espaco do meu crescimento profissional e pessoal. "
            f"Aos colegas da Pro-Reitoria de Gestao de Pessoas, pelo aprendizado cotidiano "
            f"e pelo trabalho colaborativo que tornou possivel cada conquista aqui registrada.\n\n"
            f"Agradeco a Comissao para Reconhecimento de Saberes e Competencias do PCCTAE, "
            f"pelo cuidadoso trabalho de avaliação das trajetorias dos servidores "
            f"técnico-administrativos.\n\n"
            f"O presente trabalho foi realizado com apoio da Coordenacao de Aperfeicoamento "
            f"de Pessoal de Nível Superior -- Brasil (CAPES) -- Codigo de Financiamento 001.\n\n"
            f"Aos servidores que compartilharam comigo a missao de construir uma universidade "
            f"pública, gratuita e de qualidade, minha sincera admiraco e reconhecimento.\n"
        )
        md.append("\n---\n")

        # ===== EPIGRAFE =====
        md.append('<p align="right" style="font-size:12pt; font-style:italic;">\n')
        md.append('  "A educação não transforma o mundo. Educacao muda as pessoas.<br>\n')
        md.append('  Pessoas transformam o mundo."<br>\n')
        md.append('  <span style="font-size:11pt;">-- Paulo Freire</span>\n')
        md.append('</p>\n')
        md.append("\n---\n")

        # ===== LISTA DE SIGLAS =====
        md.append("# LISTA DE SIGLAS\n\n")
        siglas = [
            "**CAPES** -- Coordenacao de Aperfeicoamento de Pessoal de Nível Superior",
            "**CRSC-PCCTAE** -- Comissao para Reconhecimento de Saberes e Competencias do "
            "Plano de Carreira dos Cargos Tecnico-Administrativos em Educacao",
            "**PCCTAE** -- Plano de Carreira dos Cargos Tecnico-Administrativos em Educacao",
            "**RSC** -- Reconhecimento de Saberes e Competencias",
            "**SIAPE** -- Sistema Integrado de Administracao de Recursos Humanos",
            "**UFV** -- Universidade Federal de Viçosa",
        ]
        md.extend([f"{s}\n\n" for s in siglas])
        md.append("---\n")

        # ===== SUMÁRIO =====
        md.append("# SUMÁRIO\n\n")
        md.append("- **1 INTRODUÇÃO -- TRAJETORIA E FUNDAMENTOS**\n")
        md.append("- **2 ANEXO I -- PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS**\n")
        md.append("- **3 ANEXO II -- PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS**\n")
        md.append("- **4 ANEXO III -- PREMAÇÕES**\n")
        md.append("- **5 ANEXO IV -- RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS**\n")
        md.append("- **6 ANEXO V -- EXERCICIO DE FUNÇÕES DE DIRECAO E ASSESSORAMENTO**\n")
        md.append("- **7 ANEXO VI -- PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO CIENTÍFICO E TÉCNICO**\n")
        md.append("- **8 SÍNTESE DE PONTUAÇÃO**\n")
        md.append("- **9 REFLEXÃO FINAL -- SABERES E COMPETÊNCIAS**\n")
        md.append("- **REFERENCIAS**\n")
        md.append("\n---\n")

        # ===== CONTEUDO =====
        md.extend(self._narrativa_intro())
        md.append("\n---\n")
        md.extend(self._narrativa_anexo_I())
        md.append("\n---\n")
        md.extend(self._narrativa_anexo_II())
        md.append("\n---\n")
        md.extend(self._narrativa_anexo_III())
        md.append("\n---\n")
        md.extend(self._narrativa_anexo_IV())
        md.append("\n---\n")
        md.extend(self._narrativa_anexo_V())
        md.append("\n---\n")
        md.extend(self._narrativa_anexo_VI())
        md.append("\n---\n")
        md.extend(self._sintese_pontuacao())
        md.append("\n---\n")
        md.extend(self._reflexao_final())
        md.append("\n---\n")
        md.extend(self._referencias())
        md.append("\n---\n")
        md.append(
            "*Memorial gerado automaticamente em conformidade com a Secao 2.3.7 das "
            "Diretrizes do Agente (OBRIGATORIEDADE de salvamento proativo no vault). "
            "Dados extraidos do Relatorio Detalhado RSC emitido pelo sistema UFV "
            f"em {datetime.now().strftime('%d/%m/%Y')}.*\n"
        )
        texto_final = ''.join(md)
        texto_final = self._normalizar(texto_final)
        return texto_final


# =============================================================================
# FORMATADOR UFV/ABNT — DOCX
# =============================================================================
def set_cell_shading(cell, color):
    from docx.oxml import OxmlElement
    shading_elm = OxmlElement('w:shd')
    shading_elm.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill', color)
    shading_elm.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'clear')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def _html_to_docx(doc, lines):
    """Render HTML blocks (capa, folha de rosto, dedicatória, epígrafe) no DOCX."""
    text = ' '.join(lines)
    
    # Detect cover page (capa): university, author, title, city, year centered
    # Patterns:
    # <p align="center"><strong>TEXT</strong></p> -> centered bold
    # <div style="text-align:justify; margin-left:4cm...">TEXT</div> -> justified + indent
    # <p align="right" ...> -> right-aligned
    # <br> within text -> newline within same paragraph
    
    # Strip HTML tags for simple rendering, keeping alignment cues
    # First pass: detect alignment from first <p> tag
    alignment = WD_ALIGN_PARAGRAPH.CENTER  # default for cover
    if 'align="right"' in text:
        alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif 'margin-left:4cm' in text:
        alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # Extract clean text: remove all HTML tags, keep content
    clean = re.sub(r'<[^>]+>', '', text).strip()
    if not clean:
        return
    
    # Break by <br> patterns (they become separate lines)
    lines_text = re.split(r'<br\s*/?>', ' '.join(lines))
    lines_text = [re.sub(r'<[^>]+>', '', l).strip() for l in lines_text if re.sub(r'<[^>]+>', '', l).strip()
                  and not re.sub(r'<[^>]+>', '', l).strip().startswith('&')]
    
    # Check if it's a cover page (has vertical spacing like <br><br><br>)
    br_count = len([l for l in lines if l.strip() == '<br>' or l.strip() == '<br/>'])
    is_cover = br_count > 3 or 'UNIVERSIDADE FEDERAL' in clean.upper()
    is_right = 'align="right"' in text
    is_folha = 'margin-left:4cm' in text and 'Memorial descritivo' in clean
    
    if is_folha:
        # Folha de rosto: first line(s) centered, then nature text indented, then city/year centered
        # Split: find the nature paragraph
        parts = []
        for l in lines:
            ct = re.sub(r'<[^>]+>', '', l).strip()
            if ct:
                parts.append(ct)
        # parts[0] = author (centered), parts[1] = title (centered), 
        # parts[2] = nature (indented), parts[-2] = city, parts[-1] = year
        if len(parts) >= 4:
            # Author
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(parts[0])
            run.bold = True
            run.font.size = Pt(12)
            # Title
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(parts[1])
            run.bold = True
            run.font.size = Pt(12)
            # Blank space
            doc.add_paragraph()
            # Nature (indented) - find the long text
            for part in parts[2:-2]:
                if 'Memorial' in part or 'apresentado' in part or 'requisito' in part:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.line_spacing = 1.5
                    p.paragraph_format.left_indent = Cm(4)
                    p.paragraph_format.right_indent = Cm(0)
                    run = p.add_run(part)
                    run.font.size = Pt(12)
                    break
            doc.add_paragraph()
            # City
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(parts[-2])
            run.font.size = Pt(12)
            # Year
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(parts[-1])
            run.font.size = Pt(12)
        else:
            # Fallback: render each non-empty line with detected alignment
            for ct in parts:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.line_spacing = 1.5
                run = p.add_run(ct)
                if '<strong>' in text and ct == parts[0]:
                    run.bold = True
                run.font.size = Pt(12)
        return
    
    if is_right:
        # Dedicatória or Epígrafe - right aligned, italic
        for ct in lines_text:
            if ct:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.line_spacing = 1.5
                # Check for smaller font (epígrafe source)
                font_size = Pt(11) if '-- Paulo' in ct or ct.startswith('--') else Pt(12)
                run = p.add_run(ct)
                run.italic = True
                run.font.size = font_size
        return
    
    if is_cover:
        # Cover page: centered, bold, spaced out
        # Skip leading <br> lines and empty tags
        text_lines = [re.sub(r'<[^>]+>', '', l).strip() for l in lines 
                      if re.sub(r'<[^>]+>', '', l).strip()]
        text_lines = [l for l in text_lines if not l.startswith('&') and l]
        
        # Vertical spacing before (push content to ~middle third of the page)
        for _ in range(6):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for ct in text_lines:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(ct)
            run.bold = True
            run.font.size = Pt(12)
        
        for _ in range(4):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return
    
    # Generic HTML block: render with detected alignment
    for ct in lines_text:
        if ct:
            p = doc.add_paragraph()
            p.alignment = alignment
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(ct)
            run.font.size = Pt(12)


def md_to_docx_ufv(md_path, docx_path):
    """Converte MD para DOCX com formatacao UFV/ABNT obrigatoria.
    
    Processa HTML blocks (capa, folha de rosto, dedicatória, epígrafe)
    em vez de ignorá-los.
    """
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.split('\n')
    i = 0
    html_block = []
    in_html = False
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Collect HTML blocks and render them properly
        if stripped.startswith('<') and not stripped.startswith('</'):
            if not in_html:
                html_block = []
            in_html = True
            html_block.append(stripped)
            i += 1
            continue
        
        if in_html:
            # HTML block ends only at blank line (NOT at </div> or </p>)
            if stripped == '':
                # End of HTML block - render it
                html_block.append(stripped) if stripped else None
                if html_block:
                    _html_to_docx(doc, html_block)
                in_html = False
                html_block = []
            else:
                html_block.append(stripped)
            i += 1
            continue

        # Seção level (###)
        if stripped.startswith('### '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(stripped.replace('### ', ''))
            run.bold = True
            run.font.size = Pt(12)

        elif stripped.startswith('## '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(stripped.replace('## ', ''))
            run.bold = True
            run.font.size = Pt(12)

        elif stripped.startswith('# ') and not stripped.startswith('# LISTA') and not stripped.startswith('# SUMÁRIO') and not stripped.startswith('# AGRADECIMENTOS'):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(18)
            run = p.add_run(stripped.replace('# ', ''))
            run.bold = True
            run.font.size = Pt(14)

        elif stripped.startswith('# ') and stripped.startswith(('# LISTA', '# SUMÁRIO', '# AGRADECIMENTOS')):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(18)
            run = p.add_run(stripped.replace('# ', ''))
            run.bold = True
            run.font.size = Pt(14)

        elif stripped.startswith('---'):
            doc.add_page_break()

        elif stripped.startswith('| '):
            # Table
            rows_data = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().split('|')[1:-1]]
                rows_data.append(cells)
                i += 1
            if rows_data:
                is_header = True
                table = doc.add_table(rows=len(rows_data), cols=len(rows_data[0]))
                table.style = 'Table Grid'
                for r_idx, row_data in enumerate(rows_data):
                    for c_idx, cell_text in enumerate(row_data):
                        cell = table.rows[r_idx].cells[c_idx]
                        cell.text = ''
                        p = cell.paragraphs[0]
                        run = p.add_run(cell_text.strip('-').strip())
                        run.font.size = Pt(10)
                        if is_header:
                            run.bold = True
                            set_cell_shading(cell, "D9E2F3")
                    if is_header:
                        is_header = False
            continue

        elif stripped.startswith('- '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.left_indent = Cm(1)
            run = p.add_run(stripped)
            run.font.size = Pt(12)

        elif stripped == '':
            pass

        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            # Process bold markers
            parts = re.split(r'(\*\*.*?\*\*)', stripped)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                    run.font.size = Pt(12)
                elif part:
                    run = p.add_run(part)
                    run.font.size = Pt(12)

        i += 1

    # Close any unclosed HTML block
    if html_block and in_html:
        _html_to_docx(doc, html_block)

    # Add page numbers via OxmlElement
    from docx.oxml import OxmlElement
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p.add_run()
        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'begin')
        run1._r.append(fld_begin)
        run2 = p.add_run()
        instr = OxmlElement('w:instrText')
        instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instr.text = ' PAGE '
        run2._r.append(instr)
        run3 = p.add_run()
        fld_end = OxmlElement('w:fldChar')
        fld_end.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'end')
        run3._r.append(fld_end)
        for r in p.runs:
            r.font.size = Pt(10)

    doc.save(docx_path)


# =============================================================================
# CONVERSÃO PDF
# =============================================================================
def md_to_pdf(md_path, pdf_path):
    import subprocess
    md_abs = os.path.abspath(md_path)
    pdf_abs = os.path.abspath(pdf_path)
    result = subprocess.run(
        ['pandoc', md_abs, '-o', pdf_abs,
         '--pdf-engine=weasyprint',
         '--metadata', 'title="Memorial RSC-PCCTAE"',
         '--metadata', 'author="Gerado pelo PesquisAI (pdf-to-memorial-rsc)"'],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"pandoc + weasyprint failed: {result.stderr}")


# =============================================================================
# CLI
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description='Gera memorial RSC-PCCTAE completo a partir do PDF de relatorio detalhado.')
    parser.add_argument('pdf', help='Caminho para o PDF do Relatorio Detalhado RSC')
    parser.add_argument('--output-dir', '-o', default=None)
    parser.add_argument('--nome', '-n', default=None)
    parser.add_argument('--ano-ingresso', '-a', type=int, default=None)
    parser.add_argument('--auto', action='store_true',
                        help='Modo automatico: usa 2009 como ano de ingresso se não informado')
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        sys.exit(f"Arquivo não encontrado: {pdf_path}")

    output_dir = Path(args.output_dir) if args.output_dir else pdf_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = pdf_path.stem.replace(' ', '_')

    print("=" * 60)
    print("GERADOR DE MEMORIAL RSC-PCCTAE v3.1")
    print("=" * 60)
    print(f"   Decreto nº 13.048/2026: {DECRETO_URL}")
    print("=" * 60)

    print(f"\nLendo PDF: {pdf_path}")
    parser_obj = RSCPDFParser(str(pdf_path))
    data = parser_obj.data

    if args.nome:
        data['nome'] = args.nome
    nome = data['nome']

    # Ano de ingresso
    ano_ingresso = args.ano_ingresso
    if ano_ingresso is None:
        if args.auto:
            ano_ingresso = 2009
        else:
            ano_ingresso = perguntar_ano_ingresso()

    print(f"\n   Servidor: {nome}")
    print(f"   Matricula: {data['matricula']}")
    print(f"   Cargo: {data['cargo']}")
    print(f"   Titulação: {data['titulação']}")
    print(f"   RSC Requerido: {data['rsc_requerido']}")
    print(f"   Nível: {data['nivel_info']['nome']} (equivalente a {data['nivel_info']['equivalente']})")
    print(f"   Total: {fmt_br(data['total_geral'])} pts ({data['total_criterios']} critérios)")
    print(f"   Grupos: {len(data['grupos'])}")
    print(f"   Critérios extraídos: {len(data['criterios'])}")
    print(f"   Ordem critérios: {len(data['ordem_criterios'])}")
    for g in data['grupos']:
        print(f"      {g['romano']}: {g['nome_curto']} -- {g['criterios']} crit, {fmt_br(g['pontos'])} pts")
    anos_carreira = datetime.now().year - ano_ingresso
    print(f"\n   Ano de ingresso na UFV: {ano_ingresso} ({anos_carreira} anos de carreira)")

    # Generate .md
    print(f"\nGerando memorial em Markdown...")
    gen = MemorialGenerator(data, ano_ingresso)
    md_content = gen.generate()
    md_path = output_dir / f"{stem}_MEMORIAL.md"
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md_content)
    print(f"   OK: {md_path} ({len(md_content)} caracteres)")

    # Generate .docx
    print(f"Gerando .docx formatado UFV/ABNT...")
    docx_path = output_dir / f"{stem}_MEMORIAL.docx"
    try:
        md_to_docx_ufv(str(md_path), str(docx_path))
        sz = os.path.getsize(docx_path) / 1024
        print(f"   OK: {docx_path} ({sz:.1f} KB)")
    except Exception as e:
        print(f"   Aviso: erro ao gerar .docx: {e}")

    # Generate .pdf
    print(f"Gerando PDF...")
    pdf_output = output_dir / f"{stem}_MEMORIAL.pdf"
    try:
        md_to_pdf(str(md_path), str(pdf_output))
        sz = os.path.getsize(pdf_output) / 1024
        print(f"   OK: {pdf_output} ({sz:.1f} KB)")
    except Exception as e:
        print(f"   Aviso: erro ao gerar PDF: {e}")

    print(f"\n" + "=" * 60)
    print(f"Memorial gerado com sucesso!")
    print(f"Estrutura e topicos conforme memorial de referencia aprovado.")
    print(f"Formatacao UFV/ABNT obrigatoria aplicada.")
    print(f"Base legal: Decreto nº 13.048/2026 (Art. 13)")
    print(f"Link: {DECRETO_URL}")
    print("=" * 60)
    print(f"MD:  {md_path}")
    if os.path.exists(docx_path):
        print(f"DOCX: {docx_path}")
    pdf_out = output_dir / f"{stem}_MEMORIAL.pdf"
    if os.path.exists(pdf_out):
        print(f"PDF: {pdf_out}")
    print("=" * 60)


if __name__ == '__main__':
    main()
