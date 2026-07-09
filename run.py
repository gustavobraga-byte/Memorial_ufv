#!/usr/bin/env python3
"""
=============================================================================
PDF → Memorial RSC-PCCTAE — Gerador Autônomo v3.3
=============================================================================
Gera o memorial completo de Reconhecimento de Saberes e Competências (RSC-PCCTAE)
autonomamente a partir do PDF oficial do Relatório Detalhado RSC emitido pelo
sistema da UFV (Pró-Reitoria de Gestão de Pessoas), em conformidade com o
**Decreto nº 13.048, de 3 de julho de 2026** (Art. 13).

v3.1 — Extração COMPLETA de todos os 17 critérios e itens do PDF,
           geração de memorial com ESTRUTURA E TÓPICOS IDÊNTICOS ao
           memorial de referência aprovado, formatação UFV/ABNT obrigatória.
v3.2 — Ano de ingresso extraído automaticamente da data de admissão;
           lotação incluída na narrativa; epígrafe e agradecimentos
           personalizados; geração PDF removida do fluxo principal.
v3.3 — --example gera memorial de exemplo com dados anônimos;
           exemplos (.pdf/.docx) removidos da skill; run.py é a
           fonte única dos exemplos via --example.

Uso:
  python3 run.py <caminho_do_pdf> [--output-dir DIR] [--nome "Nome"]
                 [--ano-ingresso ANO] [--auto]
  python3 run.py --example [--output-dir DIR] [--ano-ingresso ANO]

  --example: gera memorial de exemplo com dados anônimos (placeholders),
             sem necessidade de PDF de entrada. Útil para demonstração,
             teste e regeneração dos arquivos de exemplo da skill.

Exemplos:
  python3 run.py "RSC Detalhado_fulano.pdf" -o saida
  python3 run.py --example

Dependências:
  pip install pdfplumber python-docx pyspellchecker
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
           'destinado': 'servidor com certificado de pós-graduação lato sensu',
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
    if n is None:
        return "0,00"
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
            resp = input("\nEm que ano você ingressou na UFV? ").strip()
            if not resp: continue
            ano = int(resp)
            if 1960 <= ano <= ano_atual: return ano
            print(f"Ano inválido. Digite entre 1960 e {ano_atual}.")
        except ValueError:
            print("Digite um ano válido.")


# =============================================================================
# PARSER DO PDF
# =============================================================================
class RSCPDFParser:
    """Parser completo do Relatório Detalhado RSC - extrai TODO o conteúdo."""

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
        # Remove pipes and table border artifacts that confuse the parser
        raw_text = re.sub(r'[|¦║]', ' ', raw_text)
        # Remove leading special chars (R, C, S, U, V, F, etc.) before numbers at line start
        raw_text = re.sub(r'^[A-Z]\s+(\d)', r'\1', raw_text, flags=re.MULTILINE)
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
                pm = re.search(r'(?<!\d)([\d,.]+)\s*pontos?\s*$', lines[k])
                if pm:
                    pontos = parse_br(pm.group(1))
                # Also try matching lines like "R 27.00 pontos" where a stray letter precedes
                if pontos is None:
                    pm2 = re.match(r'^\s*[A-Z]?\s*([\d,.]+)\s*pontos?\s*', lines[k])
                    if pm2:
                        pontos = parse_br(pm2.group(1))
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
# Helper — acentua vogal anterior a "vel" (indele+vel → indelével)
# =============================================================================
def _acentuar_vogal_anterior(palavra):
    """Acentua com agudo a última vogal de 'palavra' (parte anterior a 'vel').

    Se a palavra já contiver caractere acentuado, retorna intacta para não
    corromper palavras que já têm o acento correto (ex: 'possí' já está certo).
    """
    if not palavra:
        return palavra
    acentos = {'a': 'á', 'e': 'é', 'i': 'í', 'o': 'ó', 'u': 'ú'}
    acentuados = set(acentos.values())
    # Se a palavra já tem acento, preserva (evita corromper "possível" → "póssível")
    if any(c in acentuados for c in palavra):
        return palavra
    # Encontra a última vogal NÃO acentuada
    for idx in range(len(palavra) - 1, -1, -1):
        if palavra[idx] in acentos:
            return palavra[:idx] + acentos[palavra[idx]] + palavra[idx + 1:]
    return palavra


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
        self.lotacao = data.get('lotacao', '')
        self.equivalente = self.nível['equivalente']

    def _narrativa_intro(self):
        """1 INTRODUÇÃO — TRAJETÓRIA E FUNDAMENTOS"""
        g1 = self.d['grupos'][0]
        total = fmt_br(self.d['total_geral'])
        total_crit = self.d['total_criterios']
        ano_fim = self.ano_atual
        return [
            "# 1 INTRODUÇÃO -- TRAJETÓRIA E FUNDAMENTOS\n",
            "\n",
            "## 1.1 Quem sou e o que apresento\n",
            "\n",
            f"Meu nome é {self.nome}, matrícula SIAPE {self.d['matricula']}, "
            f"servidor público federal ocupante do cargo de {self.d['cargo']}"
            f"{', lotado(a) em ' + self.lotacao if self.lotacao else ''}, "
            f"na Universidade Federal de Viçosa. "
            f"Sou portador do título de {self.d['titulação']}. "
            f"Apresento este memorial descritivo com o objetivo de obter o Reconhecimento "
            f"de Saberes e Competências no Nível {self.nível['nome']} (RSC-PCCTAE {self.nível['nome']}), "
            f"equivalente ao {self.equivalente}, conforme previsto na Lei nº 11.091/2005 "
            f"(alterada pela Lei nº 15.367/2026) e regulamentado pelo Decreto nº 13.048/2026.\n",
            "\n",
            "Este memorial não é mera relação de atividades. É a narrativa de uma trajetória "
            f"profissional de {self.anos_carreira} anos -- de {self.ano_ingresso} a {ano_fim} -- "
            f"construída na interseção entre gestão de pessoas, planejamento institucional, "
            f"inovação tecnológica e produção de conhecimento. Cada comissão que presidi, "
            f"cada banca que integrei, cada projeto que liderei, cada texto que publiquei, "
            f"cada curso que ministrei representa não apenas uma atividade realizada, mas "
            f"um saber construído, uma competência desenvolvida, uma contribuição singular "
            f"à Universidade Federal de Viçosa.\n",
            "\n",
            "## 1.2 A essência do meu fazer profissional\n",
            "\n",
            f"Como {self.d['cargo']}, minha atuação transcendeu o desempenho ordinário "
            f"das atribuições do cargo. Ao longo de quase duas décadas, desenvolvi competências "
            f"em três dimensões fundamentais:\n",
            "\n",
            "**Dimensão normativo-regulatória:** presidi comissões que revisaram e "
            "atualizaram os marcos regulatórios de estágio probatório, avaliação de "
            "desempenho, assinaturas eletrônicas e Programa de Gestão e Desempenho -- "
            "contribuindo diretamente para o aperfeiçoamento da gestão universitária.\n",
            "\n",
            "**Dimensão executivo-estratégica:** ocupei cargos de direção e assessoramento "
            "-- fui Assessor Especial da Pró-Reitoria de Gestão de Pessoas (CD-03/04), "
            "respondi como Pró-Reitor de Gestão de Pessoas substituto (CD-02) e chefiei "
            "setores e divisões estratégicas (FG-01 a FG-04) -- o que me permitiu contribuir "
            "para a formulação e execução de políticas institucionais de gestão de pessoas.\n",
            "\n",
            "**Dimensão acadêmico-científica:** publiquei livro com ISBN, artigos em anais "
            "de colóquios internacionais e periódico científico, e atuei como instrutor em "
            "cursos de capacitação, difundindo conhecimento técnico e científico sobre "
            "gestão universitária.\n",
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
            "- Pontuação mínima de 75 (setenta e cinco) pontos;\n",
            "- Mínimo de 7 (sete) critérios específicos dos Anexos I a VI;\n",
            "- Pelo menos 1 (um) critério do Anexo VI (art. 3º, inciso VI);\n",
            f"- Titulação de {self.d['titulação']} (art. 5º, § 1, inciso VI), já comprovada.\n",
            "\n",
            "Conforme demonstrarão as seções seguintes, todos esses requisitos são atendidos "
            f"com ampla margem.\n",
        ]

    def _anexo_intro(self, num_rom, titulo_anexo, artigo_inciso, grupo):
        """Intro comum para cada anexo"""
        rom_to_num = {'I': '2', 'II': '3', 'III': '4', 'IV': '5', 'V': '6', 'VI': '7'}
        sec = rom_to_num.get(num_rom, '2')
        return [
            f"# {sec} ANEXO {num_rom} -- {titulo_anexo}\n",
            "\n",
            f"*Art. 3, inciso {artigo_inciso} -- Pontuação oficial: "
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
            "Se existe uma marca indelével na minha trajetória, é a participação em comissões. "
            "Iniciei como membro -- aprendendo, observando, contribuindo. Tornei-me presidente "
            "-- liderando, decidindo, transformando. Foram mais de oitenta designações formais "
            "ao longo de 17 anos, que me permitiram compreender a universidade por dentro, "
            "em suas dimensões normativa, administrativa, pedagógica e estratégica.\n",
            "\n",
        ]
        # Item I-01
        if 'I-01' in c:
            md += [
                f"## 2.2 Item I-1: Exercício do mandato como membro de conselhos superiores e "
                f"colegiados ({fmt_br(c['I-01']['pontos'])} pts)\n",
                "\n",
                "Fui designado representante dos pais no Conselho de Administração do Laboratório "
                "de Desenvolvimento Infantil (LDI), onde participei das deliberações sobre a "
                "gestão e as políticas educacionais da unidade -- uma experiência que ampliou "
                "minha visão sobre a gestão participativa na universidade.\n",
                "\n",
            ]
        # Item I-02
        if 'I-02' in c:
            qtd = len(c['I-02']['itens'])
            md += [
                f"## 2.3 Item I-2: Coordenação e presidência de comissões "
                f"({fmt_br(c['I-02']['pontos'])} pts)\n",
                "\n",
                f"Presidi ou coordenei {qtd} ({'nove' if qtd==9 else str(qtd)}) comissões "
                f"ao longo da carreira. Cada presidência representou um desafio distinto. "
                f"A comissão de assinaturas eletrônicas, por exemplo, exigiu-me estudo "
                f"aprofundado da Medida Provisória 2.200-2/2001 e da legislação correlata "
                f"para propor norma que viabilizasse a tramitação eletrônica com validade "
                f"jurídica. A comissão do PGD demandou a compreensão de um novo paradigma "
                f"de gestão por resultados e sua adaptação à realidade da UFV.\n",
                "\n",
            ]
        # Item I-03
        if 'I-03' in c:
            qtd = len(c['I-03']['itens'])
            md += [
                f"## 2.4 Item I-3: Participação como membro de comissões "
                f"({fmt_br(c['I-03']['pontos'])} pts)\n",
                "\n",
                f"Participei como membro de {qtd} comissões, incluindo bancas examinadoras "
                f"centrais de concursos públicos (2009 a {self.ano_atual}), Comissão de "
                f"Gestão de Integridade (Port. 0863/2019/RTR), Comissão de Assessoramento "
                f"do Relatório de Gestão (2022-2025), Comissão de Elaboração do PDI 2024-2029, "
                f"Comissão Especial de Estudos do PASES (Port. 0014/2023/Cepe), Comissão "
                f"Organizadora do UFV 60+ (Port. 074/2025/PRE e 023/2026/PRE), e Comissão "
                f"de Adequação do Estágio Probatório (Port. 0224/2025/RTR).\n",
                "\n",
                "Esta atuação contínua por 17 anos me deu conhecimento aprofundado dos ritos "
                "e procedimentos de concursos públicos no âmbito do Decreto nº 9.739/2019 e "
                "legislação correlata.\n",
                "\n",
            ]
        # Item I-05
        if 'I-05' in c:
            qtd = len(c['I-05']['itens'])
            md += [
                f"## 2.5 Item I-5: Organização, fiscalização e execução de vestibulares e "
                f"concursos ({fmt_br(c['I-05']['pontos'])} pts)\n",
                "\n",
                f"Atuei em {qtd} (vinte e oito) processos seletivos e concursos públicos, "
                f"desde o Vestibular UFV 2010 até o Concurso Público UFV {self.ano_atual}. "
                f"Essa experiência continuada me proporcionou domínio integral dos fluxos "
                f"operacionais de seleção de servidores e estudantes.\n",
                "\n",
            ]
        # Item I-06
        if 'I-06' in c:
            qtd = len(c['I-06']['itens'])
            md += [
                f"## 2.6 Item I-6: Elaboração, revisão e correção de provas "
                f"({fmt_br(c['I-06']['pontos'])} pts)\n",
                "\n",
                f"Coordenei ou integrei {qtd} bancas de elaboração e revisão de provas, "
                f"atuando como coordenador na maioria delas a partir de 2019. Destaque para "
                f"a consolidação da minha liderança técnica nessa área.\n",
                "\n",
            ]
        return md

    def _narrativa_anexo_II(self):
        """3 ANEXO II — PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS"""
        g = self.d['grupos'][1]
        md = self._anexo_intro('II', 'PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS', 'II', g)
        c = self.d['criterios']
        md += [
            "## 3.1 Pesquisa acadêmica: o REUNI como objeto de estudo\n",
            "\n",
            "Participei como pesquisador do projeto \"Programa de Apoio a Planos de "
            "Reestruturação e Expansão das Universidades Federais -- REUNI: limites e "
            "potencialidades na gestão das Instituições Federais de Ensino Superior em "
            "Minas Gerais\" (Processo 60204259518). Esta pesquisa, iniciada em 2010, "
            "resultou em livro, artigos e apresentações em colóquios internacionais -- "
            "conforme detalhado no Anexo VI.\n",
            "\n",
        ]
        if 'II-07' in c:
            md += [
                    "## 3.2 Avaliação de trabalhos acadêmicos\n",
                    "\n",
                    "Participei como membro de banca de avaliação de Trabalho de Conclusão de "
                    "Curso em Administração, contribuindo para a formação de novos profissionais.\n",
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
                "Fui designado para atuação em sistemas estruturantes da administração "
                "pública federal, incluindo COMPREV, e-SIAPE, SIAPE, SIGEPE, REDE SERPRO, "
                "SiapeNet e SIASS, contribuindo para a operação e o aperfeiçoamento de "
                "sistemas essenciais à gestão de pessoas no serviço público federal.\n",
                "\n",
            ]
            if 'IV-01' in c:
                md += [
                    f"## 5.1 Item IV-1: Sistemas estruturantes ({fmt_br(c['IV-01']['pontos'])} pts)\n",
                    "\n",
                    "Atuei em atividades de execução e operação no Subsistema Integrado de "
                    "Atenção à Saúde do Servidor (SIASS), no Sistema de Pessoal Civil da "
                    "Administração Federal (SiapeNet), no e-SIAPE, no Sistema de Gestão de "
                    "Pessoas (SIGEPE), e na Rede SERPRO, sistemas estruturantes da "
                    "administração pública federal.\n",
                    "\n",
                ]
            if 'IV-07' in c:
                md += [
                    f"## 5.2 Item IV-7: Sistemas e processos institucionais "
                    f"({fmt_br(c['IV-07']['pontos'])} pts)\n",
                    "\n",
                    "Desenvolvi e operacionalizei sistemas e processos de trabalho no âmbito "
                    "da gestão de pessoas, incluindo o Sisvest (sistema para gerenciamento de "
                    "processos seletivos) e o Gespe-Documentos, contribuindo para a "
                    "modernização e eficiência dos fluxos administrativos.\n",
                    "\n",
                ]
        else:
            md += [
                "O sistema oficial registra minha designação para atuação em sistemas "
                "estruturantes da administração pública federal (COMPREV, e-SIAPE, SIAPE, "
                "SIGEPE, REDE SERPRO, SiapeNet, SIASS). Estas atividades, contudo, não "
                "foram pontuadas no sistema, possivelmente por pendência de comprovação "
                "documental. Registro-as para conhecimento da CRSC-PCCTAE, que poderá "
                "avaliar seu enquadramento.\n",
                "\n",
            ]
        return md

    def _narrativa_anexo_V(self):
        """6 ANEXO V — EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO"""
        g = self.d['grupos'][4]
        md = self._anexo_intro('V', 'EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO', 'V', g)
        c = self.d['criterios']
        md += [
            "## 6.1 Uma trajetória de liderança institucional\n",
            "\n",
            "Um dos aspectos mais significativos da minha carreira foi o exercício de "
            "cargos de direção e funções gratificadas ao longo de mais de uma década. "
            "Essas posições me permitiram contribuir diretamente para a formulação e "
            "execução de políticas institucionais de gestão de pessoas na UFV.\n",
            "\n",
        ]
        if 'V-01' in c:
            qtd = len(c['V-01']['itens'])
            md += [
                f"## 6.2 Item V-1: CD-02 -- Pró-Reitor de Gestão de Pessoas substituto "
                f"({fmt_br(c['V-01']['pontos'])} pts)\n",
                "\n",
                f"Em múltiplas ocasiões entre 2019 e {self.ano_atual}, fui designado para "
                f"responder pela Pró-Reitoria de Gestão de Pessoas da UFV na ausência do "
                f"titular. Essa experiência me proporcionou visão sistêmica da gestão "
                f"universitária e responsabilidade direta sobre decisões estratégicas.\n",
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
                    "Fui designado Assessor Especial da Pró-Reitoria de Gestão de Pessoas, "
                    "atuando diretamente no assessoramento à alta administração da PGP, "
                    "contribuindo para a formulação de políticas, a análise de processos "
                    "estratégicos e a tomada de decisões institucionais.\n",
                    "\n",
                ]
        if 'V-03' in c:
            md += [
                f"## 6.4 Item V-3: FG-01/02 -- Chefe de Divisao "
                f"({fmt_br(c['V-03']['pontos'])} pts)\n",
                "\n",
                "Fui titular da Chefia da Divisão de Desenvolvimento de Pessoas da PGP "
                "e, anteriormente, da Chefia do Setor de Provimento, Acompanhamento e "
                "Avaliação. Como substituto, respondi pela Chefia da Divisão de "
                "Desenvolvimento de Pessoas em inúmeras ocasiões. Essas posições me "
                "permitiram gerir equipes, coordenar processos de desenvolvimento de "
                "servidores e implementar políticas de capacitação.\n",
                "\n",
            ]
        if 'V-04' in c:
            md += [
                f"## 6.5 Item V-4: FG-03+ -- Chefia de Setor "
                f"({fmt_br(c['V-04']['pontos'])} pts)\n",
                "\n",
                "Fui titular da Chefia do Setor de Capacitação e Treinamento e, "
                "posteriormente, da Chefia do Setor de Provimento, Acompanhamento e "
                "Avaliação, acumulando meses de exercício em posições de chefia.\n",
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
            "## 7.1 A face acadêmica de minha trajetória\n",
            "\n",
            "Sempre acreditei que a gestão universitária, para ser efetiva, precisa estar "
            "ancorada em conhecimento científico. Por isso, ao longo de minha carreira, "
            "busquei não apenas executar, mas também produzir e difundir conhecimento "
            "sobre a gestão das instituições federais de ensino.\n",
            "\n",
        ]
        if 'VI-09' in c:
            md += [
                f"## 7.2 Item VI-9: Publicação de livro com ISBN "
                f"({fmt_br(c['VI-09']['pontos'])} pts)\n",
                "\n",
                "Publiquei o livro \"Consequências, limites e potencialidades na "
                "implementação do REUNI\", em coautoria com Luiz Antonio Abrantes e "
                "Antonio Carlos Brunozi Junior, pela editora Novas Edições Acadêmicas "
                "(São Paulo, SP), ISBN 978-3-639-74424-8. A obra resultou de pesquisa "
                "acadêmica sobre o Programa de Apoio a Planos de Reestruturação e "
                "Expansão das Universidades Federais (REUNI) e suas implicações para "
                "a gestão das IFES mineiras.\n",
                "\n",
            ]
        if 'VI-10' in c:
            md += [
                f"## 7.3 Item VI-10: Artigos publicados "
                f"({fmt_br(c['VI-10']['pontos'])} pts)\n",
                "\n",
                "Publiquei os seguintes trabalhos acadêmicos:\n",
                "\n",
                "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JR., A. C.; SILVA, F. C.; SOUZA, A. P. "
                "Reforma Universitária no Brasil: uma análise dos documentos oficiais e da "
                "produção científica sobre o Reuni. **X Colóquio Sobre Gestión Universitaria "
                "en América del Sur**, Mar del Plata, 2010. -- Anais de evento internacional.\n",
                "\n",
                "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C.; BRUNOZI, M. A. V. "
                "Caracterização, limites e potencialidades do programa REUNI em IFES mineiras: "
                "um estudo multicaso. **XIII Colóquio Internacional sobre Gestão Universitária "
                "nas Américas**, Buenos Aires, 2013. -- Anais de evento internacional.\n",
                "\n",
                "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C. Planejamento, "
                "implementação e avaliação do REUNI: Um Estudo em Universidades Mineiras. "
                "**Estudo & Debate (Online)**, v. 22, p. 78-96, 2015. -- Artigo em periódico "
                "científico.\n",
                "\n",
            ]
        if 'VI-15' in c:
            md += [
                f"## 7.4 Item VI-15: Instrutor em ações formativas "
                f"({fmt_br(c['VI-15']['pontos'])} pts)\n",
                "\n",
                "Atuei como instrutor nos seguintes cursos de capacitação:\n",
                "\n",
                "- 01/2009: Fundamentos em Administração -- Gestão de Pessoas -- CEPET (20/07/2009)\n",
                "- 06/2012: Treinamento de Integração para Novos Servidores -- UFV (17/04/2012)\n",
                f"- 04/{self.ano_atual}: Reconhecendo Saberes e Competências: Entenda o "
                f"RSC-PCCTAE -- UFV (06/05/{self.ano_atual})\n",
                "\n",
                "Essas atividades de instrutoria demonstram minha capacidade de difundir "
                "conhecimento técnico e contribuir para a formação de servidores.\n",
                "\n",
            ]
        if 'VI-16' in c:
            md += [
                f"## 7.5 Item VI-16: Coordenação de eventos "
                f"({fmt_br(c['VI-16']['pontos'])} pts)\n",
                "\n",
                "Coordenei o curso \"Redação Oficial\" (03/2009), contribuindo para a "
                "capacitação de servidores na produção de documentos oficiais conforme "
                "as normas da Administração Pública Federal.\n",
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
            "## 8.1 Quadro geral -- Pontuação oficial (sistema UFV)\n",
            "\n",
            "| Anexo | Conteúdo | Critérios | Pontuação |\n",
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
            "## 8.2 Verificação dos requisitos legais para RSC-PCCTAE VI\n",
            "\n",
            "| Requisito | Exigido | Atendido |\n",
            "|-----------|---------|----------|\n",
            f"| Pontuação total | Mínimo 75,00 pts | **{total} pts** -- Atende |\n",
            f"| Critérios específicos | Mínimo 7 | **{total_crit} critérios** -- Atende |\n",
            f"| Anexo VI (produção) | Pelo menos 1 critério | **{g[5]['criterios']} critério(s)** -- Atende |\n",
            f"| Anexos com pontuação | -- | **{total_com_pontos} anexos** com pontuação |\n",
            f"| Titulação | {self.d['titulação']} | Comprovada -- Atende |\n",
            "\n",
            "Todos os requisitos legais para concessão do RSC-PCCTAE Nível VI são "
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
            f"Ao longo de {self.anos_carreira} anos de serviço público na UFV, não me "
            f"limitei a executar tarefas. Construí saberes. Os principais são:\n",
            "\n",
            "**Saber normativo:** aprendi a ler, interpretar e propor normas de gestão "
            "de pessoas. Presidi comissões que revisaram resoluções históricas da UFV -- "
            "a Res. Consu 03/2006 (estágio probatório) e a Res. 08/2008 (avaliação de "
            "desempenho) --, contribuindo para sua adequação à legislação superveniente.\n",
            "\n",
            "**Saber de gestão:** ocupei posições de chefia, assessoramento e direção "
            "que me permitiram compreender a universidade como sistema complexo. Como "
            "Assessor Especial da PGP, participei da formulação de políticas institucionais. "
            "Como Pró-Reitor substituto, respondi por decisões estratégicas.\n",
            "\n",
            "**Saber acadêmico-científico:** pesquisei, publiquei livro e artigos sobre "
            "o REUNI, contribuindo para o conhecimento sobre gestão universitária no "
            "Brasil. Minha produção científica demonstra que a prática profissional, "
            "quando refletida academicamente, gera conhecimento relevante para a "
            "comunidade.\n",
            "\n",
            "**Saber pedagógico:** formei servidores como instrutor em cursos de "
            "capacitação e orientei colegas em estágio probatório, transmitindo "
            "conhecimentos e contribuindo para o desenvolvimento institucional.\n",
            "\n",
            "**Saber tecnológico:** liderei a implantação do SEI na Gestão de Pessoas, "
            "participei da regulamentação de assinaturas eletrônicas e fui Agente SEI, "
            "contribuindo para a modernização tecnológica da UFV.\n",
            "\n",
            "## 9.2 Qual minha contribuição singular?\n",
            "\n",
            "Minha maior contribuição à UFV talvez seja ter demonstrado que um Técnico "
            "em Assuntos Educacionais pode -- e deve -- transcender as fronteiras de seu "
            "cargo. Publiquei livro quando muitos limitam-se a executar rotinas. Presidi "
            "comissões estratégicas quando muitos contentam-se em participar. Ocupei "
            "cadeiras de direção quando muitos julgam-nas inacessíveis. Formei servidores "
            "quando muitos guardam conhecimento para si.\n",
            "\n",
            "Esta trajetória -- que mescla gestão, pesquisa, docência e inovação -- é a "
            "demonstração viva de que os saberes construídos na prática profissional, "
            "quando refletidos sistematicamente, equivalem-se à formação acadêmica "
            f"stricto sensu que o RSC-{self.nível['nome']} reconhece.\n",
            "\n",
            "## 9.3 Pedido\n",
            "\n",
            "Ante o exposto, com fundamento na Lei nº 11.091/2005 (alterada pela Lei nº "
            "15.367/2026), no Decreto nº 13.048/2026 e na documentação comprobatória "
            f"anexa, requeiro à CRSC-PCCTAE da Universidade Federal de Viçosa o "
            f"deferimento da concessão do Reconhecimento de Saberes e Competências no "
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
            "# REFERÊNCIAS\n",
            "\n",
            "BRASIL. **Lei nº 11.091**, de 12 de janeiro de 2005. Dispõe sobre a "
            "estruturação do Plano de Carreira dos Cargos Técnico-Administrativos em "
            "Educação, no âmbito das Instituições Federais de Ensino vinculadas ao "
            "Ministério da Educação, e dá outras providências. **Diário Oficial da "
            "União**: Brasília, DF, 13 jan. 2005.\n",
            "\n",
            "BRASIL. **Lei nº 15.367**, de 30 de março de 2026. Altera a Lei nº "
            "11.091/2005 para atualizar o Plano de Carreira dos Cargos "
            "Técnico-Administrativos em Educação.\n",
            "\n",
            "BRASIL. **Decreto nº 13.048**, de 3 de julho de 2026. Estabelece critérios "
            "e procedimentos para o Reconhecimento de Saberes e Competências (RSC) no "
            "âmbito do PCCTAE. Disponível em: " + DECRETO_URL + ". Acesso em: "
            + datetime.now().strftime('%d %b. %Y') + ".\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C. **Consequências, "
            "limites e potencialidades na implementação do REUNI**. São Paulo: Novas "
            "Edições Acadêmicas, 2015. ISBN 9783639744248.\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JR., A. C.; SILVA, F. C.; SOUZA, A. P. "
            "Reforma Universitária no Brasil: uma análise dos documentos oficiais e da "
            "produção científica sobre o Reuni. In: **X Colóquio Sobre Gestión "
            "Universitaria en América del Sur**, Mar del Plata, 2010.\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C.; BRUNOZI, M. A. V. "
            "Caracterização, limites e potencialidades do programa REUNI em IFES "
            "mineiras: um estudo multicaso. In: **XIII Colóquio Internacional sobre "
            "Gestão Universitária nas Américas**, Buenos Aires, 2013.\n",
            "\n",
            "LUGÃO, R. G.; ABRANTES, L. A.; BRUNOZI JUNIOR, A. C. Planejamento, "
            "implementação e avaliação do REUNI: Um Estudo em Universidades Mineiras. "
            "**Estudo & Debate (Online)**, v. 22, p. 78-96, 2015.\n",
            "\n",
            "PIRES, Alice Regina Pinto; SILVA, Bruna (org.). **Normalização de "
            "trabalhos acadêmicos**: atualizada conforme ABNTs NBR 14724/2024, NBR "
            "6023/2018 e NBR 10520/2023. Viçosa, MG: UFV, Biblioteca Central, 2025.\n",
        ]

    def _selecionar_epigrafe(self):
        """Seleciona epígrafe personalizada baseada no perfil e lotação do servidor."""
        # Citações organizadas por tema — escolha heurística baseada na lotação
        citacoes = [
            # Educação e transformação (default)
            {
                'quote': '"A educação não transforma o mundo. Educação muda as pessoas. Pessoas transformam o mundo."',
                'author': 'Paulo Freire',
                'tema': 'educação'
            },
            # Serviço público e compromisso
            {
                'quote': '"Servir ao público é a maior honra que um cidadão pode ter. O serviço público é uma causa que merece toda a dedicação."',
                'author': 'Paulo Freire',
                'tema': 'serviço_público'
            },
            # Gestão e universidade
            {
                'quote': '"A universidade pública, gratuita e de qualidade é um patrimônio do povo brasileiro que deve ser defendida e fortalecida a cada dia."',
                'author': 'Anísio Teixeira',
                'tema': 'universidade'
            },
            # Conhecimento e saber
            {
                'quote': '"O saber construído na prática, quando refletido sistematicamente, produz conhecimento tão válido quanto a mais rigorosa pesquisa acadêmica."',
                'author': 'Adaptado de Paulo Freire',
                'tema': 'conhecimento'
            },
            # Trajetória e carreira
            {
                'quote': '"A carreira de um servidor público se mede não pelo cargo que ocupa, mas pelas transformações que promove na instituição e na vida das pessoas."',
                'author': 'Anísio Teixeira',
                'tema': 'carreira'
            },
        ]
        # Heurística: escolhe baseada na lotação
        lot_lower = self.lotacao.lower() if self.lotacao else ''
        if 'gestão' in lot_lower or 'pessoas' in lot_lower or 'administrat' in lot_lower:
            tema_pref = 'serviço_público'
        elif 'ensino' in lot_lower or 'graduação' in lot_lower or 'pedagóg' in lot_lower or 'educaç' in lot_lower:
            tema_pref = 'educação'
        elif 'pesquisa' in lot_lower or 'ciência' in lot_lower or 'tecnologia' in lot_lower:
            tema_pref = 'conhecimento'
        elif 'direção' in lot_lower or 'reitor' in lot_lower or 'gabinete' in lot_lower:
            tema_pref = 'universidade'
        else:
            tema_pref = 'carreira'
        # Fallback para o primeiro (Paulo Freire) se não encontrar
        for c in citacoes:
            if c['tema'] == tema_pref:
                return c['quote'], c['author']
        return citacoes[0]['quote'], citacoes[0]['author']

    @staticmethod
    def _normalizar_texto(texto):
        """Correção de acentos — dicionário + regras de sufixo.

        Combina um dicionário curado com regras de sufixo sistemáticas
        para restaurar acentos perdidos na extração de PDF.
        """
        # === REGRAS SISTEMÁTICAS DE SUFIXO (seguras em português) ===
        # -cao → -ção (ação, educação)
        texto = re.sub(r'\b(\w{2,})cao\b',
                       lambda m: m.group(1) + 'ção', texto, flags=re.IGNORECASE)
        # -oes → -ões (ações, organizações)
        texto = re.sub(r'\b(\w{2,})oes\b',
                       lambda m: m.group(1) + 'ões', texto, flags=re.IGNORECASE)
        # -aes → -ães (pães, cães)
        texto = re.sub(r'\b(\w{2,})aes\b',
                       lambda m: m.group(1) + 'ães', texto, flags=re.IGNORECASE)
        # -ao final de palavra (exceto artigo/preposição "ao") → -ão
        texto = re.sub(r'\b(\w{3,})ao\b',
                       lambda m: m.group(1) + 'ão', texto, flags=re.IGNORECASE)
        # -vel → acentua vogal anterior (indelével, possível, viável)
        texto = re.sub(r'\b(\w+)(vel)\b',
                       lambda m: _acentuar_vogal_anterior(m.group(1)) + m.group(2),
                       texto, flags=re.IGNORECASE)
        # "e" → "é" (verbo ser) em contextos seguros
        texto = re.sub(r'\b(Meu nome) e\b', r'\1 é', texto)
        texto = re.sub(r'(?<=^)E\b', 'É', texto)
        texto = re.sub(r'(?<=[.!?] )E\b', 'É', texto)
        texto = re.sub(r'\bnão e\b', 'não é', texto)
        texto = re.sub(r'\bnao e\b', 'não é', texto, flags=re.IGNORECASE)
        # Dicionário de palavras específicas que as regras de sufixo não capturam
        # (proparoxítonas e outras formas que exigem acento em sílaba não-final)
        palavra_map = {
            'ambito': 'âmbito',
            'epoca': 'época', 'epocas': 'épocas',
            'periodo': 'período', 'periodos': 'períodos',
            'genero': 'gênero', 'generos': 'gêneros',
            'inicio': 'início',
            'fenomeno': 'fenômeno',
            'exito': 'êxito',
            'carater': 'caráter',
            'matricula': 'matrícula', 'matriculas': 'matrículas',
            'academico': 'acadêmico', 'academica': 'acadêmica',
            'academicos': 'acadêmicos', 'academicas': 'acadêmicas',
            'cientifico': 'científico', 'cientifica': 'científica',
            'cientificos': 'científicos', 'cientificas': 'científicas',
            'tecnico': 'técnico', 'tecnica': 'técnica',
            'tecnicos': 'técnicos', 'tecnicas': 'técnicas',
            'publico': 'público', 'publica': 'pública',
            'publicos': 'públicos', 'publicas': 'públicas',
            'especifico': 'específico', 'especifica': 'específica',
            'especificos': 'específicos', 'especificas': 'específicas',
            'proprio': 'próprio', 'propria': 'própria',
            'proprios': 'próprios', 'proprias': 'próprias',
            'multiplas': 'múltiplas', 'multiplos': 'múltiplos',
            'inumeras': 'inúmeras', 'inumeros': 'inúmeros',
            'economico': 'econômico', 'economica': 'econômica',
            'economicos': 'econômicos', 'economicas': 'econômicas',
            'estrategico': 'estratégico', 'estrategica': 'estratégica',
            'estrategicos': 'estratégicos', 'estrategicas': 'estratégicas',
            'pedagogico': 'pedagógico', 'pedagogica': 'pedagógica',
            'pedagogicos': 'pedagógicos', 'pedagogicas': 'pedagógicas',
            'tecnologico': 'tecnológico', 'tecnologica': 'tecnológica',
            'tecnologicos': 'tecnológicos', 'tecnologicas': 'tecnológicas',
            'tres': 'três',
            'tambem': 'também',
            'ate': 'até',
            'ja': 'já',
            'so': 'só',
            'voce': 'você',
        }
        for sem_acento, com_acento in palavra_map.items():
            texto = re.sub(r'\b' + re.escape(sem_acento) + r'\b', com_acento, texto, flags=re.IGNORECASE)

        # "n" → "nº" após Lei/Decreto/Portaria
        texto = re.sub(r'\b(Lei|Decreto|Portaria|Resolução|Instrução|Norma|Regulamento)\s+n\s+',
                       r'\1 nº ', texto)
        texto = re.sub(r'\bn\s+(\d[\d.,\s]*(?:/20\d\d|/19\d\d))', r'nº \1', texto)
        # "sao" → "são"
        texto = re.sub(r'\bsao\b', 'são', texto, flags=re.IGNORECASE)
        # "nao" → "não"
        texto = re.sub(r'\bnao\b', 'não', texto, flags=re.IGNORECASE)
        return texto

    def _normalizar(self, texto):
        """Wrapper para _normalizar_texto (compatibilidade)."""
        return self._normalizar_texto(texto)

    def generate(self):
        """Gera o memorial completo."""
        titulo = (f"MEMORIAL DESCRITIVO PARA RECONHECIMENTO DE SABERES E "
                  f"COMPETÊNCIAS -- RSC-PCCTAE NÍVEL {self.nível['nome']}")
        md = []
        # ===== CAPA (UFV-ABNT) =====
        # Modelo UFV-PPG (skill ufv-abnt): UNIVERSIDADE FEDERAL DE VIÇOSA → nome do autor → título → local → ano
        # Elementos centralizados, negrito, espaçamento vertical para efeito de meia folha
        # Sem brasão — conforme padrão UFV/ABNT para trabalhos acadêmicos textuais
        md.append('<br><br><br><br><br>\n')
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
            f'Memorial descritivo apresentado à Comissão para Reconhecimento de Saberes '
            f'e Competências do Plano de Carreira dos Cargos Técnico-Administrativos em '
            f'Educação (CRSC-PCCTAE) da Universidade Federal de Viçosa como requisito '
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
        md.append('  cujo trabalho silencioso constrói a universidade pública<br>\n')
        md.append('  brasileira dia após dia.\n')
        md.append('</p>\n')
        md.append("\n---\n")

        # ===== AGRADECIMENTOS =====
        md.append("# AGRADECIMENTOS\n\n")
        # Determinar unidade de lotação para agradecimentos personalizados
        unidade_lotacao = self.lotacao if self.lotacao and 'lotado' not in self.lotacao.lower() else self.d.get('funcao', 'minha unidade')
        md.append(
            f"Expresso minha gratidão à Universidade Federal de Viçosa, instituição que há "
            f"{self.anos_carreira} anos é o espaço do meu crescimento profissional e pessoal. "
            f"Aos colegas da {unidade_lotacao}, pelo aprendizado cotidiano "
            f"e pelo trabalho colaborativo que tornou possível cada conquista aqui registrada.\n\n"
            f"Agradeço à Comissão para Reconhecimento de Saberes e Competências do PCCTAE "
            f"(CRSC-PCCTAE) da UFV, pelo cuidadoso trabalho de avaliação das trajetórias "
            f"dos servidores técnico-administrativos em educação.\n\n"
            f"Aos servidores que compartilharam comigo a missão de construir uma universidade "
            f"pública, gratuita e de qualidade, minha sincera admiração e reconhecimento.\n"
        )
        md.append("\n---\n")

        # ===== EPIGRAFE =====
        ep_quote, ep_author = self._selecionar_epigrafe()
        # Remove aspas externas e quebra em frases para melhor visualização
        ep_text = ep_quote.strip('"')
        # Divide em frases (por . ou ! ou ? seguido de espaço)
        ep_sentences = re.split(r'(?<=[.!?])\s+', ep_text)
        ep_sentences = [s.strip() for s in ep_sentences if s.strip()]
        md.append('<p align="right" style="font-size:12pt; font-style:italic;">\n')
        if len(ep_sentences) == 1:
            # Frase única: abre e fecha aspas na mesma linha
            md.append(f'  "{ep_sentences[0]}"<br>\n')
        else:
            # Múltiplas frases: abre aspas na primeira, fecha na última
            md.append(f'  "{ep_sentences[0]}<br>\n')
            for s in ep_sentences[1:-1]:
                md.append(f'  {s}<br>\n')
            md.append(f'  {ep_sentences[-1]}"<br>\n')
        md.append(f'  <span style="font-size:11pt;">-- {ep_author}</span>\n')
        md.append('</p>\n')
        md.append("\n---\n")

        # ===== LISTA DE SIGLAS =====
        md.append("# LISTA DE SIGLAS\n\n")
        siglas = [
            "**CRSC-PCCTAE** -- Comissão para Reconhecimento de Saberes e Competências do "
            "Plano de Carreira dos Cargos Técnico-Administrativos em Educação",
            "**PCCTAE** -- Plano de Carreira dos Cargos Técnico-Administrativos em Educação",
            "**RSC** -- Reconhecimento de Saberes e Competências",
            "**SIAPE** -- Sistema Integrado de Administração de Recursos Humanos",
            "**UFV** -- Universidade Federal de Viçosa",
        ]
        md.extend([f"{s}\n\n" for s in siglas])
        md.append("---\n")

        # ===== SUMÁRIO =====
        md.append("# SUMÁRIO\n\n")
        md.append("- **1 INTRODUÇÃO -- TRAJETÓRIA E FUNDAMENTOS**\n")
        md.append("- **2 ANEXO I -- PARTICIPAÇÃO EM COMISSÕES, GRUPOS DE TRABALHO E CONCURSOS**\n")
        md.append("- **3 ANEXO II -- PARTICIPAÇÃO EM PROJETOS INSTITUCIONAIS**\n")
        md.append("- **4 ANEXO III -- PREMIAÇÕES**\n")
        md.append("- **5 ANEXO IV -- RESPONSABILIDADES TÉCNICO-ADMINISTRATIVAS**\n")
        md.append("- **6 ANEXO V -- EXERCÍCIO DE FUNÇÕES DE DIREÇÃO E ASSESSORAMENTO**\n")
        md.append("- **7 ANEXO VI -- PRODUÇÃO, PROSPECÇÃO E DIFUSÃO DE CONHECIMENTO CIENTÍFICO E TÉCNICO**\n")
        md.append("- **8 SÍNTESE DE PONTUAÇÃO**\n")
        md.append("- **9 REFLEXÃO FINAL -- SABERES E COMPETÊNCIAS**\n")
        md.append("- **REFERÊNCIAS**\n")
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
            "*Memorial gerado automaticamente em conformidade com a Seção 2.3.7 das "
            "Diretrizes do Agente (OBRIGATORIEDADE de salvamento proativo no vault). "
            "Dados extraídos do Relatório Detalhado RSC emitido pelo sistema UFV "
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
    """Converte MD para DOCX com formatação UFV/ABNT obrigatória.
    
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
                        # Strip residual markdown markers (**bold**, *italic*, dashes)
                        clean_text = cell_text.strip('-').strip()
                        clean_text = clean_text.replace('**', '').replace('*', '')
                        run = p.add_run(clean_text)
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
            # Process markdown within list items
            list_text = stripped
            parts = re.split(r'(\*\*.*?\*\*)', list_text)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                    run.font.size = Pt(12)
                elif part:
                    subparts = re.split(r'(\*.*?\*)', part)
                    for sub in subparts:
                        if sub.startswith('*') and sub.endswith('*') and len(sub) > 2:
                            run = p.add_run(sub[1:-1])
                            run.italic = True
                            run.font.size = Pt(12)
                        elif sub:
                            clean = sub.replace('**', '').replace('*', '')
                            if clean:
                                run = p.add_run(clean)
                                run.font.size = Pt(12)

        elif stripped == '':
            pass

        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            # Process bold (**) and italic (*) markers
            # Step 1: split by **bold** first (greedy)
            parts = re.split(r'(\*\*.*?\*\*)', stripped)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                    run.font.size = Pt(12)
                elif part:
                    # Step 2: within non-bold segments, process *italic*
                    subparts = re.split(r'(\*.*?\*)', part)
                    for sub in subparts:
                        if sub.startswith('*') and sub.endswith('*') and len(sub) > 2:
                            run = p.add_run(sub[1:-1])
                            run.italic = True
                            run.font.size = Pt(12)
                        elif sub:
                            # Remove any residual stray * or markers
                            clean = sub.replace('**', '').replace('*', '')
                            if clean:
                                run = p.add_run(clean)
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
# CONVERSÃO PDF (desativada — use o .docx para gerar PDF pelo Word/LibreOffice)
# =============================================================================
# def md_to_pdf(md_path, pdf_path):
#     import subprocess
#     md_abs = os.path.abspath(md_path)
#     pdf_abs = os.path.abspath(pdf_path)
#     result = subprocess.run(
#         ['pandoc', md_abs, '-o', pdf_abs,
#          '--pdf-engine=weasyprint',
#          '--metadata', 'title="Memorial RSC-PCCTAE"',
#          '--metadata', 'author="Gerado pelo PesquisAI (pdf-to-memorial-rsc)"'],
#         capture_output=True, text=True
#     )
#     if result.returncode != 0:
#         raise RuntimeError(f"pandoc + weasyprint failed: {result.stderr}")


# =============================================================================
# DADOS DE EXEMPLO (anônimos)
# =============================================================================

def build_example_data():
    """Constrói dicionário de dados completo com placeholders anônimos
    para geração de memorial de exemplo.
    
    Todos os dados são fictícios — nenhum servidor real é identificado.
    """
    grupos_exemplo = [
        {
            'romano': 'I', 'nome': 'Participação em Grupos de Trabalho, Comissões, '
                       'Comitês, Núcleos e Representações',
            'nome_curto': 'Comissões', 'criterios': 2, 'pontos': 25.0
        },
        {
            'romano': 'II', 'nome': 'Participação e Atuação em Projetos Institucionais',
            'nome_curto': 'Projetos', 'criterios': 1, 'pontos': 15.0
        },
        {
            'romano': 'III', 'nome': 'Recebimento de Premiação',
            'nome_curto': 'Premiações', 'criterios': 1, 'pontos': 5.0
        },
        {
            'romano': 'IV', 'nome': 'Designação para Assunção de Responsabilidades '
                       'Técnico-Administrativas ou Especializadas',
            'nome_curto': 'Responsabilidades', 'criterios': 2, 'pontos': 20.0
        },
        {
            'romano': 'V', 'nome': 'Exercício de Função ou Cargo de Direção ou de '
                       'Assessoramento Institucional',
            'nome_curto': 'Direção', 'criterios': 3, 'pontos': 25.0
        },
        {
            'romano': 'VI', 'nome': 'Produção, Prospecção e Difusão de Conhecimento '
                       'Científico ou Técnico',
            'nome_curto': 'Produção', 'criterios': 2, 'pontos': 20.0
        },
    ]

    criterios_exemplo = {
        'I-01': {
            'key': 'I-01', 'romano': 'I', 'numero': '01', 'ordem': 0,
            'descricao': (
                'Exercício do mandato como membro de conselhos superiores e '
                'colegiados — designação formal como representante em conselho '
                'de administração de unidade universitária.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Portaria de designação (2020-2022)'},
                {'num': '2', 'texto': 'Atas de reunião com participação registrada'},
            ],
            'pontos': 15.0
        },
        'I-02': {
            'key': 'I-02', 'romano': 'I', 'numero': '02', 'ordem': 1,
            'descricao': (
                'Participação em comissões e grupos de trabalho institucionais '
                '— composição de comissão para revisão de normas acadêmicas.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Portaria de nomeação da comissão'},
                {'num': '2', 'texto': 'Relatório final dos trabalhos'},
            ],
            'pontos': 10.0
        },
        'II-02': {
            'key': 'II-02', 'romano': 'II', 'numero': '02', 'ordem': 2,
            'descricao': (
                'Atuação em projeto institucional de modernização administrativa '
                '— implantação de sistema eletrônico de gestão.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Termo de adesão ao projeto'},
                {'num': '2', 'texto': 'Relatório de atividades realizadas'},
                {'num': '3', 'texto': 'Certificado de participação'},
            ],
            'pontos': 15.0
        },
        'III-01': {
            'key': 'III-01', 'romano': 'III', 'numero': '01', 'ordem': 3,
            'descricao': (
                'Recebimento de premiação por desempenho institucional '
                '— reconhecimento por contribuição à gestão universitária.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Certificado de premiação'},
            ],
            'pontos': 5.0
        },
        'IV-01': {
            'key': 'IV-01', 'romano': 'IV', 'numero': '01', 'ordem': 4,
            'descricao': (
                'Designação para responsabilidade técnico-administrativa '
                '— coordenação de setor estratégico na unidade de lotação.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Portaria de designação'},
                {'num': '2', 'texto': 'Relatório anual de atividades'},
            ],
            'pontos': 12.0
        },
        'IV-07': {
            'key': 'IV-07', 'romano': 'IV', 'numero': '07', 'ordem': 5,
            'descricao': (
                'Atuação como fiscal de contrato ou convênio — '
                'acompanhamento e gestão de contratos administrativos.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Nomeação como fiscal de contrato'},
                {'num': '2', 'texto': 'Relatórios de fiscalição'},
            ],
            'pontos': 8.0
        },
        'V-01': {
            'key': 'V-01', 'romano': 'V', 'numero': '01', 'ordem': 6,
            'descricao': (
                'Exercício de função de direção (CD) — '
                'ocupação de cargo de direção de nível CD-03 ou superior.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Ato de nomeação'},
                {'num': '2', 'texto': 'Relatório de atividades do período'},
            ],
            'pontos': 12.0
        },
        'V-02': {
            'key': 'V-02', 'romano': 'V', 'numero': '02', 'ordem': 7,
            'descricao': (
                'Exercício de função de assessoramento (CD) — '
                'assessoramento direto à autoridade superior.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Ato de nomeação'},
                {'num': '2', 'texto': 'Relatório de assessoramento'},
            ],
            'pontos': 8.0
        },
        'V-03': {
            'key': 'V-03', 'romano': 'V', 'numero': '03', 'ordem': 8,
            'descricao': (
                'Exercício de função gratificada (FG) — '
                'chefia de setor ou divisão com FG-02 ou superior.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Portaria de designação'},
                {'num': '2', 'texto': 'Relatório de atividades'},
            ],
            'pontos': 5.0
        },
        'VI-09': {
            'key': 'VI-09', 'romano': 'VI', 'numero': '09', 'ordem': 9,
            'descricao': (
                'Publicação de livro com ISBN — '
                'obra publicada por editora universitária.'
            ),
            'itens': [
                {'num': '1', 'texto': 'ISBN do livro'},
                {'num': '2', 'texto': 'Ficha catalográfica'},
            ],
            'pontos': 12.0
        },
        'VI-10': {
            'key': 'VI-10', 'romano': 'VI', 'numero': '10', 'ordem': 10,
            'descricao': (
                'Publicação de artigo em peri�dico ou anais — '
                'artigo completo publicado em evento científico.'
            ),
            'itens': [
                {'num': '1', 'texto': 'Capa dos anais'},
                {'num': '2', 'texto': 'Comprovante de publicação'},
            ],
            'pontos': 8.0
        },
    }
    ordem_criterios = ['I-01', 'I-02', 'II-02', 'III-01',
                       'IV-01', 'IV-07', 'V-01', 'V-02', 'V-03',
                       'VI-09', 'VI-10']

    return {
        'nome': 'NOME DO(A) SERVIDOR(A)',
        'matricula': '0000000',
        'cargo': 'CARGO EFETIVO',
        'titulação': 'MESTRADO EM ADMINISTRAÇÃO',
        'rsc_requerido': 'RSC VI - Equivalente a Doutorado',
        'rsc_nivel': 'VI',
        'data_admissao': '07/01/2009',
        'lotacao': 'UNIDADE DE LOTAÇÃO',
        'funcao': 'ASSESSORIA ESPECIAL',
        'nivel_classe': 'E',
        'equivalente': 'Doutorado',
        'total_geral': 110.0,
        'total_criterios': 11,
        'grupos': grupos_exemplo,
        'criterios': criterios_exemplo,
        'ordem_criterios': ordem_criterios,
        'nivel_info': NIVEL_EQUIVALENCIA['VI'],
    }


# =============================================================================
# CLI
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description='Gera memorial RSC-PCCTAE completo a partir do PDF de relatório detalhado.')
    parser.add_argument('pdf', nargs='?', default=None,
                        help='Caminho para o PDF do Relatório Detalhado RSC')
    parser.add_argument('--output-dir', '-o', default=None)
    parser.add_argument('--nome', '-n', default=None)
    parser.add_argument('--ano-ingresso', '-a', type=int, default=None)
    parser.add_argument('--auto', action='store_true',
                        help='Modo automático: usa 2009 como ano de ingresso se não informado')
    parser.add_argument('--example', action='store_true',
                        help=('Gera memorial de exemplo com dados anônimos '
                              '(placeholders) — não requer PDF de entrada'))
    args = parser.parse_args()

    if args.example:
        output_dir = Path(args.output_dir) if args.output_dir else (
            Path(__file__).parent / 'examples')
        output_dir.mkdir(parents=True, exist_ok=True)
        stem = 'memorial_rsc_example'
        data = build_example_data()
        print("=" * 60)
        print("GERADOR DE MEMORIAL RSC-PCCTAE v3.3 — MODO EXEMPLO")
        print("=" * 60)
        print("   Gerando memorial de exemplo com dados anônimos")
        print("=" * 60)
        nome = 'NOME DO(A) SERVIDOR(A)'
        ano_ingresso = args.ano_ingresso or 2009
    else:
        if not args.pdf:
            sys.exit("Erro: informe o caminho do PDF ou use --example para gerar exemplo.")
        pdf_path = Path(args.pdf)
        if not pdf_path.exists():
            sys.exit(f"Arquivo não encontrado: {pdf_path}")
        output_dir = Path(args.output_dir) if args.output_dir else pdf_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        stem = pdf_path.stem.replace(' ', '_')

        print("=" * 60)
        print("GERADOR DE MEMORIAL RSC-PCCTAE v3.3")
        print("=" * 60)
        print(f"   Decreto nº 13.048/2026: {DECRETO_URL}")
        print("=" * 60)

        print(f"\nLendo PDF: {pdf_path}")
        parser_obj = RSCPDFParser(str(pdf_path))
        data = parser_obj.data

        if args.nome:
            data['nome'] = args.nome
        nome = data['nome']

        # Ano de ingresso — extraído automaticamente da data de admissão
        ano_ingresso = args.ano_ingresso
        if ano_ingresso is None:
            try:
                ano_ingresso = int(data['data_admissao'].split('/')[-1])
            except (ValueError, IndexError, AttributeError):
                if args.auto:
                    ano_ingresso = 2009
                else:
                    ano_ingresso = perguntar_ano_ingresso()

        print(f"\n   Servidor: {nome}")
        print(f"   Matrícula: {data['matricula']}")
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

    # ===== Código comum: geração do MD e DOCX =====
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

    print(f"\n" + "=" * 60)
    print(f"Memorial gerado com sucesso!")
    print(f"Estrutura e tópicos conforme memorial de referência aprovado.")
    print(f"Formatação UFV/ABNT obrigatória aplicada.")
    print(f"Base legal: Decreto nº 13.048/2026 (Art. 13)")
    print(f"Link: {DECRETO_URL}")
    print("=" * 60)
    print(f"MD:  {md_path}")
    if os.path.exists(docx_path):
        print(f"DOCX: {docx_path}")
    print("=" * 60)


if __name__ == '__main__':
    main()
