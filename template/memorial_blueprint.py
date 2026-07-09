#!/usr/bin/env python3
"""
=============================================================================
BLUEPRINT: Memorial RSC-PCCTAE — Estrutura Programática v2.0
=============================================================================
Este arquivo serve como template/conceito para entender a estrutura
do memorial conforme o Decreto nº 13.048/2026 (Art. 13).

Uso:
    from memorial_blueprint import MemorialBlueprint
    bp = MemorialBlueprint()
    print(bp.structure)
=============================================================================
"""

DECRETO_URL = (
    "https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm"
)


class MemorialBlueprint:
    """Blueprint completo da estrutura do memorial RSC-PCCTAE.

    Organizado conforme o Decreto nº 13.048/2026, Art. 13:
      - Art. 13, II: descrição da trajetória profissional e individual
      - Art. 13, §1º, I: descrição das atividades e experiências
      - Art. 13, §1º, II: demonstração de alinhamento ao nível pleiteado

    E conforme o Manual de Normalização UFV 2025 e o
    modelo UFV-PPG para teses e dissertações.
    """

    # --- CONSTANTES DE FORMATAÇÃO UFV ---
    UFV_MARGINS = {'left': '3cm', 'right': '2cm', 'top': '3cm', 'bottom': '2cm'}
    UFV_PAGE = 'A4 (21 × 29.7 cm)'
    UFV_FONT = 'Arial 12pt'
    UFV_SPACING_BODY = 1.5
    UFV_SPACING_SINGLE = 1.0
    UFV_BLACK = 'RGB(0, 0, 0)'

    # --- MAPEAMENTO DE NÍVEIS RSC ---
    NIVEL_EQUIVALENCIA = {
        'VI': {
            'nome': 'VI',
            'equivalente': 'Doutor',
            'percentual': '75%',
            'destinado': 'servidor com diploma de mestrado',
            'lei': 'Art. 5º, §1º, VI',
        },
        'V': {
            'nome': 'V',
            'equivalente': 'Mestre',
            'percentual': '52%',
            'destinado': 'servidor com certificado de pós-graduação lato sensu',
            'lei': 'Art. 5º, §1º, V',
        },
        'IV': {
            'nome': 'IV',
            'equivalente': 'Graduação',
            'percentual': '30%',
            'destinado': 'servidor com diploma de graduação',
            'lei': 'Art. 5º, §1º, IV',
        },
    }

    # --- ESTRUTURA COMPLETA DO MEMORIAL (Conforme Decreto 13.048/2026) ---
    structure = {
        'pre_textuais': {
            'ordem': [
                '1. CAPA',
                '2. FOLHA DE ROSTO',
                '3. DEDICATÓRIA (opcional, sem título, alinhada à direita)',
                '4. AGRADECIMENTOS (obrigatório UFV-PPG, inclui CAPES)',
                '5. EPÍGRAFE (opcional, sem título, alinhada à direita)',
                '6. LISTA DE SIGLAS',
                '7. SUMÁRIO (último pré-textual)',
            ],
            'capa': {
                'modelo_ufv_ppg': [
                    'UNIVERSIDADE FEDERAL DE VIÇOSA',
                    'Nome do Autor',
                    'MEMORIAL DESCRITIVO PARA RECONHECIMENTO DE SABERES E '
                    'COMPETÊNCIAS — RSC-PCCTAE NÍVEL [NÍVEL]',
                    'VIÇOSA – MINAS GERAIS',
                    'Ano',
                ],
                'fonte': 'Arial 12pt bold, centralizado',
            },
            'folha_de_rosto': {
                'elementos': [
                    'Nome do Autor',
                    'Título',
                    'Natureza do trabalho (recuada 4cm da esquerda, espaço simples)',
                    'Link do Decreto nº 13.048/2026',
                    'VIÇOSA – MINAS GERAIS',
                    'Ano',
                ],
                'natureza_texto': (
                    'Modelo: "Memorial descritivo apresentado à Comissão para '
                    'Reconhecimento de Saberes e Competências do Plano de Carreira '
                    'dos Cargos Técnico-Administrativos em Educação (CRSC-PCCTAE) '
                    'da Universidade Federal de Viçosa como requisito para concessão '
                    'do RSC-PCCTAE Nível [NÍVEL], nos termos da Lei nº 11.091/2005 '
                    '(alterada pela Lei nº 15.367/2026), do Decreto nº 13.048/2026 '
                    'e da legislação correlata."'
                ),
                'observacao': (
                    'Incluir referência ao Art. 13 do Decreto nº 13.048/2026 '
                    'que estabelece os requisitos do memorial'
                ),
            },
            'agradecimentos': {
                'obrigatorio': True,
                'incluir_capes': True,
                'texto_capes': (
                    'O presente trabalho foi realizado com apoio da '
                    'Coordenação de Aperfeiçoamento de Pessoal de Nível '
                    'Superior – Brasil (CAPES) – Código de Financiamento 001.'
                ),
            },
        },
        'textuais': {
            'secao_1_trajetoria': {
                'fundamento_legal': 'Art. 13, II do Decreto nº 13.048/2026',
                'descricao': (
                    'Descrição da trajetória profissional e individual do '
                    'servidor desenvolvida ao longo da carreira, resultante '
                    'da atuação profissional na dinâmica de ensino, de '
                    'pesquisa e de extensão'
                ),
                'subsecoes': [
                    '1.1 Quem sou (identificação, cargo, tempo de serviço)',
                    '1.2 Trajetória profissional ao longo da carreira',
                    '1.3 Atuação na dinâmica de ensino, de pesquisa e de extensão',
                ],
                'conteudo': [
                    'Contextualização do autor e da trajetória',
                    'Fundamentação legal (Lei 11.091/2005, Decreto 13.048/2026)',
                    'Atuação em ensino, pesquisa e extensão ao longo da carreira',
                    'Demonstração de saberes, competências e experiências',
                ],
            },
            'secao_2_descricao_atividades': {
                'fundamento_legal': 'Art. 13, §1º, I do Decreto nº 13.048/2026',
                'descricao': (
                    'Descrição das atividades e das experiências profissionais '
                    'e individuais vinculadas aos requisitos previstos no '
                    'art. 3º, caput, incisos I a VI'
                ),
                'subsecoes': [
                    '2.1 Vinculação aos requisitos do Art. 3º (Incisos I a VI)',
                ],
            },
            'secao_3_demonstracao': {
                'fundamento_legal': 'Art. 13, §1º, II do Decreto nº 13.048/2026',
                'descricao': (
                    'Demonstração de que o conjunto da trajetória profissional '
                    'se alinha ao padrão de conhecimentos e competências que '
                    'justificam o reconhecimento naquele nível'
                ),
                'subsecoes': [
                    '3.1 Saberes, competências e nível pleiteado',
                ],
            },
            'anexos': {
                'fundamento_legal': 'Art. 3º, I a VI do Decreto nº 13.048/2026',
                'total': 6,
                'lista': [
                    'ANEXO I — Participação em Comissões, Grupos de Trabalho e Concursos',
                    'ANEXO II — Participação em Projetos Institucionais',
                    'ANEXO III — Premiações',
                    'ANEXO IV — Responsabilidades Técnico-Administrativas',
                    'ANEXO V — Exercício de Funções de Direção e Assessoramento',
                    'ANEXO VI — Produção, Prospecção e Difusão de Conhecimento',
                ],
                'correspondencia_art_3': {
                    'Anexo I': 'Art. 3º, I',
                    'Anexo II': 'Art. 3º, II',
                    'Anexo III': 'Art. 3º, III',
                    'Anexo IV': 'Art. 3º, IV',
                    'Anexo V': 'Art. 3º, V',
                    'Anexo VI': 'Art. 3º, VI',
                },
            },
            'sintese': {
                'tabela': 'Quadro geral com grupos, critérios e pontuação',
                'detalhamento': 'Distribuição dos pontos por anexo',
            },
            'reflexao_final': {
                'fundamento_legal': 'Art. 13 c/c Art. 15 do Decreto nº 13.048/2026',
                'tom': 'Narrativo-reflexivo, em primeira pessoa',
                'extensao': '3-4 parágrafos',
                'conteudo': [
                    'Síntese da trajetória documentada',
                    'Referência ao nível pleiteado e equivalência',
                    'Menção ao Art. 15 (saberes e competências diferenciados)',
                    'Submissão à CRSC-PCCTAE',
                ],
            },
        },
        'pos_textuais': {
            'referencias': {
                'formato': 'ABNT NBR 6023/2018',
                'espacamento': 'Simples, separadas por linha em branco',
                'alinhamento': 'Esquerda',
                'obrigatorias': [
                    'BRASIL. Lei nº 11.091/2005',
                    'BRASIL. Lei nº 15.367/2026',
                    f'BRASIL. Decreto nº 13.048/2026 (disponível em: {DECRETO_URL})',
                    'PIRES; SILVA. Normalização de trabalhos acadêmicos UFV, 2025',
                    'UFV. Relatório Detalhado RSC do servidor',
                ],
            },
        },
    }

    # --- CHECKLIST DE VERIFICAÇÃO ---
    checklist = """
    ✅ CONFORMIDADE COM DECRETO Nº 13.048/2026
    [ ] Art. 13, II — Descrição da trajetória profissional e individual
    [ ] Art. 13, §1º, I — Descrição das atividades vinculadas ao Art. 3º
    [ ] Art. 13, §1º, II — Demonstração de alinhamento ao nível pleiteado
    [ ] Art. 5º, §1º — Nível correto e equivalência adequada
    [ ] Art. 15 — Saberes e competências diferenciados
    [ ] Link do Decreto incluído na folha de rosto e nas referências

    ✅ FORMATAÇÃO UFV
    [ ] Papel A4 (21 × 29,7 cm)
    [ ] Margens: superior/esquerda 3 cm, inferior/direita 2 cm
    [ ] Fonte Arial 12 pt, cor preta
    [ ] Espaçamento 1,5 entre linhas (exceto: natureza, referências = 1,0)
    [ ] Paginação: algarismos arábicos, canto superior direito, a partir da introdução
    [ ] Títulos de seções primárias: 14 pt negrito
    [ ] Títulos centralizados sem numeração (AGRADECIMENTOS, SUMÁRIO, etc.)
    [ ] Natureza do trabalho na folha de rosto: recuo 4 cm, espaço simples

    ✅ ESTRUTURA
    [ ] Capa: Instituição → Autor → Título (com nível RSC) → Cidade → Ano
    [ ] Folha de rosto: Autor → Título → Natureza → Cidade → Ano
    [ ] Dedicatória (sem título, alinhada à direita)
    [ ] Agradecimentos (com CAPES)
    [ ] Epígrafe (sem título, alinhada à direita)
    [ ] Lista de siglas
    [ ] Sumário (último pré-textual)
    [ ] Seção 1 — Trajetória Profissional e Individual (Art. 13, II)
    [ ] Seção 2 — Descrição das Atividades (Art. 13, §1º, I)
    [ ] Seção 3 — Demonstração de Alinhamento (Art. 13, §1º, II)
    [ ] Anexos I a VI com referência aos incisos do Art. 3º
    [ ] Referências (ABNT, espaço simples)
    [ ] Seções primárias iniciando em nova página

    ✅ NORMALIZAÇÃO UFV-ABNT
    [ ] Skill ufv-abnt aplicada ao final do processo
    [ ] ABNT NBR 14724/2024 — Estrutura
    [ ] ABNT NBR 6023/2018 — Referências
    [ ] ABNT NBR 10520/2023 — Citações
    [ ] Manual de Normalização UFV 2025
    """

    def __str__(self):
        return self.checklist


if __name__ == '__main__':
    bp = MemorialBlueprint()
    print("=" * 60)
    print("BLUEPRINT — Memorial RSC-PCCTAE (v2.0)")
    print("Conforme Decreto nº 13.048/2026")
    print("=" * 60)
    print(bp)
    print("\nEstrutura completa dos elementos textuais (Art. 13):")
    print("  Seção 1 — TRAJETÓRIA PROFISSIONAL E INDIVIDUAL (Art. 13, II)")
    for sub in bp.structure['textuais']['secao_1_trajetoria']['subsecoes']:
        print(f"    {sub}")
    print("  Seção 2 — DESCRIÇÃO DAS ATIVIDADES E EXPERIÊNCIAS (Art. 13, §1º, I)")
    for sub in bp.structure['textuais']['secao_2_descricao_atividades']['subsecoes']:
        print(f"    {sub}")
    print("  Seção 3 — DEMONSTRAÇÃO DE ALINHAMENTO (Art. 13, §1º, II)")
    for sub in bp.structure['textuais']['secao_3_demonstracao']['subsecoes']:
        print(f"    {sub}")
    print(f"\nLink do Decreto: {DECRETO_URL}")
