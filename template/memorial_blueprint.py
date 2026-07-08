#!/usr/bin/env python3
"""
=============================================================================
BLUEPRINT: Memorial RSC-PCCTAE — Estrutura Programática
=============================================================================
Este arquivo serve como template/conceito para entender a estrutura
do memorial. Use como referência para criar versões customizadas.

Uso:
    from memorial_blueprint import MemorialBlueprint
    bp = MemorialBlueprint()
    print(bp.structure)
=============================================================================
"""

class MemorialBlueprint:
    """Blueprint completo da estrutura do memorial RSC-PCCTAE.
    
    Organizado segundo o Manual de Normalização UFV 2025 e o
    modelo UFV-PPG para teses e dissertações.
    """
    
    # --- CONSTANTES DE FORMATAÇÃO UFV ---
    UFV_MARGINS = {'left': '3cm', 'right': '2cm', 'top': '3cm', 'bottom': '2cm'}
    UFV_PAGE = 'A4 (21 × 29.7 cm)'
    UFV_FONT = 'Arial 12pt'
    UFV_SPACING_BODY = 1.5
    UFV_SPACING_SINGLE = 1.0
    UFV_BLACK = 'RGB(0, 0, 0)'
    
    # --- ESTRUTURA COMPLETA DO MEMORIAL ---
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
                    'TÍTULO DO TRABALHO',
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
                    'VIÇOSA – MINAS GERAIS',
                    'Ano',
                ],
                'natureza_texto': (
                    'Modelo: "Tipo do trabalho apresentado à Universidade '
                    'Federal de Viçosa como requisito para..."'
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
            'introducao': {
                'conteudo': [
                    'Contextualização do autor e da trajetória',
                    'Fundamentação legal (Lei 11.091/2005, Decreto 13.048/2026)',
                    'Estrutura do memorial',
                ],
            },
            'anexos': {
                'total': 6,
                'lista': [
                    'ANEXO I — Participação em Comissões, Grupos de Trabalho e Concursos',
                    'ANEXO II — Participação em Projetos Institucionais',
                    'ANEXO III — Premiações',
                    'ANEXO IV — Responsabilidades Técnico-Administrativas',
                    'ANEXO V — Exercício de Funções de Direção e Assessoramento',
                    'ANEXO VI — Produção, Prospecção e Difusão de Conhecimento',
                ],
            },
            'sintese': {
                'tabela': 'Quadro geral com grupos, critérios e pontuação',
                'detalhamento': 'Distribuição dos pontos por anexo',
            },
            'reflexao_final': {
                'tom': 'Narrativo-reflexivo, em primeira pessoa',
                'extensao': '3-4 parágrafos',
            },
        },
        'pos_textuais': {
            'referencias': {
                'formato': 'ABNT NBR 6023/2018',
                'espacamento': 'Simples, separadas por linha em branco',
                'alinhamento': 'Esquerda',
            },
        },
    }
    
    # --- CHECKLIST DE VERIFICAÇÃO ---
    checklist = """
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
    [ ] Capa: Instituição → Autor → Título → Cidade → Ano
    [ ] Folha de rosto: Autor → Título → Natureza → Cidade → Ano
    [ ] Dedicatória (sem título, alinhada à direita)
    [ ] Agradecimentos (com CAPES)
    [ ] Epígrafe (sem título, alinhada à direita)
    [ ] Lista de siglas
    [ ] Sumário (último pré-textual)
    [ ] Referências (ABNT, espaço simples)
    [ ] Seções primárias iniciando em nova página
    """
    
    def __str__(self):
        return self.checklist


if __name__ == '__main__':
    bp = MemorialBlueprint()
    print("=" * 60)
    print("BLUEPRINT — Memorial RSC-PCCTAE (UFV/ABNT)")
    print("=" * 60)
    print(bp)
    print("\nEstrutura completa dos elementos pré-textuais:")
    for i, item in enumerate(bp.structure['pre_textuais']['ordem'], 1):
        print(f"  {item}")
