# Examples — Memorial RSC-PCCTAE v3.3

Memorial de exemplo com dados anônimos (placeholders), gerado
pelo próprio `run.py --example`. **Nenhum servidor real é identificado.**

> Decreto nº 13.048/2026: https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm

| Arquivo | Descrição |
|---------|-----------|
| `memorial_rsc_example.md` | Memorial completo em Markdown (fonte de verdade) |
| `memorial_rsc_example.docx` | Gerado via `run.py --example` |

> Os arquivos .pdf e .docx **não são mais distribuídos** no ZIP da skill
> — são gerados sob demanda pelo comando abaixo.

## Regenerar os exemplos

```bash
# Gera .md e .docx na pasta examples/
python3 run.py --example -o examples/
```

O .docx pode ser aberto no Word para gerar PDF.

## Comando principal (com PDF real)

```bash
python3 run.py "RSC Detalhado_fulano.pdf"
```

## Novidades da v3.3

- ✅ `--example`: gera memorial de exemplo com dados 100% anônimos
- ✅ Exemplos (.pdf/.docx) removidos do ZIP; `run.py` é a fonte única
- ✅ `build_example_data()` como fonte de dados de demonstração
- ✅ CAPES removido dos placeholders
- ✅ Dados do servidor sempre como "NOME DO(A) SERVIDOR(A)"
