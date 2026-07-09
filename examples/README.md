# Examples — Memorial RSC-PCCTAE v2.0

Arquivos gerados a partir do relatório oficial UFV em conformidade com o
**Decreto nº 13.048, de 3 de julho de 2026** (Art. 13).

> Decreto nº 13.048/2026: https://www.planalto.gov.br/ccivil_03/_ato2023-2026/2026/decreto/d13048.htm

| Arquivo | Descrição |
|---------|-----------|
| `memorial_rsc_example.md` | Memorial completo em Markdown (fonte de verdade) |
| `memorial_rsc_example.docx` | Formatado UFV/ABNT (Arial 12pt, margens 3/2cm) |
| `memorial_rsc_example.pdf` | PDF pronto para entrega |

## Como usar estes exemplos

1. Abra o `.docx` no Microsoft Word
2. Atualize o sumário (clique → F9 → "Atualizar sumário inteiro")
3. Verifique a formatação UFV/ABNT
4. Salve como PDF final

## Comando para gerar (v2.0)

```bash
# O sistema perguntará o ano de ingresso interativamente
python3 run.py "RSC Detalhado_03jun.pdf" --output-dir examples --nome "NOME"

# Ou informe o ano de ingresso diretamente
python3 run.py "RSC Detalhado_03jun.pdf" --output-dir examples --nome "NOME" --ano-ingresso 1992
```

## Novidades da v2.0

- ✅ Detecção automática do nível RSC (V ou VI) e equivalência (Mestrado/Doutorado)
- ✅ Pergunta interativa do ano de ingresso na UFV
- ✅ Trajetória profissional e individual conforme Art. 13, II do Decreto
- ✅ Descrição das atividades vinculadas aos requisitos do Art. 3º (§1º, I)
- ✅ Demonstração de alinhamento ao nível pleiteado (§1º, II)
- ✅ Link do Decreto nº 13.048/2026 incluído em todas as saídas
- ✅ Chamada da skill ufv-abnt para normalização final
