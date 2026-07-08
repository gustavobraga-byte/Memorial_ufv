# Examples — Memorial RSC-PCCTAE

Arquivos gerados a partir do relatório oficial UFV:
`RSC Detalhado_03jun.pdf` (Ricardo Gandini Lugão, 698 pts, 15 critérios)

| Arquivo | Descrição |
|---------|-----------|
| `memorial_rsc_example.md` | Memorial completo em Markdown (fonte de verdade) |
| `memorial_rsc_example.docx` | Formatado UFV/ABNT (Arial 12pt, margens 3/2cm) |
| `memorial_rsc_example.pdf` | PDF pronto para entrega |

## Como usar estes exemplos

1. Abra o `.docx` no Microsoft Word
2. Atualize o sumário (clique → F9 → "Atualizar sumário inteiro")
3. Verifique a formatação
4. Salve como PDF final

## Comando para gerar

```bash
python3 run.py "RSC Detalhado_03jun.pdf" --output-dir examples --nome "NOME"
```
