# Contas Mensais — controle de parcelas, dívidas e gastos do mês

Planilha de Excel gerada por código para quem quer **saber quanto deve, quanto sobra do salário e quanto pode gastar por semana**.

## O que a planilha faz

**Aba "Contas Mensais"**
- Uma linha por conta: parcelamentos, dívidas com pessoas, contas fixas (aluguel, assinaturas) e metas de economia ("Caixinha").
- Você informa o valor total, o nº de parcelas e o mês da 1ª parcela, e a planilha distribui tudo pelos meses.
- Status automático: *Ativo*, *Última parcela*, *Quitado* ou *Fixo*. A última parcela fica destacada em rosa.
- Totais por mês: fatura do cartão, Pix/boleto, caixinha, total, **sobra do salário** e sobra por dia.
- Escolhendo um mês no topo, você vê a fatura, o total e quanto ainda deve.

**Aba "Gastos do Mês"**
- Anote os gastos do dia a dia (data, descrição, categoria, forma de pagamento, valor).
- Limite do mês dividido em **limites semanais editáveis**, com saldo acumulado e alerta quando passar.
- Gastos no **cartão** entram sozinhos na fatura do mês seguinte (na aba Contas Mensais).
- Mostra o "respiro": quanto vai sobrar quando o próximo salário cair.

## Como gerar a planilha

**Opção 1 — sem instalar nada:** baixe `modelo/Contas_Mensais_exemplo.xlsx`, apague os dados de exemplo e use.

**Opção 2 — pelo GitHub:** na aba **Actions**, abra "Gerar planilha de exemplo" e baixe o arquivo em *Artifacts*.

**Opção 3 — no seu computador (Python 3):**
```bash
pip install -r requirements.txt
python gerar_planilha.py                      # gera com os dados de exemplo
python gerar_planilha.py --dados caminho/meus_dados.json --saida Minhas_Contas.xlsx
```

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `gerar_planilha.py` | O código que monta a planilha |
| `config.json` | As regras: dia de vencimento do cartão, categorias, formas de pagamento, nº de meses |
| `dados/exemplo.json` | Dados **fictícios** usados no modelo |
| `modelo/` | Planilha de exemplo pronta |

## Usando com os seus dados (sem publicar nada)

1. Copie `dados/exemplo.json` para uma pasta **fora deste repositório** (ou para `privado/`, que é ignorada pelo Git) com o nome `meus_dados.json`.
2. Troque os valores pelos seus.
3. Rode `python gerar_planilha.py --dados privado/meus_dados.json --saida Minhas_Contas.xlsx`.

O `.gitignore` impede que `privado/`, `meus_dados*.json` e planilhas `.xlsx` (exceto a de exemplo) sejam enviados ao GitHub.

## Personalizar

Edite o `config.json`:
- `cartao.dia_vencimento` e `cartao.dias_entre_fechamento_e_vencimento`
- `categorias_gastos`, `formas_gastos`, `formas_contas`
- `meses_na_tabela`, `linhas_de_contas`, `linhas_de_gastos`

## Licença

MIT. Pode usar, copiar e adaptar.
