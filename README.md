# Feed do catálogo — Rede Central Farma

Arquivo lido automaticamente pelo **Meta Commerce Manager** para alimentar o catálogo de
produtos usado no atendimento por WhatsApp das lojas.

- **Catálogo:** "Rede Central Farma - Catalogo Lojas" (ID `964862383322100`)
- **Portfólio empresarial:** Rede Central Farma (`807408591759806`)
- **URL do feed:** `https://raw.githubusercontent.com/redecentralauditoria-cmyk/feed-catalogo-lojas/main/catalogo.csv`

## O que tem no arquivo

Produtos **não medicamentosos** (grupo HPC do cadastro FarmaPRO) com giro na Loja 12.
Medicamento não entra no catálogo.

| Coluna | Origem |
|---|---|
| `id` | código do produto no FarmaPRO |
| `title` / `description` | nome do produto no cadastro da loja online (Instabuy) |
| `price` | P.M.C. — preço **já com o desconto do cadastro Fidelidade** |
| `image_link` | foto pública do cadastro da loja online |
| `link` | loja online |
| `brand` | marca do cadastro da loja online |

## Como atualizar

Gerar o CSV novamente a partir das planilhas do FarmaPRO, substituir `catalogo.csv` e:

```
git add catalogo.csv && git commit -m "atualiza precos" && git push
```

A Meta relê o arquivo no horário agendado em Commerce Manager → Catálogo → Fontes de dados.
