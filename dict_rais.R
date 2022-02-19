library("basedosdados")
# Defina o seu projeto no Google Cloud
set_billing_id("focus-tree-338418")
# Para carregar o dado direto no R
query <- bdplyr("br_me_rais.dicionario")
df <- bd_collect(query)