###########################         Librarys            ###########################

library("basedosdados")
library("yaml")
library("readxl")

###########################           Config            ###########################

# Configuração dos municípios, CNAEs e CBOs desejadas na pesquisa
config_filters <- read_excel("filtros.xlsx")

municipios <- config_filters$cod_mun
cnaes <- config_filters$cnae
cbos <- config_filters$cbo
ano_consulta <- config_filters$ano

# billing_id necessário para rodar querys no basedosdados
config <- yaml.load_file("config.yml")
basedosdados::set_billing_id(config$billing_id)

###########################           Get data            ###########################

#Preparing the query

query <- "SELECT id_municipio_trabalho, cbo_2002, cnae_2, valor_remuneracao_media, faixa_etaria, grau_instrucao_apos_2005 FROM `basedosdados.br_me_rais.microdados_vinculos` WHERE "

#Municipios
if(length(municipios) == 1) {
    query_mun <- paste("id_municipio_trabalho = \"", municipios, "\"",sep="")
} else {
    query_mun <- "id_municipio_trabalho in ("
    for(i in municipios) {
        query_mun <- paste(query_mun, "\"", i, "\"", sep=", ")
    }
    query_mun = gsub("\\(, ", "(", query_mun)
    query_mun <- paste(query_mun, ")", sep="")
}
#CBOs
if(length(cbos) == 1) {
    query_cbos <- paste(" and cbo_2002 = \"", cbos, "\"", sep="")
} else {
    query_cbos <- " and cbo_2002 in ("
    for(i in cbos) {
        query_cbos <- paste(query_cbos, "\"", i, "\"", sep=", ")
    }
    query_cbos = gsub("\\(, ", "(", query_cbos)
    query_cbos <- paste(query_cbos, ")", sep="")
}
#CNAEs
if(!is.na(cnaes)){
    if(length(cnaes) == 1) {
        query_cnaes <- paste(" and cnae_2 = ", cnaes, sep="")
    } else {
        query_cnaes <- " and cnae_2 in ("
        for(i in cnaes) {
            query_cnaes <- paste(query_cnaes, i, sep=", ")
        }
        query_cnaes = gsub("\\(, ", "(", query_cnaes)
        query_cnaes <- paste(query_cnaes, ")", sep="")
    }
} else {query_cnae <- ""}

#Query final
query <- paste(query, query_mun, query_cbos, query_cnae, " and ano = ", ano_consulta, sep="")

#Executa query
jobs <- read_sql(query)