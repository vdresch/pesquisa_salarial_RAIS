###########################         Librarys            ###########################

library(basedosdados)
library(yaml)
library(readxl)
library(openxlsx)
library(dplyr)

###########################           Config            ###########################

# Configura????o dos munic??pios, CNAEs e CBOs desejadas na pesquisa
config_filters <- read_excel("filtros.xlsx", col_types = c("numeric", "numeric", "numeric", "numeric"))

municipios <- config_filters$cod_mun
cnaes <- config_filters$cnae
cbos <- config_filters$cbo
ano_consulta <- config_filters$ano

# billing_id necess??rio para rodar querys no basedosdados
config <- yaml.load_file("config.yml")
basedosdados::set_billing_id(config$billing_id)
bq_auth(email = config$email)

#Tabela de suporte com nomes dos munic??pios e estados
tab_mun <- read_excel("suporte/municipios.xlsx")
tab_uf <- read_excel("suporte/ufs.xlsx")

###########################           Get data            ###########################

#Preparing the query

query <- "SELECT id_municipio_trabalho, cbo_2002, cnae_2, valor_remuneracao_media, tempo_emprego, faixa_etaria, escolaridade_apos_2005 FROM `basedosdados.br_me_rais.microdados_vinculos` WHERE "

#Municipios
if(length(municipios) == 1) {
    query_mun <- paste("id_municipio_trabalho = \"", municipios, "\"",sep="")
} else {
    query_mun <- "id_municipio_trabalho in ("
    for(i in municipios) {
        query_mun <- paste(query_mun, ", \"", i, "\"", sep="")
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
        query_cbos <- paste(query_cbos, ", \"", i, "\"", sep="")
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
query <- paste(query, query_mun, query_cbos, query_cnae, " and ano = ", ano_consulta[1], sep="")

#Executa query
jobs <- read_sql(query)


###########################           To xlsx            ###########################

wb <- createWorkbook()

#Pega o nome das cidades para a planilha
municipios <- data.frame("cod_mun" = config_filters$cod_mun)
municipios <- merge(x=municipios, y=tab_mun, by.x='cod_mun', by.y='cod_mun', all.x = TRUE)
municipios <- merge(x=municipios, y=tab_uf, by.x='UF', by.y='cod_uf', all.x = TRUE)

#Para cada cidade, uma aba
for(i in 1:nrow(municipios)) {
    #Cria df e aba que ser??o salvos
    tabela_mun <- data.frame()
    addWorksheet(wb, paste(municipios[i, "nome_mun"], municipios[i, "sigla"], sep=" - "))
  
    #Filtra cidade
    jobs_municipio <- jobs %>% filter(id_municipio_trabalho == municipios[i, "cod_mun"])
    jobs_municipio <- jobs_municipio[jobs_municipio$valor_remuneracao_media != 0, ]

    #Para cada CBO, um linha
    for(j in  1:length(cbos)) {
        jobs_cbos <- jobs_municipio %>% filter(cbo_2002 == cbos[j])
        
        tabela <- data.frame()
        tabela[1, 'cbo'] <- cbos[j]
        #tabela['cbo_nome'] <- cbos[i]
        tabela[1, 'total'] <- lengths(jobs_cbos)[1]
        tabela[2, 'total'] <- mean(jobs_cbos$valor_remuneracao_media)
        tabela[1, 'faixa1'] <- table(jobs_cbos$faixa_etaria == 1)[2]
        tabela[1, 'faixa2'] <- table(jobs_cbos$faixa_etaria == 2)[2]
        tabela[1, 'faixa3'] <- table(jobs_cbos$faixa_etaria == 3)[2]
        tabela[1, 'faixa4'] <- table(jobs_cbos$faixa_etaria == 4)[2]
        tabela[1, 'faixa5'] <- table(jobs_cbos$faixa_etaria == 5)[2]
        tabela[1, 'faixa6'] <- table(jobs_cbos$faixa_etaria == 6)[2]
        tabela[1, 'faixa7'] <- table(jobs_cbos$faixa_etaria == 7)[2]
        tabela[1, 'faixa8'] <- table(jobs_cbos$faixa_etaria == 8)[2]
        tabela[2, 'faixa1'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 1, ]$valor_remuneracao_media)
        tabela[2, 'faixa2'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 2, ]$valor_remuneracao_media)
        tabela[2, 'faixa3'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 3, ]$valor_remuneracao_media)
        tabela[2, 'faixa4'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 4, ]$valor_remuneracao_media)
        tabela[2, 'faixa5'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 5, ]$valor_remuneracao_media)
        tabela[2, 'faixa6'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 6, ]$valor_remuneracao_media)
        tabela[2, 'faixa7'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 7, ]$valor_remuneracao_media)
        tabela[2, 'faixa8'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 8, ]$valor_remuneracao_media)
        #tabela[1, 'escol1'] <- table(jobs_cbos$escolaridade_apos_2005 == 1)[2]
        #tabela[1, 'escol2'] <- table(jobs_cbos$escolaridade_apos_2005 == 2)[2]
        #tabela[1, 'escol3'] <- table(jobs_cbos$escolaridade_apos_2005 == 3)[2]
        #tabela[1, 'escol4'] <- table(jobs_cbos$escolaridade_apos_2005 == 4)[2]
        #tabela[1, 'escol5'] <- table(jobs_cbos$escolaridade_apos_2005 == 5)[2]
        #tabela[1, 'escol6'] <- table(jobs_cbos$escolaridade_apos_2005 == 6)[2]
        #tabela[1, 'escol7'] <- table(jobs_cbos$escolaridade_apos_2005 == 7)[2]
        #tabela[1, 'escol8'] <- table(jobs_cbos$escolaridade_apos_2005 == 8)[2]
        #tabela[1, 'escol9'] <- table(jobs_cbos$escolaridade_apos_2005 == 9)[2]
        #tabela[1, 'escol10'] <- table(jobs_cbos$escolaridade_apos_2005 == 10)[2]
        #tabela[1, 'escol11'] <- table(jobs_cbos$escolaridade_apos_2005 == 11)[2]
        #tabela[2, 'escol1'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 1, ]$valor_remuneracao_media)
        #tabela[2, 'escol2'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 2, ]$valor_remuneracao_media)
        #tabela[2, 'escol3'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 3, ]$valor_remuneracao_media)
        #tabela[2, 'escol4'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 4, ]$valor_remuneracao_media)
        #tabela[2, 'escol5'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 5, ]$valor_remuneracao_media)
        #tabela[2, 'escol6'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 6, ]$valor_remuneracao_media)
        #tabela[2, 'escol7'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 7, ]$valor_remuneracao_media)
        #tabela[2, 'escol8'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 8, ]$valor_remuneracao_media)
        #tabela[2, 'escol9'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 9, ]$valor_remuneracao_media)
        #tabela[2, 'escol10'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 10, ]$valor_remuneracao_media)
        #tabela[2, 'escol11'] <- mean(jobs_cbos[jobs_cbos$escolaridade_apos_2005 == 11, ]$valor_remuneracao_media)
        #tabela[1, 'tempo1'] <- table(jobs_cbos$tempo_emprego == 1)[2]
        #tabela[1, 'tempo2'] <- table(jobs_cbos$tempo_emprego == 2)[2]
        #tabela[1, 'tempo3'] <- table(jobs_cbos$tempo_emprego == 3)[2]
        #tabela[1, 'tempo4'] <- table(jobs_cbos$tempo_emprego == 4)[2]
        #tabela[1, 'tempo5'] <- table(jobs_cbos$tempo_emprego == 5)[2]
        #tabela[1, 'tempo6'] <- table(jobs_cbos$tempo_emprego == 6)[2]
        #tabela[1, 'tempo7'] <- table(jobs_cbos$tempo_emprego == 7)[2]
        #tabela[1, 'tempo8'] <- table(jobs_cbos$tempo_emprego == 8)[2]
        #tabela[2, 'tempo1'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 1, ]$valor_remuneracao_media)
        #tabela[2, 'tempo2'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 2, ]$valor_remuneracao_media)
        #tabela[2, 'tempo3'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 3, ]$valor_remuneracao_media)
        #tabela[2, 'tempo4'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 4, ]$valor_remuneracao_media)
        #tabela[2, 'tempo5'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 5, ]$valor_remuneracao_media)
        #tabela[2, 'tempo6'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 6, ]$valor_remuneracao_media)
        #tabela[2, 'tempo7'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 7, ]$valor_remuneracao_media)
        #tabela[2, 'tempo8'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 8, ]$valor_remuneracao_media)
        
        tabela_mun <- rbind(tabela_mun, tabela)
    }
    
    #Salva aba do municipio
    writeData(wb, paste(municipios[i, "nome_mun"], municipios[i, "sigla"], sep=" - "), tabela_mun, startRow = 1, startCol = 1)
}

#Salva planilha
saveWorkbook(wb, file = "resultado.xlsx", overwrite = TRUE)
