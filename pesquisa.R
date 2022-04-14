setwd(dirname(rstudioapi::getSourceEditorContext()$path))

###########################         Librarys            ###########################

library(basedosdados)
library(yaml)
library(readxl)
library(openxlsx)
library(dplyr)
library(bigrquery)

###########################           Config            ###########################

# Configuração dos municípios, CNAEs e CBOs desejadas na pesquisa
config_filters <- read_excel("filtros.xlsx", col_types = c("numeric", "numeric", "numeric", "numeric"))

municipios <- config_filters$cod_mun
cnaes <- config_filters$cnae
cbos <- config_filters$cbo
ano_consulta <- config_filters$ano

# billing_id necessário para rodar querys no basedosdados
config <- yaml.load_file("config.yml")
basedosdados::set_billing_id(config$billing_id)
bq_auth(email = config$email)

#Tabela de suporte com nomes dos municípios, estados e CBOS
tab_mun <- read_excel("suporte/municipios.xlsx")
tab_uf <- read_excel("suporte/ufs.xlsx")
tab_cbos <- read.csv("suporte/lista_cbos.csv", encoding = "UTF-8")

###########################           Get data            ##########################

#Preparing the query

query <- "SELECT id_municipio_trabalho, cbo_2002, cnae_2, valor_remuneracao_media, tempo_emprego, faixa_etaria, grau_instrucao_apos_2005 FROM `basedosdados.br_me_rais.microdados_vinculos` WHERE "

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
        query_cnaes <- paste(" and cnae_2 between \"", cnaes, "000\" and \"", cnaes, "999\"", sep="")
    } else {
        query_cnaes <- " and ((cnae_2 between \""
        for(i in cnaes) {
            query_cnaes <- paste(query_cnaes, i, "000\" and \"", i, "999\" ", sep="")
        }
        query_cnaes <- paste(query_cnaes, "x)", sep="")
        query_cnaes = gsub("9\" ", "9\") or (cnae_2 between \"", query_cnaes)
        query_cnaes = gsub("9\"\\) or \\(cnae_2 between \"x\\)", "9\")) ", query_cnaes)
    }
} else {query_cnaes <- ""}

#Query final
query <- paste(query, query_mun, query_cbos, query_cnaes, " and ano = ", ano_consulta[1], sep="")

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
    #Cria df e aba que serão salvos
    tabela_mun <- data.frame()
    addWorksheet(wb, paste(municipios[i, "nome_mun"], municipios[i, "sigla"], sep=" - "))
  
    #Filtra cidade
    jobs_municipio <- jobs %>% filter(id_municipio_trabalho == municipios[i, "cod_mun"])
    jobs_municipio <- jobs_municipio[jobs_municipio$valor_remuneracao_media != 0, ]

    #Para cada CBO, um linha
    for(j in  1:length(cbos)) {
        jobs_cbos <- jobs_municipio %>% filter(cbo_2002 == cbos[j])
        tabela <- data.frame()
        tabela[1, 'CBO'] <- cbos[j]
        tabela[2, 'CBO'] <- ""
        cbo_i <- subset(tab_cbos, CBO == cbos[j])
        tabela[1, 'cbo_nome'] <- cbo_i$Nome[1]
        tabela[2, 'cbo_nome'] <- ""
        tabela[1, 'Total de trabalhadores'] <- lengths(jobs_cbos)[1]
        tabela[2, 'Total de trabalhadores'] <- mean(jobs_cbos$valor_remuneracao_media)
        tabela[1, '10 a 14 anos'] <- table(jobs_cbos$faixa_etaria == 1)[2]
        tabela[1, '15 a 17 anos'] <- table(jobs_cbos$faixa_etaria == 2)[2]
        tabela[1, '18 a 24 anos'] <- table(jobs_cbos$faixa_etaria == 3)[2]
        tabela[1, '25 a 29 anos'] <- table(jobs_cbos$faixa_etaria == 4)[2]
        tabela[1, '30 a 39 anos'] <- table(jobs_cbos$faixa_etaria == 5)[2]
        tabela[1, '40 a 49 anos'] <- table(jobs_cbos$faixa_etaria == 6)[2]
        tabela[1, '50 a 64 anos'] <- table(jobs_cbos$faixa_etaria == 7)[2]
        tabela[1, '65 anos ou mais'] <- table(jobs_cbos$faixa_etaria == 8)[2]
        tabela[2, '10 a 14 anos'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 1, ]$valor_remuneracao_media)
        tabela[2, '15 a 17 anos'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 2, ]$valor_remuneracao_media)
        tabela[2, '18 a 24 anos'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 3, ]$valor_remuneracao_media)
        tabela[2, '25 a 29 anos'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 4, ]$valor_remuneracao_media)
        tabela[2, '30 a 39 anos'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 5, ]$valor_remuneracao_media)
        tabela[2, '40 a 49 anos'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 6, ]$valor_remuneracao_media)
        tabela[2, '50 a 64 anos'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 7, ]$valor_remuneracao_media)
        tabela[2, '65 anos ou mais'] <- mean(jobs_cbos[jobs_cbos$faixa_etaria == 8, ]$valor_remuneracao_media)
        tabela[1, 'Analfabeto'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 1)[2]
        tabela[1, 'Até 5º ano'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 2)[2]
        tabela[1, '5º ano'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 3)[2]
        tabela[1, '6º ao 9º ano'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 4)[2]
        tabela[1, 'Fundamental Completo'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 5)[2]
        tabela[1, 'Médio Incompleto'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 6)[2]
        tabela[1, 'Médio Completo'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 7)[2]
        tabela[1, 'Superior Incompleto'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 8)[2]
        tabela[1, 'Superior Completo'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 9)[2]
        tabela[1, 'Mestrado'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 10)[2]
        tabela[1, 'Doutorado'] <- table(jobs_cbos$grau_instrucao_apos_2005 == 11)[2]
        tabela[2, 'Analfabeto'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 1, ]$valor_remuneracao_media)
        tabela[2, 'Até 5º ano'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 2, ]$valor_remuneracao_media)
        tabela[2, '5º ano'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 3, ]$valor_remuneracao_media)
        tabela[2, '6º ao 9º ano'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 4, ]$valor_remuneracao_media)
        tabela[2, 'Fundamental Completo'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 5, ]$valor_remuneracao_media)
        tabela[2, 'Médio Incompleto'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 6, ]$valor_remuneracao_media)
        tabela[2, 'Médio Completo'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 7, ]$valor_remuneracao_media)
        tabela[2, 'Superior Incompleto'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 8, ]$valor_remuneracao_media)
        tabela[2, 'Superior Completo'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 9, ]$valor_remuneracao_media)
        tabela[2, 'Mestrado'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 10, ]$valor_remuneracao_media)
        tabela[2, 'Doutorado'] <- mean(jobs_cbos[jobs_cbos$grau_instrucao_apos_2005 == 11, ]$valor_remuneracao_media)
        tabela[1, 'Ate 2,9 meses'] <- table(jobs_cbos$tempo_emprego == 1)[2]
        tabela[1, '3,0 a 5,9 meses'] <- table(jobs_cbos$tempo_emprego == 2)[2]
        tabela[1, '6,0 a 11,9 meses'] <- table(jobs_cbos$tempo_emprego == 3)[2]
        tabela[1, '12,0 a 23,9 meses'] <- table(jobs_cbos$tempo_emprego == 4)[2]
        tabela[1, '24,0 a 35,9 meses'] <- table(jobs_cbos$tempo_emprego == 5)[2]
        tabela[1, '36,0 a 59,9 meses'] <- table(jobs_cbos$tempo_emprego == 6)[2]
        tabela[1, '60,0 a 119,9 meses'] <- table(jobs_cbos$tempo_emprego == 7)[2]
        tabela[1, '120,0 meses ou mais'] <- table(jobs_cbos$tempo_emprego == 8)[2]
        tabela[2, 'Ate 2,9 meses'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 1, ]$valor_remuneracao_media)
        tabela[2, '3,0 a 5,9 meses'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 2, ]$valor_remuneracao_media)
        tabela[2, '6,0 a 11,9 meses'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 3, ]$valor_remuneracao_media)
        tabela[2, '12,0 a 23,9 meses'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 4, ]$valor_remuneracao_media)
        tabela[2, '24,0 a 35,9 meses'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 5, ]$valor_remuneracao_media)
        tabela[2, '36,0 a 59,9 meses'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 6, ]$valor_remuneracao_media)
        tabela[2, '60,0 a 119,9 meses'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 7, ]$valor_remuneracao_media)
        tabela[2, '120,0 meses ou mais'] <- mean(jobs_cbos[jobs_cbos$tempo_emprego == 8, ]$valor_remuneracao_media)
        tabela <- tabela %>% replace(is.na(.), 0)
        
        tabela_mun <- rbind(tabela_mun, tabela)
        
        #Merge cells
        row_begin = 2*j
        row_end = (2*j)+1
        mergeCells(wb, paste(municipios[i, "nome_mun"], municipios[i, "sigla"], sep=" - "), cols = 1:1, rows = row_begin:row_end)
        mergeCells(wb, paste(municipios[i, "nome_mun"], municipios[i, "sigla"], sep=" - "), cols = 2:2, rows = row_begin:row_end)

    }
    
    #Salva aba do municipio
    writeData(wb, paste(municipios[i, "nome_mun"], municipios[i, "sigla"], sep=" - "), tabela_mun, startRow = 1, startCol = 1)
    
    
}

#Salva planilha
saveWorkbook(wb, file = "Results/resultado.xlsx", overwrite = TRUE)
