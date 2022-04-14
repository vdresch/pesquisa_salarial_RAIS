# Pesquisa salarial RAIS #
Esse script em R recebe como entrada uma lista de municípios, CNAEs e CBOs, e dá como saída um arquivo em excel com os salários das ocupações.

O modelo do arquivo excel é o seguinte:
- uma aba para cada munícipio pesquisado
- uma linha para cada ocupação pesquisada
- as colunas irão representar três variáveis: tempo de emprego, idade e escolaridade

## Utilização ##
- Para a utilização do script, primeiro de tudo é preciso ter as bibliotecas instaladas. O programa ```install_packages.R``` tem as bibliotecas necessárias, assim como o comando para a instalação.

- O passo seguinte é criar um projeto no Google Cloud para fornecer como credencial para o programa. O script usa a biblioteca da Base dos Dados, que exige tal credencial. Uma vez criado um projeto, o id do projeto deve ser salvo no arquivo ```config.yml```. O email responsável pelo projeto também deverá ser adicionado. É possível ser necessário fornecer permissão para o projeto ser utilizado, mas isso deve ocorrer de forma automática.

- Por fim, o arquivo ```filtros.xlsx``` deve ser editado para conter os municípios desejados, assim como as CBOs desejadas. Esses campos são obrigatórios. O ano de consulta também é obrigatório. O ano de consulta é limitado pela disponibilidade dos dados pelo Ministério da Economia, e geralmente é a partir de dois anos anteriores. O campo de CNAEs é opcional, se o campo estiver em branco será feito uma pesquisa para todas as indústrias. Se estiver preenchido, será feito um filtro por segmento econômico. Um exemplo de como o arquivo ```filtros.xlsx``` deve ser preenchido está no diretório.

Quando os pré requisitos anteriores estiverem completos, deve ser apenas executar o programa e o resultado será automáticamente gerado na mesma pasta.
```
Rscript pesquisa.R
```

## Observações ##
No repositório está presente um arquivo chamado ```lista_cbos.csv```, que foi criado a partir de um PDF do Ministério da Economia em um [outro projeto meu](https://github.com/vdresch/pdf_to_csv_CBOs). Essa lista pode ser atualizada com novas CBOs, então uma atualização periódica da tabela pode ser necessária para o programa funcionar com  as novas ocupações.

Também foi deixado no repositório um programinha que auxilia a visualizar a estrutura do banco de dados da Base dos Dados, auxiliando na modificação do projeto para usos distintos.

## Agradecimentos ##
Esse projeto só foi possível devido ao trabalho da [Base dos Dados](https://basedosdados.org/). Um apoio para esse projeto é sempre legal.
