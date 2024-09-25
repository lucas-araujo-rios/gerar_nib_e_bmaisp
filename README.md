# gerar_nib_e_bmaisp

Código que puxa bases de dados da Embrapii, filtra para NIB e B+P, e exporta as planilhas filtradas.

## Como utilizar

Por padrão, todos os arquivos serão baixados e exportados para a pasta onde está localizado o projeto. 
Logo, basta fazer os passos:
1. Criar em seu dispositivo a pasta do projeto
2. Clonar o repositório: `git clone https://github.com/lucas-araujo-rios/gerar_nib_e_bmaisp.git`
3. Certifique-se de que todos os pacotes necessários (pandas, ...) estão instalados
4. Rodar os arquivos .py em ordem, dado um período de referência


## Requisitos

É necessário criar o arquivo .env, com padrão:

```env
PASTA_DOWNLOAD=<pasta para onde vão os seus downloads>
ROOT=<pasta onde está localizado o projeto>
 
sharepoint_email=<seu e-mail da embrapii>
sharepoint_password=<sua senha>
sharepoint_url_site="https://embrapii.sharepoint.com/sites/GEEDD"
sharepoint_site_name="GEEDD"
sharepoint_doc_library="Documentos Compartilhados/"
```
