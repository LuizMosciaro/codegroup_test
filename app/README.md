# Teste CodeGroup

Teste realizado para avaliacao de dev python

## Instalação

1. Clone o repositório para o seu ambiente local.
2. Navegue até o diretório do projeto.
3. Crie e ative um ambiente virtual.
4. Instale as dependências usando o arquivo `requirements.txt`:
    pip install -r requirements.txt
5. Renomeie o arquivo `.env.example` para `.env` e preencha as credenciais necessárias:
    WEATHER_API_KEY=<sua_chave_de_api>
    EMAIL=<seu_email_e_linkedin>
    EMAIL_PWD=<sua_senha_do_email>
    PASSWORD=<sua_senha_linkedin>
    RECIPIENT=<email_destinarario>
6. Execute o projeto.

## Utilização
O projeto tentara realizar login no linkedin, depois pegara informacoes dos usuarios listados na planilha e finalmente enviara um email para o destinatario com as informacoes coletadas na planilha como anexo