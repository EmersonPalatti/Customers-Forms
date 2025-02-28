# Formulário de Cadastro - Streamlit
Este projeto é uma aplicação web construída com Streamlit que permite aos usuários preencher um formulário de cadastro detalhado, gerar arquivos Excel baseados em templates pré-definidos (Merck e Sigma) e enviar esses arquivos por e-mail junto com anexos adicionais. A aplicação é útil para processos de cadastro de empresas ou pessoas físicas, incluindo informações fiscais, de endereço, contatos e comprovantes.

https://registration-form-merck-sigma.streamlit.app/

## Funcionalidades
- Formulário Interativo: Interface amigável para inserção de dados como razão social, CNPJ, endereço, informações fiscais e comprovantes.

- Geração de Arquivos Excel: Preenche automaticamente templates Excel (`Merck.xlsx` e `Sigma.xlsx`) com os dados fornecidos.

- Suporte a Imagens: Permite upload de comprovantes (endereço, Receita Federal, Sintegra, etc.) que são inseridos nos arquivos Excel.

- Envio por E-mail: Envia os arquivos gerados e comprovantes adicionais por e-mail usando SMTP (compatível com Gmail).

- Validação Básica: Verifica se o campo "Razão Social" foi preenchido antes de processar o envio.

- Feedback ao Usuário: Exibe mensagens de sucesso ou erro e inclui animações (balões) após o envio bem-sucedido.

## Estrutura do Projeto
```
├── templates/
│   ├── Merck.xlsx        # Template Excel para Merck
│   ├── Sigma.xlsx        # Template Excel para Sigma
├── merck1.jpg            # Logo exibido na interface
├── app.py                # Código principal da aplicação
├── README.md             # Este arquivo
└── requirements.txt      # Dependências do projeto
```
## Pré-requisitos
- Python 3.8 ou superior

- Conta de e-mail Gmail com "Senha de Aplicativo" para envio de e-mails (devido à autenticação SMTP)

- Arquivos Excel de template (`Merck.xlsx` e `Sigma.xlsx`) na pasta templates/

- Imagem do logo (`merck1.jpg`) no diretório raiz

## Instalação
1. Clone o repositório:
```bash

git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
```
2. Crie um ambiente virtual (opcional, mas recomendado):
```bash

python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```
3. Instale as dependências:
Crie um arquivo `requirements.txt` com o seguinte conteúdo:
```
streamlit
pandas
openpyxl
pillow
```
Em seguida, execute:
```bash

pip install -r requirements.txt
```
4. Configure as credenciais de e-mail:
Crie um arquivo .streamlit/secrets.toml no diretório do projeto com as seguintes variáveis:
```toml

SENDER_EMAIL = "seu-email@gmail.com"
RECEIVER_EMAIL = "destinatario@example.com"
EMAIL_PASSWORD = "sua-senha-de-aplicativo"
```
Nota: Para o Gmail, gere uma "Senha de Aplicativo" em Gerenciar Conta Google > Segurança > Senhas de aplicativo.

5. Adicione os templates e logo:
- Coloque os arquivos `Merck.xlsx` e `Sigma.xlsx` na pasta templates/.

- Coloque a imagem `merck1.jpg` no diretório raiz.

## Uso
1. Execute a aplicação:
```bash

streamlit run app.py
```
2. Preencha o formulário:
- Acesse a aplicação no navegador (geralmente em http://localhost:8501).

- Preencha os campos necessários nos expanders: Dados de Cadastro, Endereço, Informações de Contato, etc.

- Faça upload dos comprovantes nos formatos .jpg, .jpeg ou .png.

3. Envie os dados:
- Clique no botão "Enviar".

- Aguarde o processamento (os arquivos Excel serão gerados e enviados por e-mail).

- Em caso de sucesso, uma mensagem personalizada com um cupom fictício será exibida.

## Estrutura do Formulário
- Dados de Cadastro: Razão social, CNPJ/CPF, inscrição estadual, telefone, e-mail, etc.

- Endereço: Endereço de faturamento e opcionalmente de entrega, com complementos (universidade, laboratório, etc.).

- Informações de Contato: Dados do solicitante.

- Informações de Contribuição: Tipo de empresa, uso dos produtos, incentivos fiscais (ICMS, IPI, PIS, COFINS).

- Empresas Coligadas: Até 4 empresas associadas (opcional).

- Comprovantes: Upload de documentos como comprovante de endereço e cartão da Receita Federal.

## Notas Técnicas
- Templates Excel: Os arquivos Merck.xlsx e Sigma.xlsx devem ter as abas "FICHA CADASTRAL (Sold-to)" e "FICHA CADASTRAL (Ship-to)" com as células mapeadas conforme o dicionário cells_sold_to e cells_ship_to.

- Imagens: Os comprovantes são inseridos nas células especificadas e removidos após o processamento para evitar acúmulo de arquivos temporários.

- Segurança: As credenciais de e-mail são armazenadas em secrets.toml para evitar exposição no código.

## Contribuição
1. Faça um fork do repositório.

2. Crie uma branch para sua feature (git checkout -b feature/nova-funcionalidade).

3. Commit suas alterações (git commit -m 'Adiciona nova funcionalidade').

4. Push para o repositório remoto (git push origin feature/nova-funcionalidade).

5. Abra um Pull Request.

