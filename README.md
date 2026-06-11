# Microsoft Graph API Integration (Mass Mailing / Envio em Massa)

[Português](#português) | [English](#english)

---

## Português

Esta aplicação permite a conectividade segura e autenticada com o ecossistema Microsoft 365 (Outlook) via Tenant para o **envio automatizado de e-mails em massa**. A solução utiliza processamento assíncrono para otimizar a fila de envios e uma interface gráfica moderna desenvolvida com `customtkinter`.

###  Funcionalidades
*   **Autenticação Segura:** Integração com Microsoft Graph API via MSAL (Microsoft Authentication Library).
*   **Envio em Massa Assíncrono:** Arquitetura preparada para processar múltiplos e-mails eficientemente sem bloquear a interface.
*   **Interface Gráfica (GUI):** Interface de utilizador moderna, limpa e responsiva.

###  Limites Importantes do Microsoft 365 (Anti-Spam)
Para evitar que a sua conta seja bloqueada pela Microsoft durante o envio em massa, respeite os seguintes limites padrão por utilizador:
*   **Máximo de destinatários por dia:** 10.000 destinatários.
*   **Máximo de destinatários por mensagem:** 500 destinatários.
*   **Limite de taxa (Rate Limit):** Aproximadamente 30 e-mails por minuto. A aplicação deve implementar pequenos intervalos (*throttling*) entre os lotes.

###  Pré-requisitos (Ubuntu)
Antes de começar, certifique-se de que tem o Python 3 e o suporte para interfaces gráficas instalados no seu sistema:
```bash
sudo apt update
sudo apt install python3-pip python3-venv python3-tk -y
```

###  1. Configurar o Ambiente Virtual (venv)
Siga estes passos para clonar o repositório, criar o ambiente isolado e instalar todas as dependências necessárias:

```bash
# Criar o ambiente virtual
python3 -m venv venv

# Ativar o ambiente virtual
source venv/bin/activate

# Atualizar o pip e instalar as dependências
pip install --upgrade pip
pip install -r requirements.txt
```

###  2. Configuração do Ambiente

Crie um ficheiro chamado `.env` na raiz do projeto. Este ficheiro é obrigatório para que a aplicação consiga autenticar-se no Azure (Microsoft Entra ID) e enviar os e-mails.

| Variável | Descrição | Onde encontrar no Portal do Azure |
| :--- | :--- | :--- |
| `TENANT_ID` | Identificador único da sua organização (Directory ID). | **Visão Geral (Overview)** do seu Microsoft Entra ID. |
| `CLIENT_ID` | Identificador da aplicação registada (Application ID). | **Registos de aplicações (App registrations)** > Selecione a sua App. |
| `CLIENT_SECRET` | Chave secreta gerada para autenticação da aplicação. | Na sua App > **Certificados e segredos (Certificates & secrets)** > Novo segredo. |
| `EMAIL_USER` | Endereço de e-mail da caixa de correio que vai disparar a campanha. | O e-mail institucional ou da conta partilhada (ex: `news@suaempresa.com`). |

Exemplo do ficheiro `.env`:
```env
TENANT_ID=12345678-abcd-1234-abcd-123456789abc
CLIENT_ID=abcdefgh-1234-abcd-1234-abcdefghijkl
CLIENT_SECRET=ExAmPlE_SeCrEt_KeY_DoNoTsHaRe_12345
EMAIL_USER=marketing@suaempresa.com
```

>  **Nota de Segurança:** Nunca envie o ficheiro `.env` para o GitHub. Certifique-se de que o ficheiro `.gitignore` inclui a linha `.env`.

###  3. Executar o Projeto em Desenvolvimento
Com o `venv` ativo, execute o script principal:
```bash
python3 main.py
```

###  4. Como Gerar o Executável (Ubuntu)
Para compilar a aplicação num único ficheiro executável e sem a consola de comandos visível em segundo plano, utilize o PyInstaller:

```bash
pyinstaller --onefile --noconsole main.py
```
O executável final estará disponível na pasta **`dist/main`**. Pode executá-lo diretamente com o comando:
```bash
./dist/app_api
```

---

## English

This application provides secure, authenticated connectivity with the Microsoft 365 ecosystem (Outlook) via Tenant for **automated mass email campaigns**, using an asynchronous workflow to optimize the delivery queue and a modern graphical user interface built with `customtkinter`.

###  Features
*   **Secure Authentication:** Microsoft Graph API integration using MSAL (Microsoft Authentication Library).
*   **Asynchronous Mass Mailing:** Built to process multiple emails efficiently without freezing the user interface.
*   **Graphical Interface (GUI):** Modern and responsive user interface.

###  Important Microsoft 365 Limits (Anti-Spam)
To prevent your account from being flagged or blocked by Microsoft during mass mailing, respect these default limits per user:
*   **Recipient limit per day:** 10,000 recipients.
*   **Recipient limit per message:** 500 recipients.
*   **Rate limit (Throttling):** Approximately 30 emails per minute. The application should handle throttling by adding short delays between batches.

###  Prerequisites (Ubuntu)
Before starting, ensure you have Python 3 and GUI development support installed on your system:
```bash
sudo apt update
sudo apt install python3-pip python3-venv python3-tk -y
```

###  1. Setting Up the Virtual Environment (venv)
Follow these steps to clone the repository, create the isolated environment, and install all required dependencies:

```bash
# Create the virtual environment
python3 -m venv venv

# Activate the virtual environment
source venv/bin/activate

# Upgrade pip and install dependencies
pip install --upgrade pip
pip install -r requirements.txt
```

###  2. Environment Configuration

Create a file named `.env` in the root directory of the project. This file is mandatory for the application to authenticate with Azure (Microsoft Entra ID) and send emails.

| Variable | Description | Where to find in the Azure Portal |
| :--- | :--- | :--- |
| `TENANT_ID` | Unique identifier of your organization (Directory ID). | **Overview** page of your Microsoft Entra ID. |
| `CLIENT_ID` | Identifier of your registered application (Application ID). | **App registrations** > Select your registered App. |
| `CLIENT_SECRET` | Secret key generated for application authentication. | Inside your App > **Certificates & secrets** > New client secret. |
| `EMAIL_USER` | Email address of the mailbox that will send the messages. | The corporate email or shared mailbox address (e.g., `news@yourcompany.com`). |

Example `.env` file:
```env
TENANT_ID=12345678-abcd-1234-abcd-123456789abc
CLIENT_ID=abcdefgh-1234-abcd-1234-abcdefghijkl
CLIENT_SECRET=ExAmPlE_SeCrEt_KeY_DoNoTsHaRe_12345
EMAIL_USER=marketing@suaempresa.com
```

>  **Security Note:** Never commit your `.env` file to GitHub. Ensure your `.gitignore` file contains a `.env` entry.

###  3. Running the Project in Development
With the `venv` active, run the main script:
```bash
python3 main.py
```

###  4. How to Generate the Executable (Ubuntu)
To compile the application into a single standalone executable file without the background terminal console, use PyInstaller:

```bash
pyinstaller --onefile --noconsole main.py
```
The final executable will be available inside the **`dist/main`** folder. You can run it directly using:
```bash
./dist/app_api
```
## Licença / License

### Português
Este projeto está licenciado sob a **GNU GPLv3**. Consulte o ficheiro [LICENSE](LICENSE) para obter mais detalhes.

### English
This project is licensed under the **GNU GPLv3**. See the [LICENSE](LICENSE) file for more details.
