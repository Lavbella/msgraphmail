import asyncio
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress

async def enviar_email_graph():
    # 1. Configuração das credenciais do Azure AD
    TENANT_ID = ""
    CLIENT_ID = ""
    CLIENT_SECRET = ""
    
    # Conta que vai originar o envio (ex: info@suaempresa.com)
    USER_EMAIL = ""

    # 2. Inicialização do cliente de autenticação e do Graph Client
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    graph_client = GraphServiceClient(credential)

    # 3. Construção da estrutura da mensagem de e-mail
    mensagem = Message(
        subject = "Teste de Envio via Microsoft Graph API",
        body = ItemBody(
            content_type = BodyType.Html,
            content = "<h1>Sucesso</h1><p>Este e-mail foi enviado usando a Graph API em Python.</p>"
        ),
        to_recipients = [
            Recipient(
                email_address = EmailAddress(
                    address = "xxxx17@gmail.com"
                )
            )
        ]
    )

    # 4. Criação do corpo do pedido de envio
    corpo_pedido = SendMailPostRequestBody(
        message = mensagem,
        save_to_sent_items = True
    )

    try:
        # 5. Execução do envio do e-mail
        await graph_client.users.by_user_id(USER_EMAIL).send_mail.post(corpo_pedido)
        print("E-mail enviado com sucesso com a Graph API!")
    except Exception as erro:
        print(f"Erro ao enviar o e-mail: {erro}")
    finally:
        # Garante o fecho correto das ligações assíncronas
        await credential.close()

# Executar a função assíncrona
if __name__ == "__main__":
    asyncio.run(enviar_email_graph())