import os
import asyncio
import re
import base64
import mimetypes
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
import asyncio  # Necessário para o temporizador sleep

# Importações oficiais do Microsoft Graph SDK
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.attachment import Attachment
from msgraph.generated.models.file_attachment import FileAttachment

load_dotenv()

# Configuração estética do sistema
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class InterfaceModernaApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Microsoft Graph API Mailer Pro")
        self.geometry("900x650")
        self.resizable(False, False)
        
        self.caminhos_anexos = []
        self.caminho_excel = None
        
        # Credenciais lidas do ficheiro .env
        self.tenant_id = os.getenv("TENANT_ID")
        self.client_id = os.getenv("CLIENT_ID")        # Application (client) ID
        self.client_secret = os.getenv("CLIENT_SECRET")  # Valor do Segredo (Value)
        self.usuario = os.getenv("MS_USER")                  # E-mail remetente da organização

        # Verificação da integridade das chaves
        self.env_valido = all([self.tenant_id, self.client_id, self.client_secret, self.usuario])

        # --- ESTRUTURA DE LAYOUT (Grelha Principal) ---
        self.grid_columnconfigure(0, weight=1, minsize=380) 
        self.grid_columnconfigure(1, weight=2, minsize=520) 
        self.grid_rowconfigure(0, weight=1)

        # ==========================================
        # PAINEL ESQUERDO: Configurações e Destino
        # ==========================================
        self.frame_esquerda = ctk.CTkFrame(self, corner_radius=15, fg_color="#1e1e24")
        self.frame_esquerda.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        
        self.lbl_logo = ctk.CTkLabel(self.frame_esquerda, text="🚀 Graph API Mailer", font=ctk.CTkFont(size=24, weight="bold"))
        self.lbl_logo.pack(pady=(25, 5), padx=20, anchor="w")
        
        status_conta = f"🔒 Conectado: {self.usuario}" if self.env_valido else "❌ Falha nas chaves do .env"
        cor_conta = "#2ecc71" if self.env_valido else "#e74c3c"
        self.lbl_auth = ctk.CTkLabel(self.frame_esquerda, text=status_conta, text_color=cor_conta, font=ctk.CTkFont(size=12))
        self.lbl_auth.pack(pady=(0, 20), padx=20, anchor="w")

        self.card_destino = ctk.CTkFrame(self.frame_esquerda, fg_color="#2a2a35", corner_radius=10)
        self.card_destino.pack(fill="x", padx=20, pady=10)
        
        self.lbl_card_tit = ctk.CTkLabel(self.card_destino, text="Destinatários", font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_card_tit.pack(pady=(10, 5), padx=15, anchor="w")

        self.switch_excel = ctk.CTkSwitch(self.card_destino, text="Enviar para Lista Excel", command=self.alternar_modo, progress_color="#3b82f6")
        self.switch_excel.pack(pady=10, padx=15, anchor="w")

        self.entry_para = ctk.CTkEntry(self.card_destino, placeholder_text="Destinatário (email individual)", height=40, border_color="#47475a")
        self.entry_para.pack(fill="x", padx=15, pady=(5, 15))

        self.btn_excel = ctk.CTkButton(self.card_destino, text="📊 Importar Ficheiro Excel", command=self.selecionar_excel, fg_color="#6366f1", hover_color="#4f46e5", height=35)
        self.lbl_excel = ctk.CTkLabel(self.card_destino, text="Nenhum Excel selecionado", text_color="#a1a1aa", font=ctk.CTkFont(size=11))

        self.card_anexos = ctk.CTkFrame(self.frame_esquerda, fg_color="#2a2a35", corner_radius=10)
        self.card_anexos.pack(fill="x", padx=20, pady=10)
        
        self.btn_anexar = ctk.CTkButton(self.card_anexos, text="📎 Adicionar Anexos", command=self.selecionar_anexos, fg_color="#4b5563", hover_color="#374151")
        self.btn_anexar.pack(fill="x", padx=15, pady=(15, 5))
        
        self.lbl_anexos = ctk.CTkLabel(self.card_anexos, text="Nenhum ficheiro anexado", text_color="#a1a1aa", font=ctk.CTkFont(size=11), wraplength=300)
        self.lbl_anexos.pack(pady=(0, 15), padx=15)

        self.btn_enviar = ctk.CTkButton(self.frame_esquerda, text="Iniciar o envio", command=self.processar_envio, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#22c55e", hover_color="#16a34a", height=50, corner_radius=8)
        self.btn_enviar.pack(fill="x", padx=20, pady=(30, 10), side="bottom")

        self.lbl_status = ctk.CTkLabel(self.frame_esquerda, text="Pronto para enviar", text_color="#a1a1aa", font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_status.pack(pady=5, side="bottom")

        # ==========================================
        # PAINEL DIREITO: Editor de Conteúdo HTML
        # ==========================================
        self.frame_direita = ctk.CTkFrame(self, corner_radius=15, fg_color="#18181b")
        self.frame_direita.grid(row=0, column=1, padx=(0, 15), pady=15, sticky="nsew")

        self.lbl_assunto = ctk.CTkLabel(self.frame_direita, text="Assunto da Mensagem", font=ctk.CTkFont(size=13, weight="bold"), text_color="#d4d4d8")
        self.lbl_assunto.pack(pady=(20, 5), padx=25, anchor="w")

        self.entry_assunto = ctk.CTkEntry(self.frame_direita, placeholder_text="Escreva o assunto aqui...", height=40, border_color="#27272a", fg_color="#09090b")
        self.entry_assunto.pack(fill="x", padx=25, pady=(0, 15))

        self.lbl_corpo = ctk.CTkLabel(self.frame_direita, text="Conteúdo HTML do Email", font=ctk.CTkFont(size=13, weight="bold"), text_color="#d4d4d8")
        self.lbl_corpo.pack(pady=(5, 5), padx=25, anchor="w")

        self.txt_mensagem = ctk.CTkTextbox(self.frame_direita, border_color="#27272a", fg_color="#09090b", font=ctk.CTkFont(family="Consolas", size=12))
        self.txt_mensagem.pack(fill="both", expand=True, padx=25, pady=(0, 25))
        self.txt_mensagem.insert("0.0", "<!-- Código HTML Fixo ou Dinâmico -->\n<h1>Olá {{Nome}}</h1>\n<p>O seu email de teste formatado.</p>")

    def alternar_modo(self):
        if self.switch_excel.get() == 1:
            self.entry_para.pack_forget()
            self.btn_excel.pack(fill="x", padx=15, pady=(5, 2))
            self.lbl_excel.pack(pady=(0, 15), padx=15)
        else:
            self.btn_excel.pack_forget()
            self.lbl_excel.pack_forget()
            self.entry_para.pack(fill="x", padx=15, pady=(5, 15))
            self.caminho_excel = None
            self.lbl_excel.configure(text="Nenhum Excel selecionado")

    def selecionar_excel(self):
        ficheiro = filedialog.askopenfilename(title="Selecionar Lista Excel", filetypes=[("Documentos Excel", "*.xlsx *.xls")])
        if ficheiro:
            self.caminho_excel = ficheiro
            self.lbl_excel.configure(text=f"📂 {os.path.basename(ficheiro)}", text_color="#a78bfa")

    def selecionar_anexos(self):
        ficheiros = filedialog.askopenfilenames(title="Adicionar Ficheiros")
        if ficheiros:
            self.caminhos_anexos = list(ficheiros)
            nomes = [os.path.basename(f) for f in self.caminhos_anexos]
            self.lbl_anexos.configure(text=f"📎 {len(nomes)} ficheiro(s): " + ", ".join(nomes), text_color="#60a5fa")

    def substituir_variaveis(self, texto, linha_excel):
        for coluna, valor in linha_excel.items():
            # Evita erros de conversão se a célula estiver em branco (NaN) ou for um número
            valor_limpo = "" if pd.isna(valor) else str(valor)
            padrao = f"\\{{\\{{{coluna}\\}}\\}}"
            texto = re.sub(padrao, str(valor_limpo), texto)
        return texto

    # --- NOVO FLUXO ASSÍNCRONO DA GRAPH API ---
    async def executar_envio_graph(self, destinatarios_lista, assunto_base, corpo_html_base, dados_linhas=None):
        # Inicializa o motor de autenticação da aplicação do Azure
        credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret
        )
        graph_client = GraphServiceClient(credential)

        # Configuração do ficheiro de Log
        ficheiro_log = "envios_log.txt"

        # Processamento dos anexos em formato compatível com a API Graph
        lista_anexos_graph = []
        for caminho in self.caminhos_anexos:
            if os.path.exists(caminho):
                with open(caminho, "rb") as f:
                    conteudo_binario = f.read()
                
                # Converte o ficheiro em Base64
                #conteudo_base64 = base64.b64encode(conteudo_binario).decode('utf-8')
                tipo_mime, _ = mimetypes.guess_type(caminho)
                if not tipo_mime:
                    tipo_mime = "application/octet-stream"

                anexo = FileAttachment(
                    odata_type="#microsoft.graph.fileAttachment",
                    name=os.path.basename(caminho),
                    content_type=tipo_mime,
                    content_bytes=conteudo_binario
                )
                lista_anexos_graph.append(anexo)

        sucessos = 0
        falhas = 0
        TEMPO_PAUSA = 2  # Segundos de intervalo entre cada e-mail enviado

        # Envio iterativo para cada destinatário da lista
        for i, email_dest in enumerate(destinatarios_lista):

            # Temporizador: Se não for o primeiro e-mail, aguarda o intervalo definido
            if i > 0:
                self.lbl_status.configure(text=f"⏳ A aguardar {TEMPO_PAUSA}s (Evitar bloqueio)...", text_color="#f59e0b")
                self.update_idletasks()
                await asyncio.sleep(TEMPO_PAUSA)

            # Personalização dinâmica se vier do arquivo Excel
            assunto_final = assunto_base
            corpo_final = corpo_html_base
            if dados_linhas is not None:
                linha_atual = dados_linhas[i]
                assunto_final = self.substituir_variaveis(assunto_base, linha_atual)
                corpo_final = self.substituir_variaveis(corpo_html_base, linha_atual)

            # Construção do objeto de mensagem Microsoft Graph
            mensagem = Message(
                subject=assunto_final,
                body=ItemBody(
                    content_type=BodyType.Html,
                    content=corpo_final
                ),
                to_recipients=[
                    Recipient(email_address=EmailAddress(address=self.usuario))
                ],
                bcc_recipients=[
                    Recipient(email_address=EmailAddress(address=email_dest))
                ]
            )
            
            if lista_anexos_graph:
                mensagem.attachments = lista_anexos_graph

            corpo_pedido = SendMailPostRequestBody(
                message=mensagem,
                save_to_sent_items=True
            )

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            try:
                self.lbl_status.configure(text=f"🔄 A enviar em BCC para: {email_dest}...", text_color="#3b82f6")
                self.update_idletasks()
                
                # Pedido POST para a API de envio de email da Microsoft
                await graph_client.users.by_user_id(self.usuario).send_mail.post(body=corpo_pedido)
                sucessos += 1

                # Registo de Log em caso de Sucesso
                with open(ficheiro_log, "a", encoding="utf-8") as log:
                    log.write(f"[{timestamp}] SUCESSO | Destinatário BCC: {email_dest} | Assunto: {assunto_final}\n")

            except TypeError as erro_tipo:
                # Captura e ignora o erro de 'NoneType' gerado internamente pelo SDK após o envio
                if "NoneType" in str(erro_tipo) or "cannot be awaited" in str(erro_tipo).lower():
                    sucessos += 1

                    # Registo de Log em caso de Sucesso
                    with open(ficheiro_log, "a", encoding="utf-8") as log:
                        log.write(f"[{timestamp}] SUCESSO | Destinatário BCC: {email_dest} | Assunto: {assunto_final}\n")

                else:
                    print(f"Erro de tipo inesperado em {email_dest}: {erro_tipo}")
                    falhas += 1

                # Registo de Log em caso de Falha
                with open(ficheiro_log, "a", encoding="utf-8") as log:
                    log.write(f"[{timestamp}] ERRO    | Destinatário BCC: {email_dest} | Motivo: {e}\n")

            except Exception as e:
                print(f"Erro em {email_dest}: {e}")
                falhas += 1

                # Registo de Log em caso de Falha
                with open(ficheiro_log, "a", encoding="utf-8") as log:
                    log.write(f"[{timestamp}] ERRO    | Destinatário BCC: {email_dest} | Motivo: {e}\n")

        credential.close()
        return sucessos, falhas

    def processar_envio(self):
        if not self.env_valido:
            self.lbl_status.configure(text="Erro: Chaves da Graph API em falta no .env.", text_color="#ef4444")
            return

        assunto_base = self.entry_assunto.get().strip()
        corpo_html_base = self.txt_mensagem.get("0.0", tk.END)

        if not assunto_base:
            self.lbl_status.configure(text="Erro: O assunto está vazio.", text_color="#ef4444")
            return

        destinatarios = []
        dados_linhas = None

        # Validação da origem dos dados (Manual vs Excel)
        if self.switch_excel.get() == 1:
            if not self.caminho_excel:
                self.lbl_status.configure(text="Erro: Selecione um ficheiro Excel.", text_color="#ef4444")
                return
            try:
                df = pd.read_excel(self.caminho_excel)
                # Procura uma coluna chamada 'Email' ou 'email'
                coluna_email = next((col for col in df.columns if col.lower() == 'email'), None)
                if not coluna_email:
                    self.lbl_status.configure(text="Erro: Coluna 'Email' não encontrada no Excel.", text_color="#ef4444")
                    return
                
                destinatarios = df[coluna_email].dropna().astype(str).tolist()
                dados_linhas = df.to_dict(orient="records")
            except Exception as e:
                self.lbl_status.configure(text=f"Erro ao ler Excel: {e}", text_color="#ef4444")
                return
        else:
            email_unico = self.entry_para.get().strip()
            if not email_unico:
                self.lbl_status.configure(text="Erro: Insira o e-mail de destino.", text_color="#ef4444")
                return
            destinatarios = [email_unico]

        if not destinatarios:
            self.lbl_status.configure(text="Erro: Nenhum destinatário detetado.", text_color="#ef4444")
            return

        self.lbl_status.configure(text="🔄 A ligar à Graph API da Microsoft...", text_color="#3b82f6")
        self.update_idletasks()

        # Execução do loop assíncrono dentro da função síncrona do Tkinter
        try:
            sucessos, falhas = asyncio.run(self.executar_envio_graph(
                destinatarios, assunto_base, corpo_html_base, dados_linhas
            ))
            self.lbl_status.configure(
                text=f"✨ Concluído! Sucessos: {sucessos} | Falhas: {falhas}", 
                text_color="#22c55e" if falhas == 0 else "#f59e0b"
            )

        except Exception as erro_geral:

        # Importa a biblioteca de diagnóstico do Python
            import traceback
            
            # Captura a árvore completa do erro em formato texto
            detalhes_do_erro = traceback.format_exc()
            
            # Imprime o relatório detalhado no seu terminal/consola
            print("\n" + "="*50)
            print("🚨 RELATÓRIO DE DEBUG DA GRAPH API:")
            print("="*50)
            print(detalhes_do_erro)
            print("="*50 + "\n")

            self.lbl_status.configure(text=f"Erro na Graph API: {erro_geral}", text_color="#ef4444")

if __name__ == "__main__":
    app = InterfaceModernaApp()
    app.mainloop()