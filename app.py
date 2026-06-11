import os
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import re
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

# Configuração estética do sistema
ctk.set_appearance_mode("Dark")  # Força o modo escuro para um aspeto mais moderno
ctk.set_default_color_theme("blue")

class InterfaceModernaApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Microsoft SMTP Mailer Pro")
        self.geometry("900x650")
        self.resizable(False, False)
        
        self.caminhos_anexos = []
        self.caminho_excel = None
        self.usuario = os.getenv("MS_USER")
        self.senha = os.getenv("MS_PASSWORD")

        # --- ESTRUTURA DE LAYOUT (Grelha Principal) ---
        self.grid_columnconfigure(0, weight=1, minsize=380) # Painel de Controlo (Esquerda)
        self.grid_columnconfigure(1, weight=2, minsize=520) # Painel de Conteúdo (Direita)
        self.grid_rowconfigure(0, weight=1)

        # ==========================================
        # PAINEL ESQUERDO: Configurações e Destino
        # ==========================================
        self.frame_esquerda = ctk.CTkFrame(self, corner_radius=15, fg_color="#1e1e24")
        self.frame_esquerda.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        
        # Cabeçalho da Aplicação
        self.lbl_logo = ctk.CTkLabel(self.frame_esquerda, text="🚀 SMTP Mailer", font=ctk.CTkFont(size=24, weight="bold"))
        self.lbl_logo.pack(pady=(25, 5), padx=20, anchor="w")
        
        # Indicador de Conta Autenticada
        status_conta = f"🔒 Ativo: {self.usuario}" if self.usuario and self.senha else "❌ Falha no ficheiro .env"
        cor_conta = "#2ecc71" if self.usuario and self.senha else "#e74c3c"
        self.lbl_auth = ctk.CTkLabel(self.frame_esquerda, text=status_conta, text_color=cor_conta, font=ctk.CTkFont(size=12))
        self.lbl_auth.pack(pady=(0, 20), padx=20, anchor="w")

        # Cartão de Destinatário (Modo Switch)
        self.card_destino = ctk.CTkFrame(self.frame_esquerda, fg_color="#2a2a35", corner_radius=10)
        self.card_destino.pack(fill="x", padx=20, pady=10)
        
        self.lbl_card_tit = ctk.CTkLabel(self.card_destino, text="Destinatários", font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_card_tit.pack(pady=(10, 5), padx=15, anchor="w")

        self.switch_excel = ctk.CTkSwitch(self.card_destino, text="Enviar para Lista Excel", command=self.alternar_modo, progress_color="#3b82f6")
        self.switch_excel.pack(pady=10, padx=15, anchor="w")

        self.entry_para = ctk.CTkEntry(self.card_destino, placeholder_text="Destinatário (email individual)", height=40, border_color="#47475a")
        self.entry_para.pack(fill="x", padx=15, pady=(5, 15))

        # Elementos dinâmicos do Excel (escondidos inicialmente)
        self.btn_excel = ctk.CTkButton(self.card_destino, text="📊 Importar Ficheiro Excel", command=self.selecionar_excel, fg_color="#6366f1", hover_color="#4f46e5", height=35)
        self.lbl_excel = ctk.CTkLabel(self.card_destino, text="Nenhum Excel selecionado", text_color="#a1a1aa", font=ctk.CTkFont(size=11))

        # Cartão de Anexos
        self.card_anexos = ctk.CTkFrame(self.frame_esquerda, fg_color="#2a2a35", corner_radius=10)
        self.card_anexos.pack(fill="x", padx=20, pady=10)
        
        self.btn_anexar = ctk.CTkButton(self.card_anexos, text="📎 Adicionar Anexos", command=self.selecionar_anexos, fg_color="#4b5563", hover_color="#374151")
        self.btn_anexar.pack(fill="x", padx=15, pady=(15, 5))
        
        self.lbl_anexos = ctk.CTkLabel(self.card_anexos, text="Nenhum ficheiro anexado", text_color="#a1a1aa", font=ctk.CTkFont(size=11), wraplength=300)
        self.lbl_anexos.pack(pady=(0, 15), padx=15)

        # Zona de Ação Inferior (Botão de Envio Principal)
        self.btn_enviar = ctk.CTkButton(self.frame_esquerda, text="Iniciar Disparo", command=self.processar_envio, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#22c55e", hover_color="#16a34a", height=50, corner_radius=8)
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

        # Caixa de texto profissional com fundo contrastante
        self.txt_mensagem = ctk.CTkTextbox(self.frame_direita, border_color="#27272a", fg_color="#09090b", font=ctk.CTkFont(family="Consolas", size=12))
        self.txt_mensagem.pack(fill="both", expand=True, padx=25, pady=(0, 25))
        self.txt_mensagem.insert("0.0", "<!-- Código HTML Fixo ou Dinâmico -->\n<h1>Olá {{Nome}}</h1>\n<p>O seu email de teste formatado.</p>")

    def alternar_modo(self):
        if self.switch_excel.get() == 1:
            # Oculta campo manual com animação visual limpa
            self.entry_para.pack_forget()
            # Mostra utilitários do Excel
            self.btn_excel.pack(fill="x", padx=15, pady=(5, 2))
            self.lbl_excel.pack(pady=(0, 15), padx=15)
        else:
            # Reverte para o estado inicial
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
            padrao = f"\\{{\\{{{coluna}\\}}\\}}"
            texto = re.sub(padrao, str(valor), texto)
        return texto

    def processar_envio(self):
        if not self.usuario or not self.senha:
            self.lbl_status.configure(text="Erro: Sem .env configurado.", text_color="#ef4444")
            return

        assunto_base = self.entry_assunto.get()
        corpo_html_base = self.txt_mensagem.get("0.0", tk.END)

        self.lbl_status.configure(text="🔄 A ligar ao servidor Microsoft...", text_color="#3b82f6")
        self.update_idletasks()

        try:
            server = smtplib.SMTP("office365.com", 587)
            server.starttls()
            server.login(self.usuario, self.senha)
        except Exception as e:
            self.lbl_status.configure(text=f"❌ Erro de Login: {str(e)[:30]}...", text_color="#ef4444")
            return

        # Execução com base no modo selecionado
        if self.switch_excel.get() == 1:
            if not self.caminho_excel:
                self.lbl_status.configure(text="⚠️ Selecione o ficheiro Excel primeiro.", text_color="#f59e0b")
                server.quit()
                return

            try:
                df = pd.read_excel(self.caminho_excel)
                if 'Email' not in df.columns:
                    self.lbl_status.configure(text="⚠️ Falta a coluna 'Email' no Excel.", text_color="#ef4444")
                    server.quit()
                    return
            except Exception as e:
                self.lbl_status.configure(text="❌ Erro ao processar o Excel.", text_color="#ef4444")
                server.quit()
                return

            resultados = []
            sucessos = 0
            possui_variaveis = len(df.columns) > 1

            for index, linha in df.iterrows():
                destinatario = str(linha['Email']).strip()
                agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                if not destinatario or "@" not in destinatario or destinatario.lower() == "nan":
                    resultados.append(f"Inválido ({agora})")
                    continue

                assunto = self.substituir_variaveis(assunto_base, linha) if possui_variaveis else assunto_base
                corpo = self.substituir_variaveis(corpo_html_base, linha) if possui_variaveis else corpo_html_base

                msg = self.criar_mensagem(self.usuario, destinatario, assunto, corpo)
                
                try:
                    server.sendmail(self.usuario, destinatario, msg.as_string())
                    sucessos += 1
                    resultados.append(f"Sucesso ({agora})")
                except Exception as e:
                    resultados.append(f"Falha: {str(e)[:20]} ({agora})")

                self.lbl_status.configure(text=f"Progressão: {index + 1}/{len(df)} emails", text_color="#f59e0b")
                self.update_idletasks()

            df['Estado do Envio'] = resultados
            nome_relatorio = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            caminho_relatorio = os.path.join(os.path.dirname(self.caminho_excel), nome_relatorio)
            df.to_excel(caminho_relatorio, index=False)
            
            self.lbl_status.configure(text=f"✅ Concluído! Relatório: {nome_relatorio}", text_color="#10b981")
        
        else:
            destinatario = self.entry_para.get()
            if not destinatario or "@" not in destinatario:
                self.lbl_status.configure(text="⚠️ Indique um email de destino válido.", text_color="#f59e0b")
                server.quit()
                return

            msg = self.criar_mensagem(self.usuario, destinatario, assunto_base, corpo_html_base)
            try:
                server.sendmail(self.usuario, destinatario, msg.as_string())
                self.lbl_status.configure(text="✅ Email enviado com sucesso!", text_color="#10b981")
            except Exception as e:
                self.lbl_status.configure(text="❌ Erro no envio individual.", text_color="#ef4444")

        server.quit()

    def criar_mensagem(self, de, para, assunto, html):
        msg = MIMEMultipart()
        msg['From'] = de
        msg['To'] = para
        msg['Subject'] = assunto
        msg.attach(MIMEText(html, 'html', 'utf-8'))

        for caminho in self.caminhos_anexos:
            if os.path.exists(caminho):
                with open(caminho, "rb") as anexo:
                    parte = MIMEBase("application", "octet-stream")
                    parte.set_payload(anexo.read())
                encoders.encode_base64(parte)
                parte.add_header("Content-Disposition", f"attachment; filename={os.path.basename(caminho)}")
                msg.attach(parte)
        return msg

if __name__ == "__main__":
    app = InterfaceModernaApp()
    app.mainloop()
