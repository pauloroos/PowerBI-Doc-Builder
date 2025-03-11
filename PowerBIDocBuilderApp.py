import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import webbrowser
import os
import requests
import ttkbootstrap as ttk  # Suporte a temas avançados

from core.helpers import *
from core.ai_description import *
from core.diagram_generator import *
from core.pbi_extractor import *

# Janela de configuração
class ConfigWindow(tk.Toplevel):
    """Janela de configuração para definir caminhos."""
    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("Configuração de Caminhos")
        self.geometry("750x360")  # Aumentei a altura para melhor exibição
        # centralizar_janela(self, 750, 300)
        self.config(bg="#1E1B3A")
        self.resizable(False, False)  # Impede redimensionamento
        
        self.style = ttk.Style()
        self.style.theme_use("superhero")

        self.config_data = config

        # Criar um frame para melhor organização
        self.frame = ttk.Frame(self)
        self.frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # 1️⃣ Label da API Gemini
        ttk.Label(self.frame, text="API Gemini", foreground="white", background="#1E1B3A",
                anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(2, 2))

        # 2️⃣ Frame horizontal para campo + botão
        api_frame = ttk.Frame(self.frame)
        api_frame.grid(row=1, column=0, columnspan=2, padx=10, sticky="w")

        # 3️⃣ Campo de entrada da chave
        self.api_entry = ttk.Entry(api_frame, width=70)
        self.api_entry.insert(0, self.config_data.get("api_key", ""))
        self.api_entry.pack(side=tk.LEFT)

        # 4️⃣ Botão de acesso à chave (mais largo)
        ttk.Button(api_frame, text="🔗 Acessar para Gerar Token", width=30, bootstyle="info",
                command=lambda: webbrowser.open("https://aistudio.google.com/app/apikey")).pack(side=tk.LEFT, padx=(10, 0))
        
        # Criando os campos organizados (Títulos acima das caixas)
        self.criar_campo("DAX Studio CLI", "cmd", 2)
        self.criar_campo("Analysis Services DLL", "ssas_dll", 3)
        self.criar_campo("Power BI Desktop", "pbi_desktop", 4)
        
        # Botão de salvar
        ttk.Button(self.frame, text="Salvar Configuração", bootstyle="success", command=self.salvar).grid(row=10, column=0, columnspan=2, pady=15)

    def criar_campo(self, label_text, key, row):
        """Cria um campo com título acima, entrada de texto e botão de seleção."""
        ttk.Label(self.frame, text=label_text, foreground="white", background="#1E1B3A", anchor="w", font=("Arial", 10, "bold")).grid(row=row*2, column=0, columnspan=2, sticky="w", padx=10, pady=2)

        entry = ttk.Entry(self.frame, width=105)
        entry.insert(0, self.config_data.get(key, ""))
        entry.grid(row=row*2+1, column=0, padx=10, pady=5, sticky="w")

        button = ttk.Button(self.frame, text="📂", width=3, command=lambda: self.selecionar_arquivo(entry))
        button.grid(row=row*2+1, column=1, padx=5, pady=5, sticky="w")

        setattr(self, f"{key}_entry", entry)  # Armazena a entrada de texto dinamicamente

    def selecionar_arquivo(self, entry):
        """Abre o seletor de arquivos e insere o caminho na caixa de texto."""
        caminho = filedialog.askopenfilename()
        if caminho:
            entry.delete(0, tk.END)
            entry.insert(0, caminho)

    def salvar(self):
        """Salva os caminhos no arquivo de configuração."""
        self.config_data["cmd"] = self.cmd_entry.get()
        self.config_data["ssas_dll"] = self.ssas_dll_entry.get()
        self.config_data["pbi_desktop"] = self.pbi_desktop_entry.get()
        self.config_data["api_key"] = self.api_entry.get()
        salvar_config(self.config_data)
        messagebox.showinfo("Configuração", "Caminhos salvos com sucesso!")
        self.destroy()

# Aplicativo principal
class PowerBIDocApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerBI Doc Builder")
        self.root.geometry("1080x675")  # 1440x900 escalado para 75%
        # centralizar_janela(self.root, 1080, 675)  # Centraliza
        self.root.configure(bg="#0D0B1A")
        self.root.resizable(False, False)  # Impede redimensionamento

        # Aplicar tema inicial
        self.style = ttk.Style()
        self.style.theme_use("superhero")

        # Verificar e carregar configuração
        self.config_data = carregar_config()
        if not self.config_data:
            self.abrir_config()
        
        # Menu para abrir configurações
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)
        config_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Configurações", menu=config_menu)
        config_menu.add_command(label="Definir Caminhos", command=self.abrir_config)
                
        # Caminho da imagem baseado na pasta do script
        image_path = os.path.join(os.path.dirname(__file__), "assets/image.png")
        self.bg_image = Image.open(image_path)
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)

        self.bg_label = tk.Label(self.root, borderwidth=0)
        self.bg_label.place(x=0, y=0, width=400, height=675)

        # Atualizar a imagem para preencher corretamente
        self.update_background()

        # Container para os elementos
        self.container = tk.Frame(self.root, bg="#1E1B3A", bd=5, relief="flat")
        self.container.place(x=420, y=80, width=600, height=580)

        # Ícone (imagem) para o título (50% maior)
        icon_path = os.path.join(os.path.dirname(__file__), "assets/icon.png")
        if os.path.exists(icon_path):
            self.icon_image = Image.open(icon_path)
            self.icon_image = self.icon_image.resize((75, 75), Image.LANCZOS)  # Aumento de 50%
            self.icon_photo = ImageTk.PhotoImage(self.icon_image)
            self.icon_label = tk.Label(self.container, image=self.icon_photo, bg="#1E1B3A")
            self.icon_label.pack(pady=10)

        # Título
        self.title_label = tk.Label(
            self.container, text="PowerBI Doc Builder", fg="white", bg="#1E1B3A",
            font=("Arial", 22, "bold")
        )
        self.title_label.pack()

        # Subtítulo
        self.subtitle_label = tk.Label(
            self.container, text="by Paulo Roos", fg="gray", bg="#1E1B3A",
            font=("Arial", 12)
        )
        self.subtitle_label.pack()

        # Texto de instrução atualizado e alinhado à esquerda
        self.instruction_label = tk.Label(
            self.container,
            text="Este aplicativo gera uma documentação detalhada para modelos PBIX do Power BI, incluindo informações sobre tabelas, colunas, medidas, "
                 "relacionamentos, partições e grupos de cálculo.",
            fg="white", bg="#1E1B3A", font=("Arial", 12), wraplength=500, justify="left", anchor="w"
        )
        self.instruction_label.pack(pady=10, fill="x", padx=20)

        # Descrição do aplicativo ajustada
        self.description_label = tk.Label(
            self.container, text="Selecione a pasta onde estão o(s) arquivo(s) PBIX:",
            fg="white", bg="#1E1B3A", font=("Arial", 10), anchor="w"
        )
        self.description_label.pack(pady=5, fill="x", padx=20)

        # Frame de entrada
        self.entry_frame = tk.Frame(self.container, bg="#1E1B3A")
        self.entry_frame.pack(padx=20, pady=5, anchor="w")

        # Variável para armazenar o caminho da pasta
        self.caminho_var = tk.StringVar()

        # Entrada de texto estilizada (SOMENTE LEITURA)
        self.path_entry = ttk.Entry(
            self.entry_frame, textvariable=self.caminho_var, width=50, font=("Arial", 12), state="readonly"
        )
        self.path_entry.pack(side=tk.LEFT, ipady=5, padx=(0, 10))

        # Botão de seleção de pasta (FONTE MAIOR)
        self.browse_button = ttk.Button(
            self.entry_frame, text="📂", command=self.selecionar_pasta, width=4, style="primary"
        )
        self.browse_button.pack(side=tk.RIGHT)
        self.browse_button.configure(style="Big.TButton")

        # Criando um novo estilo com fonte maior
        self.style.configure("Big.TButton", font=("Arial", 18, "bold"))

        # **Botão para gerar o relatório (50% maior)**
        self.generate_button = ttk.Button(
            self.container, text="Gerar Relatório", command=self.gerar_documentacao,
            width=25, bootstyle="success"
        )
        self.generate_button.pack(pady=30)
        self.generate_button.configure(style="success.TButton")
        self.style.configure("success.TButton", font=("Arial", 16, "bold"))

        # **Botões de redes sociais**
        self.social_frame = tk.Frame(self.container, bg="#1E1B3A")
        self.social_frame.pack(pady=10)

        self.add_social_button("LinkedIn", LINKEDIN_URL, ICON_URLS["LinkedIn"])
        self.add_social_button("GitHub", GITHUB_URL, ICON_URLS["GitHub"])
        self.add_social_button("E-mail", EMAIL_URL, ICON_URLS["E-mail"])
        self.add_social_button("Site", SITE_URL, ICON_URLS["Site"])

    def add_social_button(self, name, url, icon_url):
        """Baixa um ícone e cria um botão para rede social"""
        try:
            response = requests.get(icon_url, stream=True)
            response.raise_for_status()
            icon = Image.open(response.raw).resize((35, 35), Image.LANCZOS)  # **Tamanho igual ao botão da pasta**
            icon_photo = ImageTk.PhotoImage(icon)

            # button = ttk.Button(
            #     self.social_frame, image=icon_photo, command=lambda: webbrowser.open(url),
            #     bootstyle="secondary", width=3
            # )
            # button.image = icon_photo  # Evita garbage collection
            # button.pack(side=tk.LEFT, padx=10)

            label = tk.Label(self.social_frame, image=icon_photo, bg="#1E1B3A", cursor="hand2")
            label.image = icon_photo  # Evita garbage collection
            label.pack(side=tk.LEFT, padx=10)
            label.bind("<Button-1>", lambda event: webbrowser.open(url))
            
        except requests.RequestException:
            print(f"Erro ao carregar ícone de {name}")

    def update_background(self):
        """ Atualiza o fundo para preencher corretamente sem distorções."""
        bg_resized = self.bg_image.resize((400, 675), Image.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(bg_resized)
        self.bg_label.config(image=self.bg_photo)

    def selecionar_pasta(self):
        """Abre um seletor de diretórios e atualiza a variável de caminho."""
        pasta = filedialog.askdirectory()
        if pasta:
            self.caminho_var.set(pasta)

    def gerar_documentacao(self):
        """Executa o processamento da documentação."""
        pasta = self.caminho_var.get()
        if not pasta:
            messagebox.showerror("Erro", "Selecione uma pasta primeiro!")
            return

        # Procurar arquivos .pbix na pasta
        arquivos_pbix = [f for f in os.listdir(pasta) if f.lower().endswith(".pbix")]
        
        if not arquivos_pbix:
            messagebox.showwarning("Aviso", "Nenhum arquivo .PBIX encontrado na pasta selecionada.")
            return

        # Listar arquivos encontrados
        lista_pbix = "\n".join(arquivos_pbix)
        confirmar = messagebox.askokcancel(
            "Confirmar geração",
            f"Os seguintes arquivos PBIX serão processados:\n\n{lista_pbix}\n\nDeseja continuar?"
        )

        if not confirmar:
            return  # Usuário cancelou

        try:
            # ⚠️ Aqui você deve chamar a função real de geração da documentação ⚠️:
            processar_pbix(pasta)

            messagebox.showinfo(
                "Sucesso",
                f"Documentação gerada com sucesso!\n\nArquivos processados:\n{lista_pbix}\n\n"
                f"Pasta de saída:\n{os.path.join(pasta, 'Resultado')}"
            )

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar a documentação:\n{str(e)}")


    def abrir_config(self):
        """Abre a janela de configuração"""
        config_window = ConfigWindow(self.root, self.config_data)
        self.root.wait_window(config_window)

        # Atualiza os rótulos após salvar os caminhos
        self.config_data = carregar_config()

# Executar o aplicativo        
if __name__ == "__main__":
    root = tk.Tk()
    app = PowerBIDocApp(root)
    root.mainloop()
