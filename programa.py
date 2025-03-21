import os
import tkinter as tk
from tkinter import ttk, messagebox
import win32print
import win32api

# Configurações: defina a pasta onde estão os arquivos
os.chdir(r"S:\PROJETO\PUBLICIDADE\Cabide de Kits\CabidesPDF")
PASTA_ARQUIVOS = os.getcwd()

def imprimir_arquivo(caminho_arquivo, copias=1, printer_name=None, paper_size="Carta"):
    """Imprime um arquivo no Windows."""
    if printer_name is None:
        printer_name = win32print.GetDefaultPrinter()
    # Aqui você pode implementar lógica para paper_size, se necessário.
    for _ in range(copias):
        win32api.ShellExecute(
            0,
            "print",
            caminho_arquivo,
            '/d:"%s"' % printer_name,
            ".",
            0
        )

class ArquivoRow:
    """Linha de widgets para cada arquivo."""
    def __init__(self, master, nome_arquivo, spinbox_style):
        self.nome = nome_arquivo
        self.var_selecionado = tk.BooleanVar(value=False)

        # Widgets com larguras definidas para alinhamento
        self.check = ttk.Checkbutton(master, variable=self.var_selecionado, width=2)
        self.label = ttk.Label(master, text=nome_arquivo, anchor="w", width=50)
        self.spin = ttk.Spinbox(master, from_=1, to=100, width=4, justify="center", style=spinbox_style)
        # Garante que comece com o valor "1"
        self.spin.delete(0, 'end')
        self.spin.insert(0, '1')

    def grid(self, row):
        """Posiciona os widgets na grid."""
        self.check.grid(row=row, column=0, padx=(5,2), pady=2, sticky="w")
        self.label.grid(row=row, column=1, padx=2, pady=2, sticky="w")
        self.spin.grid(row=row, column=2, padx=5, pady=2, sticky="w")

    def esta_selecionado(self):
        return self.var_selecionado.get()

    def get_copias(self):
        try:
            return int(self.spin.get())
        except ValueError:
            return 1

class PrintAutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação de Impressão")
        self.geometry("650x500")
        self.configure(bg="#2e2e2e")

        # Armazenamento dos arquivos e estado (para manter seleção e cópias)
        self.arquivos_todos = []
        self.arquivos = []
        self.arquivos_widgets = []
        # Dicionário para armazenar o estado de cada arquivo
        # Exemplo: {"nome_do_arquivo": {"selected": True, "copies": 2}, ...}
        self.estado_arquivos = {}

        # Estilos
        self.style = ttk.Style(self)
        self._configurar_estilos()

        # Criação dos elementos
        self.criar_widgets()
        self.carregar_arquivos()

    def _configurar_estilos(self):
        """Configura o tema escuro e cria um estilo para o Spinbox com fonte preta."""
        self.style.theme_use("clam")
        fonte_padrao = ("Segoe UI", 11)
        
        self.style.configure("TFrame", background="#2e2e2e")
        self.style.configure("TLabel", background="#2e2e2e", foreground="#ffffff", font=fonte_padrao)
        self.style.configure("TButton", background="#454545", foreground="#ffffff", font=fonte_padrao)
        self.style.configure("TEntry", fieldbackground="#454545", foreground="#ffffff", font=fonte_padrao)
        self.style.configure("TCheckbutton", background="#2e2e2e", foreground="#ffffff", font=fonte_padrao)
        self.style.configure("TRadiobutton", background="#2e2e2e", foreground="#ffffff", font=fonte_padrao)
        
        # Estilo especial para Spinbox com fonte preta
        self.style.configure(
            "SpinboxBlack.TSpinbox",
            foreground="black",
            fieldbackground="#ffffff",
            font=fonte_padrao
        )

    def criar_widgets(self):
        """Monta toda a interface."""
        # Frame de configurações (impressora e tamanho do papel)
        config_frame = ttk.Frame(self)
        config_frame.pack(fill="x", padx=10, pady=(10,5))
        ttk.Label(config_frame, text="Impressora:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entrada_impressora = ttk.Entry(config_frame, width=20)
        self.entrada_impressora.insert(0, "Samsung M4020")
        self.entrada_impressora.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(config_frame, text="Tamanho do Papel:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.tamanho_papel = tk.StringVar(value="Carta")
        rb_carta = ttk.Radiobutton(config_frame, text="Carta", variable=self.tamanho_papel, value="Carta")
        rb_a4 = ttk.Radiobutton(config_frame, text="A4", variable=self.tamanho_papel, value="A4")
        rb_carta.grid(row=0, column=3, padx=(5,0), pady=5, sticky="w")
        rb_a4.grid(row=0, column=4, padx=(0,5), pady=5, sticky="w")

        sep = ttk.Separator(self, orient="horizontal")
        sep.pack(fill="x", padx=10, pady=(0,5))

        # Frame para busca
        filtro_frame = ttk.Frame(self)
        filtro_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(filtro_frame, text="Buscar:").pack(side="left", padx=5)
        self.entrada_filtro = ttk.Entry(filtro_frame, width=25)
        self.entrada_filtro.pack(side="left", padx=5)
        self.entrada_filtro.bind("<KeyRelease>", lambda event: self.filtrar_lista())

        # Cabeçalho da lista
        cab_frame = ttk.Frame(self)
        cab_frame.pack(fill="x", padx=10)
        ttk.Label(cab_frame, text="Sel", width=5, anchor="w").grid(row=0, column=0, padx=(5,2), pady=2, sticky="w")
        ttk.Label(cab_frame, text="Nome do Arquivo", width=50, anchor="w").grid(row=0, column=1, padx=2, pady=2, sticky="w")
        ttk.Label(cab_frame, text="Cópias", width=6, anchor="w").grid(row=0, column=2, padx=5, pady=2, sticky="w")

        # Área de rolagem para a lista
        self.canvas = tk.Canvas(self, bg="#2e2e2e", highlightthickness=0)
        self.frame_lista = ttk.Frame(self.canvas)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        self.canvas.create_window((0, 0), window=self.frame_lista, anchor="nw")
        self.frame_lista.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Botão de imprimir
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text="Imprimir", command=self.processar_impressao).pack()

    def carregar_arquivos(self):
        """Carrega os arquivos da pasta."""
        lista = os.listdir(PASTA_ARQUIVOS)
        lista.sort()
        self.arquivos_todos = [arq for arq in lista if os.path.isfile(os.path.join(PASTA_ARQUIVOS, arq))]
        self.arquivos = self.arquivos_todos.copy()
        self.montar_lista()

    def update_estado(self):
        """Atualiza o dicionário com o estado atual (seleção e cópias) dos widgets visíveis."""
        for widget in self.arquivos_widgets:
            self.estado_arquivos[widget.nome] = {
                "selected": widget.esta_selecionado(),
                "copies": widget.get_copias()
            }

    def montar_lista(self):
        """Cria as linhas de arquivo no frame_lista, restaurando estado se existir."""
        for widget in self.frame_lista.winfo_children():
            widget.destroy()
        self.arquivos_widgets = []

        for i, nome in enumerate(self.arquivos):
            arquivo_row = ArquivoRow(self.frame_lista, nome, spinbox_style="SpinboxBlack.TSpinbox")
            # Se houver estado salvo, restaura
            if nome in self.estado_arquivos:
                estado = self.estado_arquivos[nome]
                arquivo_row.var_selecionado.set(estado.get("selected", False))
                arquivo_row.spin.delete(0, 'end')
                arquivo_row.spin.insert(0, str(estado.get("copies", 1)))
            self.arquivos_widgets.append(arquivo_row)
            arquivo_row.grid(row=i)

    def filtrar_lista(self):
        """Atualiza o estado, filtra a lista e reconstrói os widgets."""
        self.update_estado()
        termo = self.entrada_filtro.get().lower()
        self.arquivos = [nome for nome in self.arquivos_todos if termo in nome.lower()]
        self.montar_lista()

    def processar_impressao(self):
        """Envia para impressão os arquivos selecionados."""
        selecionados = [row for row in self.arquivos_widgets if row.esta_selecionado()]
        if not selecionados:
            messagebox.showinfo("Atenção", "Nenhum arquivo selecionado!")
            return

        resposta = messagebox.askyesno("Confirmação", f"Confirma imprimir {len(selecionados)} arquivos?")
        if not resposta:
            return

        printer_nome = self.entrada_impressora.get().strip() or win32print.GetDefaultPrinter()
        papel = self.tamanho_papel.get()  # "Carta" ou "A4"

        for row in selecionados:
            caminho = os.path.join(PASTA_ARQUIVOS, row.nome)
            copias = row.get_copias()
            try:
                imprimir_arquivo(caminho, copias, printer_nome, papel)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao imprimir {row.nome}: {e}")
        messagebox.showinfo("Concluído", "Impressão enviada para a impressora.")

if __name__ == "__main__":
    app = PrintAutomationApp()
    app.mainloop()
