import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, scrolledtext
import pdfplumber
import pandas as pd
import os

class ModernPDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LBS - Extrator de Dados de PDF para Excel")
        self.root.geometry("900x600")
        self.root.configure(bg="#2e3440")

        style = ttk.Style()
        style.theme_use('clam')

        style.configure("TLabel", background="#2e3440", foreground="#d8dee9", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI Semibold", 10), padding=6)
        style.configure("TEntry", font=("Segoe UI", 11))
        style.configure("TFrame", background="#3b4252")
        style.map("TButton",
                  background=[('active', '#81a1c1')],
                  foreground=[('active', '#2e3440')])

        # Frame do logo
        self.frame_logo = ttk.Frame(root, width=120, height=120, style="TFrame", relief="ridge")
        self.frame_logo.grid(row=0, column=0, rowspan=7, padx=(20,10), pady=20, sticky="n")
        self.frame_logo.grid_propagate(False)
        logo_label = ttk.Label(self.frame_logo, text="LOGO\nAqui", anchor="center", font=("Segoe UI Black", 14), foreground="#88c0d0", background="#3b4252")
        logo_label.place(relx=0.5, rely=0.5, anchor="center")

        # Frame principal
        self.frame_main = ttk.Frame(root, style="TFrame")
        self.frame_main.grid(row=0, column=1, sticky="nsew", padx=10, pady=20)
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(6, weight=1)

        #variável que receberá o caminho do PDF
        self.pdf_path = ""
        
        self.resultados = []

        self.btn_carregar = ttk.Button(self.frame_main, text="Selecionar PDF", command=self.carregar_pdf)
        self.btn_carregar.grid(row=0, column=0, sticky="w")

        lbl_palavras = ttk.Label(self.frame_main, text="Palavras-chave (separadas por vírgula):")
        lbl_palavras.grid(row=1, column=0, sticky="w", pady=(15,5))

        self.entry_palavras = ttk.Entry(self.frame_main, width=60)
        self.entry_palavras.grid(row=2, column=0, sticky="we")

        self.btn_buscar = ttk.Button(self.frame_main, text="Buscar", command=self.buscar_no_pdf)
        self.btn_buscar.grid(row=3, column=0, sticky="w", pady=(15,5))

        self.btn_limpar = ttk.Button(self.frame_main, text="Limpar Pesquisa", command=self.limpar_pesquisa)
        self.btn_limpar.grid(row=3, column=0, sticky="e", pady=(15,5))

        self.btn_cor = ttk.Button(self.frame_main, text="Selecionar Cor de Fundo", command=self.selecionar_cor)
        self.btn_cor.grid(row=4, column=0, sticky="w", pady=(5,10))

        self.txt_resultado = scrolledtext.ScrolledText(self.frame_main, wrap=tk.WORD, height=15, font=("Segoe UI", 11), background="#434c5e", foreground="#d8dee9", insertbackground="#d8dee9")
        self.txt_resultado.grid(row=5, column=0, sticky="nsew", pady=(0,10))
        self.frame_main.grid_rowconfigure(5, weight=1)

        self.btn_exportar = ttk.Button(self.frame_main, text="Exportar para Excel", command=self.exportar_excel)
        self.btn_exportar.grid(row=6, column=0, sticky="e")

    #Funcao para carregar o PDF
    def carregar_pdf(self):
        
        #path recebe o valor da função askopenfilename, que é o arquivo selecionado na caixa de diálogo (filedialog)
        #filetypes especifica o filtro para listar apenas arquivos .PDF
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        
        #verifica se o valor de path não está nulo
        if path:
            
            #a variável pdf_path recebe o valor de path, que foi recebido da função askopenfilename
            self.pdf_path = path
            #
            messagebox.showinfo("PDF Selecionado", f"Arquivo carregado: {os.path.basename(self.pdf_path)}")

    def buscar_no_pdf(self):
        if not self.pdf_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo PDF primeiro.")
            return

        palavras = [p.strip().lower() for p in self.entry_palavras.get().split(',') if p.strip()]
        if not palavras:
            messagebox.showwarning("Aviso", "Digite pelo menos uma palavra-chave.")
            return

        self.resultados.clear()
        self.txt_resultado.delete(1.0, tk.END)

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if texto:
                        for linha in texto.split("\n"):
                            if any(p in linha.lower() for p in palavras):
                                self.resultados.append({"Linha": linha})
                                self.txt_resultado.insert(tk.END, linha + "\n")

            if not self.resultados:
                self.txt_resultado.insert(tk.END, "Nenhum resultado encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o PDF: {e}")

    def exportar_excel(self):
        if not self.resultados:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel Files", "*.xlsx")],
                                               title="Salvar como")
        if caminho:
            df = pd.DataFrame(self.resultados)
            try:
                df.to_excel(caminho, index=False)
                messagebox.showinfo("Sucesso", f"Dados exportados para: {caminho}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar: {e}")

    def limpar_pesquisa(self):
        self.entry_palavras.delete(0, tk.END)
        self.txt_resultado.delete(1.0, tk.END)
        self.resultados.clear()

    def selecionar_cor(self):
        cor = colorchooser.askcolor(title="Escolha a cor de fundo")
        if cor[1]:
            self.root.configure(bg=cor[1])
            self.frame_logo.configure(style=None)
            self.frame_logo.configure(background=cor[1])
            self.frame_main.configure(style=None)
            self.frame_main.configure(background=cor[1])
            self.txt_resultado.configure(background=cor[1])

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernPDFExtractorApp(root)
    root.mainloop()
