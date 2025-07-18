import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, scrolledtext
from PIL import Image, ImageTk
import pdfplumber
import pandas as pd
import os

class ModernPDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LBS - Extrator de Dados de PDF para Excel")
        self.root.geometry("900x600")
        self.root.configure(bg="#B21F16")

        style = ttk.Style()
        style.theme_use('clam')
        
        im=Image.open("favicon.png")
        ft=ImageTk.PhotoImage(im)
        
        #style.configure("TLabel", background="white", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI Semibold", 10), padding=6)
        style.configure("TEntry", font=("Segoe UI", 11))
        style.configure("TFrame", background="#B21F16")
        style.map("TButton",
                  background=[('active', '#81a1c1')],
                  foreground=[('active', '#2e3440')])

        # Frame do logo
        
        self.frame_logo = ttk.Frame(root, width=120, height=120, style="TFrame", relief="ridge")
        self.frame_logo.grid(row=0, column=0, rowspan=7, padx=(20,10), pady=20, sticky="n")
        self.frame_logo.grid_propagate(False)
        logo_label = ttk.Label(self.frame_logo, image=ft, anchor="center", font=("Segoe UI Black", 14))
        logo_label.image=ft
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

        self.btn_exportar = ttk.Button(self.frame_main, text="Exportar para CSV", command=self.exportar_excel)
        self.btn_exportar.grid(row=6, column=0, sticky="e")

    #funcao para carregar o PDF
    def carregar_pdf(self):
        
        #path recebe o valor da função askopenfilename, o valor é o caminho do arquivo selecionado na caixa de diálogo (filedialog)
        #filetypes especifica o filtro para listar apenas arquivos .PDF
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        
        #verifica se o valor de path não está nulo (se está nulo, nenhum arquivo foi selecionado)
        if path:
            
            #a variável pdf_path recebe o valor de path (que foi recebido da função askopenfilename)
            self.pdf_path = path
            
            #caixa de diálogo para printar o valor de pdf_path (caminho do arquivo PDF)
            messagebox.showinfo("PDF Selecionado", f"Arquivo carregado: {os.path.basename(self.pdf_path)}")

    def buscar_no_pdf(self):
        
        #verifica se o valor de pdf_path está nulo (caso não tenha entrado na condição if path, ou seja, nenhum arquivo selecionado)
        if not self.pdf_path:
            messagebox.showwarning("Aviso", "Selecione um arquivo PDF primeiro.")
            return
        
        #palavras é um array que recebe as strings correspondentes de acordo com o delimitador ','
        palavras = [p.strip().lower() for p in self.entry_palavras.get().split(',') if p.strip()]
        
        #se o array palavras estiver vazio, significa que o campo entry_palavras também está vazio
        if not palavras:
            messagebox.showwarning("Aviso", "Digite pelo menos uma palavra-chave.")
            return
            
        #limpa o campo de resultados
        self.resultados.clear()
        self.txt_resultado.delete(1.0, tk.END)

        try:
            #aponta o caminho do PDF (pdf_path), abre o arquivo e atribui a pdf
            with pdfplumber.open(self.pdf_path) as pdf:
                
                #laço de repeticao para percorrer todas as páginas do PDF, e cada uma corresponderá a variável pagina
                for pagina in pdf.pages:
                    
                    #extrair o texto de cada página por vez
                    texto = pagina.extract_text()
                    
                    #se o texto não for nulo
                    if texto:
                        #divide a variável texto por quebra de linha, e cada valor retornado (correspondente a uma linha) está atribuído a variável linha
                        for linha in texto.split("\n"):
                            
                            #a funcao any verifica se há correspondencias entre cada string do array palavras, com cada linha do texto
                            if any(p in linha.lower() for p in palavras):
                                #se houver alguma correspondencia, a linha do PDF será atribuída a resultados
                                self.resultados.append({"Linha": linha})
                                self.txt_resultado.insert(tk.END, linha + "\n")

            #verifica se resultados está nulo (ou seja, se não houve correspondencia na verificação any)
            if not self.resultados:
                self.txt_resultado.insert(tk.END, "Nenhum resultado encontrado.")
                
        #em caso de erro, printar uma mensagem
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o PDF: {e}")

    def exportar_excel(self):
        if not self.resultados:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".csv",
                                               filetypes=[("CSV Files", "*.csv")],
                                               title="Salvar como")
        if caminho:
            df = pd.DataFrame(self.resultados)
            try:
                df.to_csv(caminho, index=False)
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
