import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def selecionar_excel():
    caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls *.xlsx")])
    entrada_excel.delete(0, tk.END)
    entrada_excel.insert(0, caminho)

def selecionar_pasta():
    caminho = filedialog.askdirectory()
    entrada_pasta.delete(0, tk.END)
    entrada_pasta.insert(0, caminho)

def comparar():
    caminho_excel = entrada_excel.get()
    caminho_pasta = entrada_pasta.get()

    if not caminho_excel or not caminho_pasta:
        messagebox.showwarning("Atenção", "Selecione o arquivo Excel e a pasta com XMLs.")
        return

    try:
        df = pd.read_excel(caminho_excel)
        chaves_excel = df['Chave de Acesso'].astype(str).str.strip()

        arquivos_xml = os.listdir(caminho_pasta)
        chaves_xml = [arquivo.split('-')[0] for arquivo in arquivos_xml if arquivo.endswith('.xml')]

        faltantes = [chave for chave in chaves_excel if chave not in chaves_xml]

        resultado_texto.delete(1.0, tk.END)
        resultado_texto.insert(tk.END, f"Total de chaves no Excel: {len(chaves_excel)}\n")
        resultado_texto.insert(tk.END, f"Total de XMLs encontrados: {len(chaves_xml)}\n")
        resultado_texto.insert(tk.END, f"Chaves faltando ({len(faltantes)}):\n\n")
        for chave in faltantes:
            resultado_texto.insert(tk.END, f"{chave}\n")

        # Salvar em Excel
        pd.DataFrame(faltantes, columns=['Chaves Faltantes']).to_excel('faltando_xmls.xlsx', index=False)

        messagebox.showinfo("Concluído", "Comparação finalizada. Resultado salvo em 'faltando_xmls.xlsx'.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

# Interface gráfica
janela = tk.Tk()
janela.title("Comparador de XMLs - Edivim")
janela.geometry("600x500")

tk.Label(janela, text="Arquivo Excel (.xls/.xlsx):").pack()
entrada_excel = tk.Entry(janela, width=80)
entrada_excel.pack()
tk.Button(janela, text="Selecionar Excel", command=selecionar_excel).pack(pady=5)

tk.Label(janela, text="Pasta com arquivos XML:").pack()
entrada_pasta = tk.Entry(janela, width=80)
entrada_pasta.pack()
tk.Button(janela, text="Selecionar Pasta", command=selecionar_pasta).pack(pady=5)

tk.Button(janela, text="Comparar", command=comparar, bg="green", fg="white").pack(pady=10)

resultado_texto = scrolledtext.ScrolledText(janela, width=70, height=15)
resultado_texto.pack(pady=10)

janela.mainloop()