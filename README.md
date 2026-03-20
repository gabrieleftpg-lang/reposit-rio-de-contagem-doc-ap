import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def contar_condenacoes(root_dir):
    resultados = []  # lista de dicts: {'AP':..., 'CPF':..., 'QTD':...}
    aps_com_condenacao = set()

    # percorre diretório root_dir procurando por pastas "CONDENAÇÃO"
    for ap_name in os.listdir(root_dir):
        ap_path = os.path.join(root_dir, ap_name)
        if not os.path.isdir(ap_path):
            continue
        # procurar recursivamente por pastas chamadas exatamente "CONDENAÇÃO"
        for dirpath, dirnames, filenames in os.walk(ap_path):
            # normalizar nome da pasta atual
            base = os.path.basename(dirpath)
            if base.upper() == "CONDENAÇÃO" or base.upper() == "CONDENACAO":
                # em CONDENAÇÃO, listar subpastas (cada subpasta corresponde a um CPF)
                for cpf_name in sorted(os.listdir(dirpath)):
                    cpf_path = os.path.join(dirpath, cpf_name)
                    if os.path.isdir(cpf_path):
                        # contar todos os arquivos recursivamente dentro da pasta CPF
                        qtd = 0
                        for _, _, files in os.walk(cpf_path):
                            qtd += len(files)
                        resultados.append({'AP': ap_name, 'CPF': cpf_name, 'QTD': qtd})
                        aps_com_condenacao.add(ap_name)
                # não descer mais dentro dessa árvore (se desejar evitar dupla contagem)
                # removendo subdiretórios evita que o os.walk entre novamente, mas aqui já tratamos
                # continue para procurar outras ocorrências se existirem
    return resultados, sorted(aps_com_condenacao)

def salvar_excel(resultados, aps_com_condenacao, caminho_saida):
    df = pd.DataFrame(resultados, columns=['AP', 'CPF', 'QTD'])
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Contagem', index=False)
        df_aps = pd.DataFrame({'APs_com_CONDENAÇÃO': aps_com_condenacao})
        df_aps.to_excel(writer, sheet_name='APs_COM_CONDENAÇÃO', index=False)

def selecionar_e_processar():
    pasta = filedialog.askdirectory(title="Selecione a pasta MASTER")
    if not pasta:
        return
    resultado, aps = contar_condenacoes(pasta)
    if not resultado:
        messagebox.showinfo("Resultado", "Nenhuma pasta CONDENAÇÃO encontrada.")
        return
    # pergunta onde salvar o excel
    arquivo_saida = filedialog.asksaveasfilename(
        title="Salvar arquivo Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not arquivo_saida:
        return
    salvar_excel(resultado, aps, arquivo_saida)
    messagebox.showinfo("Concluído", f"Processamento finalizado.\nArquivo salvo em:\n{arquivo_saida}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Contador de ARQUIVOS em CONDENAÇÃO")
    root.geometry("400x120")
    lbl = tk.Label(root, text="Clique em Processar para selecionar a pasta MASTER e iniciar.")
    lbl.pack(pady=10)
    btn = tk.Button(root, text="Processar", width=20, command=selecionar_e_processar)
    btn.pack(pady=10)
    root.mainloop()
