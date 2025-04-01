import os
import shutil
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox

# Caminhos das pastas
PASTA_INTEGRAR = r"D:\Programa xml\cte\integrar"
PASTA_PROCESSADOS = r"D:\Programa xml\cte\processados"


def modificar_xml(novo_cfop):
    if not novo_cfop.isdigit():
        messagebox.showerror("Erro", "O CFOP deve conter apenas números!")
        return

    # Verifica se a pasta existe
    if not os.path.exists(PASTA_INTEGRAR):
        messagebox.showerror("Erro", f"A pasta {PASTA_INTEGRAR} não foi encontrada!")
        return

    arquivos = [f for f in os.listdir(PASTA_INTEGRAR) if f.endswith('.xml')]

    if not arquivos:
        messagebox.showinfo("Info", "Nenhum arquivo XML encontrado para processar.")
        return

    for arquivo in arquivos:
        caminho_arquivo = os.path.join(PASTA_INTEGRAR, arquivo)

        try:
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()

            # Namespace do XML (extraído do próprio XML)
            ns = {'cte': 'http://www.portalfiscal.inf.br/cte'}

            # Encontrar e modificar a tag <CFOP>
            cfop_element = root.find('.//cte:CFOP', ns)
            if cfop_element is not None:
                cfop_element.text = novo_cfop
            else:
                messagebox.showwarning("Aviso", f"CFOP não encontrado em {arquivo}.")
                continue

            # Salvar o XML modificado na pasta processados
            novo_caminho = os.path.join(PASTA_PROCESSADOS, arquivo)
            tree.write(novo_caminho, encoding="utf-8", xml_declaration=True)

            # Mover o arquivo original para a pasta processados
            shutil.move(caminho_arquivo, os.path.join(PASTA_PROCESSADOS, f"backup_{arquivo}"))

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar {arquivo}: {str(e)}")

    messagebox.showinfo("Sucesso", f"Processamento concluído! {len(arquivos)} arquivos alterados.")


# Criando interface gráfica
def criar_interface():
    root = tk.Tk()
    root.title("Alteração de CFOP em XML")
    root.geometry("400x200")

    tk.Label(root, text="Novo CFOP:").pack(pady=10)
    entrada_cfop = tk.Entry(root)
    entrada_cfop.pack(pady=5)

    botao_executar = tk.Button(root, text="Alterar XML", command=lambda: modificar_xml(entrada_cfop.get()))
    botao_executar.pack(pady=10)

    root.mainloop()


# Executa o programa
criar_interface()