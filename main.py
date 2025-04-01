import os
import shutil
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox

# Caminhos das pastas
PASTA_INTEGRAR = r"D:\Programa xml\cte\integrar"
PASTA_PROCESSADOS = r"D:\Programa xml\cte\processados"
PASTA_BACKUP = r"D:\Programa xml\cte\backup"

# Variável de controle para cancelamento
cancelar_processo = False


def modificar_xml():
    global cancelar_processo
    cancelar_processo = False

    arquivos = [f for f in os.listdir(PASTA_INTEGRAR) if f.endswith('.xml')]

    if not arquivos:
        messagebox.showinfo("Info", "Nenhum arquivo XML encontrado para processar.")
        return

    total_arquivos = len(arquivos)
    progresso["maximum"] = total_arquivos

    for index, arquivo in enumerate(arquivos, start=1):
        if cancelar_processo:
            label_progresso["text"] = "Processo cancelado!"
            progresso["value"] = 0
            return

        caminho_arquivo = os.path.join(PASTA_INTEGRAR, arquivo)

        try:
            tree = ET.parse(caminho_arquivo)
            xml_root = tree.getroot()  # Usando xml_root para evitar conflitos
            ns = {'cte': 'http://www.portalfiscal.inf.br/cte'}

            alterado = False  # Flag para verificar se o XML foi modificado

            # Modificando CFOP
            cfop_element = xml_root.find('.//cte:CFOP', ns)
            if cfop_element is not None and entrada_cfop.get():
                cfop_element.text = entrada_cfop.get()
                alterado = True

            # Modificando Transportadora
            transportadora = xml_root.find('.//cte:transporta', ns)
            if transportadora is not None:
                for tag, entry in entradas.items():
                    element = transportadora.find(f'cte:{tag}', ns)
                    if element is not None and entry.get():
                        element.text = entry.get()
                        alterado = True

            # Alterando Local de Entrega
            local_entrega_element = xml_root.find(".//cte:ObsCont[@xCampo='LocalDeEntrega']/cte:xTexto", ns)
            if local_entrega_element is not None and entrada_local_entrega.get():
                local_entrega_element.text = entrada_local_entrega.get()
                alterado = True

            # Criando diretórios caso não existam
            os.makedirs(PASTA_BACKUP, exist_ok=True)
            os.makedirs(PASTA_PROCESSADOS, exist_ok=True)

            # Mover para backup
            shutil.move(caminho_arquivo, os.path.join(PASTA_BACKUP, f"backup_{arquivo}"))

            if alterado:
                # Salvar arquivo modificado na pasta processados
                novo_caminho = os.path.join(PASTA_PROCESSADOS, arquivo)
                tree.write(novo_caminho, encoding="utf-8", xml_declaration=True)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar {arquivo}: {str(e)}")
            continue

        # Atualizando barra de progresso e contador
        progresso["value"] = index
        label_progresso["text"] = f"Processando {index}/{total_arquivos} arquivos..."
        janela.update_idletasks()  # Correção aqui!

    messagebox.showinfo("Sucesso", f"Processamento concluído! {total_arquivos} arquivos alterados.")


def cancelar_processo_xml():
    global cancelar_processo
    cancelar_processo = True
    label_progresso["text"] = "Cancelando processo..."
    janela.update_idletasks()


# Criando interface gráfica
janela = tk.Tk()  # Alterado de root para janela
janela.title("Alteração de XML de Notas Fiscais")
janela.geometry("500x750")
janela.configure(bg="#121212")  # Fundo escuro

# Estilo moderno
style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Arial", 12), padding=5)
style.configure("TLabel", font=("Arial", 10), foreground="white", background="#121212")

# CFOP
ttk.Label(janela, text="Novo CFOP:").pack(pady=2)
entrada_cfop = ttk.Entry(janela)
entrada_cfop.pack(pady=2)

# Transportadora
ttk.Label(janela, text="--- Dados da Transportadora ---", font=("Arial", 10, "bold")).pack(pady=5)

campos_transportadora = {
    "CNPJ": "cnpj", "IE": "ie", "Razão Social": "xNome", "Nome Fantasia": "xFant",
    "Logradouro": "xLgr", "Número": "nro", "Bairro": "xBairro",
    "Código Município": "cMun", "Município": "xMun", "CEP": "CEP",
    "UF": "UF", "Telefone": "fone"
}

entradas = {}
for label_text, tag in campos_transportadora.items():
    ttk.Label(janela, text=f"Novo {label_text}:").pack(pady=2)
    entrada = ttk.Entry(janela)
    entrada.pack(pady=2)
    entradas[tag] = entrada

# Local de Entrega
ttk.Label(janela, text="Novo Local de Entrega:").pack(pady=2)
entrada_local_entrega = ttk.Entry(janela)
entrada_local_entrega.pack(pady=2)

# Barra de progresso
progresso = ttk.Progressbar(janela, orient="horizontal", length=400, mode="determinate", style="TProgressbar")
progresso.pack(pady=10)

# Label de status da barra de progresso
label_progresso = ttk.Label(janela, text="Aguardando...")
label_progresso.pack(pady=5)

# Botões estilizados
frame_botoes = tk.Frame(janela, bg="#121212")
frame_botoes.pack(pady=10)

botao_executar = ttk.Button(frame_botoes, text="Alterar XML", command=modificar_xml)
botao_executar.pack(side="left", padx=5)

botao_cancelar = ttk.Button(frame_botoes, text="Cancelar", command=cancelar_processo_xml)
botao_cancelar.pack(side="left", padx=5)

janela.mainloop()
