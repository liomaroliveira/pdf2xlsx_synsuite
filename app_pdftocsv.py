import os
import re
import datetime
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Funções de processamento

def processar_tarefa(task):
    tarefa_line = task['tarefa_line']
    descricao = "\n".join(task['descricao_lines']).strip()
    relato = task.get('relato_line', '')

    protocolo = ''
    m = re.search(r'Protocolo (\d+)', tarefa_line)
    if m:
        protocolo = m.group(1)

    partes = tarefa_line.split(" - ")
    tipo_solicitacao = partes[1].strip() if len(partes) >= 2 else ''
    cliente = " - ".join(partes[2:]).strip() if len(partes) >= 3 else ''

    data_abertura = ''
    m = re.search(r'(\d{2}/\d{2}/\d{4}) - (\d{2}:\d{2})', relato)
    if m:
        dt = datetime.datetime.strptime(f"{m.group(1)} {m.group(2)}", "%d/%m/%Y %H:%M")
        data_abertura = dt.strftime("%d/%m/%Y  %H:%M:%S")

    return {
        "QTD": 1,
        "ID Atendimento": "Relato",
        "Protocolo": protocolo,
        "Data Abertura": data_abertura,
        "Cliente": cliente,
        "Tipo Solicitação": tipo_solicitacao,
        "Tipo Pessoa": "",
        "Data Encerramento": data_abertura,
        "Tempo": "",
        "Média Horas": "",
        "Status": "Inclusão de Relato",
        "Contexto": descricao,
        "Problema": "",
        "Solução": "Padrão",
        "Catálogo de Serviço": "Protocolo" if protocolo else "Tarefa",
        "SLA": "",
        "Equipe": "ServiceDesk N2",
        "Setor": "T.I.",
        "Item de Serviço": "",
        "Cidade": "",
        "Bairro": "",
        "Categoria 1": "",
        "Categoria 2": "",
        "Categoria 3": "",
        "Categoria 4": "",
        "Categoria 5": "",
    }


def extrair_tarefas_de_pdf(caminho_pdf, file_progress=None):
    tasks = []
    current = None
    state = None
    with pdfplumber.open(caminho_pdf) as pdf:
        total_pages = len(pdf.pages)
        if file_progress:
            file_progress.config(maximum=total_pages)
            file_progress['value'] = 0
        for idx, page in enumerate(pdf.pages, 1):
            for line in page.extract_text().split("\n"):
                if line.startswith("Tarefa:"):
                    if current:
                        tasks.append(processar_tarefa(current))
                    current = {"tarefa_line": line, "descricao_lines": [], "relato_line": ""}
                    state = None
                elif current is not None:
                    if line.startswith("Descrição:"):
                        state = "descricao"
                    elif line.startswith("Relato"):
                        if not current["relato_line"]:
                            current["relato_line"] = line
                        state = None
                    else:
                        if state == "descricao":
                            current["descricao_lines"].append(line)
            if file_progress:
                file_progress['value'] = idx
                root.update_idletasks()
    if current:
        tasks.append(processar_tarefa(current))
    return tasks


def converter_pdfs(lista_pdfs, pasta_saida, total_progress, file_progress, salva_csv, salva_xlsx):
    total = len(lista_pdfs)
    total_progress.config(maximum=total)
    total_progress['value'] = 0

    if not salva_csv and not salva_xlsx:
        messagebox.showwarning("Aviso", "Selecione ao menos um formato de saída.")
        return

    for idx, pdf_path in enumerate(lista_pdfs, 1):
        destino = pasta_saida or os.path.dirname(pdf_path)
        os.makedirs(destino, exist_ok=True)

        tarefas = extrair_tarefas_de_pdf(pdf_path, file_progress)
        df = pd.DataFrame(tarefas)
        nome_base = os.path.splitext(os.path.basename(pdf_path))[0]

        if salva_csv:
            csv_path = os.path.join(destino, f"{nome_base}.csv")
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        if salva_xlsx:
            try:
                xlsx_path = os.path.join(destino, f"{nome_base}.xlsx")
                df.to_excel(xlsx_path, index=False)
            except ImportError:
                messagebox.showerror("Erro", "openpyxl não está instalado. Execute: pip install openpyxl")
                return

        total_progress['value'] = idx
        root.update_idletasks()

    messagebox.showinfo(
        "Sucesso",
        f"Convertidos {total} PDF(s).\nArquivos salvo em:\n{pasta_saida or 'origem de cada PDF'}"
    )

# Interface Gráfica
root = tk.Tk()
root.title("PDF 2 CSV - RELATOS VOALLE")
root.geometry("700x600")

lbl1 = tk.Label(root, text="1) Selecionar os arquivos PDF:", font=("Arial", 12, "bold"), justify="center")
lbl1.pack(pady=(10, 5), anchor="center")

frame_files = tk.Frame(root, bd=1, relief="sunken", width=650, height=140)
frame_files.pack(padx=10, pady=5)
frame_files.pack_propagate(False)

listbox_pdfs = tk.Listbox(frame_files, height=6)
scrollbar_pdfs = tk.Scrollbar(frame_files, orient="vertical", command=listbox_pdfs.yview)
listbox_pdfs.config(yscrollcommand=scrollbar_pdfs.set)
listbox_pdfs.pack(side="left", fill="both", expand=True)
scrollbar_pdfs.pack(side="right", fill="y")

def selecionar_pdfs():
    arquivos = filedialog.askopenfilenames(title="Selecione um ou mais PDFs", filetypes=[("PDF files", "*.pdf")])
    if arquivos:
        listbox_pdfs.delete(0, tk.END)
        for f in arquivos:
            listbox_pdfs.insert(tk.END, f)
        # Define o caminho da pasta do primeiro PDF selecionado
        var_saida.set(os.path.dirname(arquivos[0]))

btn_select_pdfs = tk.Button(root, text="Selecionar Arquivos PDF", command=selecionar_pdfs)
btn_select_pdfs.pack(pady=5, anchor="center")

lbl2 = tk.Label(root, text="2) Selecionar a pasta de destino:", font=("Arial", 12, "bold"), justify="center")
lbl2.pack(pady=(15, 5), anchor="center")

var_saida = tk.StringVar()
entry_saida = tk.Entry(root, textvariable=var_saida, width=60, state="readonly")
entry_saida.pack(padx=10, pady=5)

def selecionar_saida():
    pasta = filedialog.askdirectory(title="Selecione a pasta de destino")
    if pasta:
        var_saida.set(pasta)
    elif listbox_pdfs.size() > 0:
        var_saida.set(os.path.dirname(listbox_pdfs.get(0)))

btn_select_saida = tk.Button(root, text="Salvar Arquivos em...", command=selecionar_saida)
btn_select_saida.pack(pady=5, anchor="center")

# Checkbuttons para formatos
format_frame = tk.Frame(root)
format_frame.pack(pady=(10,5))
var_csv = tk.BooleanVar(value=True)
var_xlsx = tk.BooleanVar(value=False)
chk_csv = tk.Checkbutton(format_frame, text="Salvar em CSV", variable=var_csv)
chk_xlsx = tk.Checkbutton(format_frame, text="Salvar em Excel (.xlsx)", variable=var_xlsx)
chk_csv.pack(side="left", padx=10)
chk_xlsx.pack(side="left", padx=10)

lbl_file_prog = tk.Label(root, text="Progresso por arquivo:")
lbl_file_prog.pack(pady=(20, 2))
file_progress = ttk.Progressbar(root, length=600, mode='determinate')
file_progress.pack(pady=2)

lbl_total_prog = tk.Label(root, text="Progresso total:")
lbl_total_prog.pack(pady=(15, 2))
total_progress = ttk.Progressbar(root, length=600, mode='determinate')
total_progress.pack(pady=2)

btn_converter = tk.Button(
    root,
    text="3) Converter arquivos...",
    bg="#4CAF50",
    fg="white",
    font=("Arial", 11, "bold"),
    command=lambda: converter_pdfs(
        listbox_pdfs.get(0, tk.END),
        var_saida.get().strip(),
        total_progress,
        file_progress,
        var_csv.get(),
        var_xlsx.get()
    )
)
btn_converter.pack(pady=20, anchor="center")

root.mainloop()
