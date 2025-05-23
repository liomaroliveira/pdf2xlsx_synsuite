# 📄 PDF 2 CSV - RELATOS VOALLE

Uma aplicação com interface gráfica para extrair tarefas e relatos de arquivos PDF gerados pelo sistema Voalle e convertê-los para os formatos **CSV** e **Excel (.xlsx)**.

---

## 🚀 Funcionalidades

- Leitura automática de arquivos PDF com estrutura de **Tarefa / Descrição / Relato**.
- Extração de informações como **protocolo**, **tipo de solicitação**, **cliente** e **data de abertura**.
- Exportação para arquivos `.csv` e `.xlsx` com estrutura compatível para relatórios.
- Interface gráfica simples e intuitiva (construída com `tkinter`).
- Barra de progresso por arquivo e progresso total da conversão.

---

## 📸 Interface

<p align="center">
  <img src="https://via.placeholder.com/700x400.png?text=Exemplo+da+Interface+Gr%C3%A1fica" alt="Interface gráfica">
</p>

---

## 🛠️ Requisitos

Certifique-se de ter o **Python 3.7+** instalado.

⚙️ Tecnologias usadas:
pdfplumber
pandas
openpyxl
tkinter

### Instalar dependências:

```bash
pip install pdfplumber pandas openpyxl

### Etapas na interface

```python
python seu_arquivo.py