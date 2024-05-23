import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re

def obter_texto_paragrafo(para):
    """
    Função para obter o texto de um parágrafo do documento DOCX,
    mantendo o estilo de negrito e itálico quando aplicado.

    Args:
        para (docx.text.paragraph.Paragraph): Parágrafo do documento DOCX.

    Returns:
        str: Texto do parágrafo com estilos de formatação aplicados.
    """
    texto = ""
    if para.runs:
        for run in para.runs:
            texto_run = run.text
            if run.bold:
                texto_run = f"**{texto_run}**"
            if run.italic:
                texto_run = f"*{texto_run}*"
            texto += texto_run
    else:
        texto = para.text
    return texto

def salvar_imagem(bytes_imagem, diretorio_imagem, nome_imagem):
    """
    Função para salvar uma imagem em um diretório específico.

    Args:
        bytes_imagem (bytes): Bytes da imagem a ser salva.
        diretorio_imagem (str): Diretório onde a imagem será salva.
        nome_imagem (str): Nome da imagem a ser salva.

    Returns:
        str: Caminho completo da imagem salva.
    """
    if not os.path.exists(diretorio_imagem):
        os.makedirs(diretorio_imagem)
    caminho_imagem = os.path.join(diretorio_imagem, nome_imagem)
    with open(caminho_imagem, "wb") as arquivo_imagem:
        arquivo_imagem.write(bytes_imagem)
    return caminho_imagem

def converter_docx_para_markdown(caminho_docx, caminho_md_saida, diretorio_imagem):
    """
    Função para converter um documento DOCX para Markdown.

    Args:
        caminho_docx (str): Caminho do arquivo DOCX a ser convertido.
        caminho_md_saida (str): Caminho de saída para o arquivo Markdown convertido.
        diretorio_imagem (str): Diretório onde as imagens serão salvas.
    """
    documento = Document(caminho_docx)
    conteudo_markdown = []

    def adicionar_paragrafo(paragrafo):
        texto = obter_texto_paragrafo(paragrafo)
        if paragrafo.style.name.startswith('Heading'):
            nivel = int(re.search(r'\d+', paragrafo.style.name).group())
            conteudo_markdown.append(f"{'#' * nivel} {texto}\n")
        elif paragrafo.style.name == 'Normal' and paragrafo.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            conteudo_markdown.append(f"<p align='center'>{texto}</p>\n")
        else:
            conteudo_markdown.append(f"{texto}\n")

    contador_imagem = 1
    mapa_imagens = {}

    for relacao in documento.part.rels.values():
        if "image" in relacao.reltype:
            parte_imagem = relacao._target
            bytes_imagem = parte_imagem.blob
            nome_imagem = f"imagem{contador_imagem}.png"
            contador_imagem += 1
            caminho_imagem_salva = salvar_imagem(bytes_imagem, diretorio_imagem, nome_imagem)
            caminho_imagem_relativa = os.path.relpath(caminho_imagem_salva, os.path.dirname(caminho_md_saida))
            mapa_imagens[relacao.rId] = f"![{nome_imagem}]({caminho_imagem_relativa})\n"

    for paragrafo in documento.paragraphs:
        for run in paragrafo.runs:
            if run._element.xpath(".//w:drawing") or run._element.xpath(".//w:pict"):
                for rId in mapa_imagens:
                    if run._element.xpath(f".//*[@r:embed='{rId}']"):
                        conteudo_markdown.append(mapa_imagens[rId])
                        break
            else:
                adicionar_paragrafo(paragrafo)

    for tabela in documento.tables:
        conteudo_markdown.append("\n")
        # Extrair cabeçalhos da tabela
        celulas_cabecalho = tabela.rows[0].cells
        cabecalhos = [celula.text.strip() for celula in celulas_cabecalho]
        conteudo_markdown.append("| " + " | ".join(cabecalhos) + " |")
        conteudo_markdown.append("| " + " | ".join(["---"] * len(cabecalhos)) + " |")
        # Extrair linhas da tabela
        for linha in tabela.rows[1:]:
            dados_linha = [celula.text.strip() for celula in linha.cells]
            conteudo_markdown.append("| " + " | ".join(dados_linha) + " |")
        conteudo_markdown.append("\n")

    with open(caminho_md_saida, "w", encoding="utf-8") as arquivo_md:
        arquivo_md.write("\n".join(conteudo_markdown))

def iniciar_conversao():
    """
    Função para iniciar o processo de conversão de DOCX para Markdown.
    """
    try:
        caminho_docx = filedialog.askopenfilename(filetypes=[("Arquivos Word", "*.docx")])
        diretorio_saida = filedialog.askdirectory()
        caminho_md_saida = os.path.join(diretorio_saida, "saida.md")
        diretorio_imagem = os.path.join(diretorio_saida, "media")
        converter_docx_para_markdown(caminho_docx, caminho_md_saida, diretorio_imagem)
        messagebox.showinfo("Sucesso", "Conversão realizada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    raiz = tk.Tk()
    raiz.geometry("300x150")
    botao_converter = tk.Button(raiz, text="Converter Arquivo Word Para MarkDown", command=iniciar_conversao)
    botao_converter.pack()
    raiz.mainloop()
