import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from datetime import datetime

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

def salvar_imagem(bytes_imagem, diretorio_imagem, nome_imagem, nome_arquivo):
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
    nome_imagem = f"{nome_arquivo}_{nome_imagem}"
    caminho_imagem = os.path.join(diretorio_imagem, nome_imagem)
    with open(caminho_imagem, "wb") as arquivo_imagem:
        arquivo_imagem.write(bytes_imagem)
    return caminho_imagem

def converter_docx_para_markdown(caminho_docx, caminho_md_saida, diretorio_imagem, nome_arquivo, tipo_cabecalho):
    """
    Função para converter um documento DOCX para Markdown.

    Args:
        caminho_docx (str): Caminho do arquivo DOCX a ser convertido.
        caminho_md_saida (str): Caminho de saída para o arquivo Markdown convertido.
        diretorio_imagem (str): Diretório onde as imagens serão salvas.
    """
    documento = Document(caminho_docx)
    conteudo_markdown = []

    # Adicionando o cabeçalho do arquivo Markdown
    data_atual = datetime.now()
    dia = f'{data_atual.day:02d}'
    mes = f'{data_atual.month:02d}'
    ano = f'{data_atual.year:04d}'

    if(tipo_cabecalho == "indice"):
        conteudo_markdown.append('---')
        conteudo_markdown.append('title: "Titulo da sua documentação"')
        conteudo_markdown.append('type: docs')
        conteudo_markdown.append('last_updated: ')
        conteudo_markdown.append(f'\tdate: "{mes}/{dia}/{ano}"')
        conteudo_markdown.append('\tauthor: "Seu nome"')
        conteudo_markdown.append('sidebar_position: 1')
        conteudo_markdown.append('---\n')
    else:
        conteudo_markdown.append('---')
        conteudo_markdown.append('title: "Titulo da sua documentação"')
        conteudo_markdown.append('type: docs')
        conteudo_markdown.append('menu: ')
        conteudo_markdown.append('\tmain:')
        conteudo_markdown.append('\t\tsidebar_position: 1')
        conteudo_markdown.append('description: "Descrição da sua documentação"')
        conteudo_markdown.append('---\n')

    # Definindo a função para adicionar parágrafos
    def adicionar_paragrafo(paragrafo):
        texto = obter_texto_paragrafo(paragrafo)
        if paragrafo.style.name.startswith('Heading'):
            nivel = int(re.search(r'\d+', paragrafo.style.name).group())
            conteudo_markdown.append(f"{'#' * nivel} {texto}\n")
        elif paragrafo.style.name == 'Normal' and paragrafo.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            conteudo_markdown.append(f"<p align='center'>{texto}</p>\n")
        else:
            conteudo_markdown.append(f"{texto}\n")

    # Definindo a função para adicionar tabelas
    def adicionar_tabela(tabela):
        conteudo_markdown.append("\n")
        # Extrair cabeçalhos da tabela
        cabecalhos = [celula.text.strip() for celula in tabela.rows[0].cells]
        conteudo_markdown.append("| " + " | ".join(cabecalhos) + " |")
        conteudo_markdown.append("| " + " | ".join(["---"] * len(cabecalhos)) + " |")
        # Extrair linhas da tabela
        for linha in tabela.rows[1:]:
            dados_linha = [celula.text.strip() for celula in linha.cells]
            conteudo_markdown.append("| " + " | ".join(dados_linha) + " |")
        conteudo_markdown.append("\n")

    contador_imagem = 1
    mapa_imagens = {}

    # Processamento de imagens
    for relacao in documento.part.rels.values():
        if "image" in relacao.reltype:
            parte_imagem = relacao._target
            bytes_imagem = parte_imagem.blob
            nome_imagem = f"imagem{contador_imagem}.png"
            contador_imagem += 1
            caminho_imagem_salva = salvar_imagem(bytes_imagem, diretorio_imagem, nome_imagem, nome_arquivo)
            caminho_imagem_relativa = os.path.relpath(caminho_imagem_salva, os.path.dirname(caminho_md_saida))
            mapa_imagens[relacao.rId] = f"![{nome_imagem}]({caminho_imagem_relativa})\n"

    # Processamento do conteúdo do documento
    for elemento in documento.element.body:
        if isinstance(elemento, CT_P):
            for paragrafo in documento.paragraphs:
                if paragrafo._element == elemento:
                    tem_imagem = any(run._element.xpath(".//w:drawing") or run._element.xpath(".//w:pict") for run in paragrafo.runs)
                    if tem_imagem:
                        for run in paragrafo.runs:
                            if run._element.xpath(".//w:drawing") or run._element.xpath(".//w:pict"):
                                for rId in mapa_imagens:
                                    if run._element.xpath(f".//*[@r:embed='{rId}']"):
                                        conteudo_markdown.append(mapa_imagens[rId])
                                        break
                    else:
                        adicionar_paragrafo(paragrafo)
                    break
        elif isinstance(elemento, CT_Tbl):
            for tabela in documento.tables:
                if tabela._element == elemento:
                    adicionar_tabela(tabela)
                    break

    # Escrever o conteúdo no arquivo Markdown
    with open(caminho_md_saida, "w", encoding="utf-8") as arquivo_md:
        arquivo_md.write("\n".join(conteudo_markdown))

def iniciar_conversao():
    """
    Função para iniciar o processo de conversão de DOCX para Markdown.
    """
    try:
        caminho_docx = filedialog.askopenfilename(filetypes=[("Arquivos Word", "*.docx")])
        nome_arquivo = os.path.splitext(os.path.basename(caminho_docx))[0]
        diretorio_saida = filedialog.askdirectory()
        caminho_md_saida = os.path.join(diretorio_saida, f"{nome_arquivo}.md")
        diretorio_imagem = os.path.join(diretorio_saida, f"img_{nome_arquivo}")
        tipo_cabecalho = var.get()
        converter_docx_para_markdown(caminho_docx, caminho_md_saida, diretorio_imagem, nome_arquivo, tipo_cabecalho)
        messagebox.showinfo("Sucesso", "Conversão realizada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    raiz = tk.Tk()
    raiz.geometry("300x150")

    var = tk.StringVar(value = "indice")
    r1 = tk.Radiobutton(raiz, text="Índice", variable=var, value="indice")
    r1.pack()
    r2 = tk.Radiobutton(raiz, text="Sub Índice", variable=var, value="subindice")
    r2.pack()

    botao_converter = tk.Button(raiz, text="Converter Arquivo Word Para MarkDown", command=iniciar_conversao)
    botao_converter.pack()
    raiz.mainloop()
