# WordToMd

WordToMd é uma aplicação desenvolvida para facilitar o desenvolvimento de arquivos Markdown (.md) apartir de um documento Word. A aplicação reconhece e converte a formatação do documento Word, incluindo negrito, itálico, links e imagens, para um arquivo Markdown compatível com GitHub Pages.

## Funcionalidades

- Conversão de documentos Word (.docx) para Markdown (.md)
- Manutenção de estilos de formatação (negrito, itálico)
- Identificação e formatação de links
- Extração e salvamento de imagens
- Interface gráfica amigável usando `customtkinter`
- Suporte para conversão de múltiplos arquivos em lote

## Requisitos

- Python
- Bibliotecas Python Utilizadas:
  - `tkinter`
  - `python-docx`
  - `Pillow`
  - `tldextract`
  - `customtkinter`
  - `pyinstaller` (para criar o executável)

## Instalação

1. Clone o repositório:
   ```sh
   git clone https://github.com/seu-usuario/WordToMd.git
   cd WordToMd

2. Execute o arquivo main.py ou use o executavel.
    ```python
    python main.py