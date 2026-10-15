# GeraDoc

O GeraDoc é um utilitário de automação desenvolvido em Python com interface gráfica (Tkinter) projetado para simplificar a criação em massa de documentos do Word (.docx) a partir de planilhas de dados (.xlsx), permitindo também a conversão direta e segura para PDF.

## Funcionalidades

- Geração Individual: Cria um documento personalizado para cada linha da planilha.
- Agrupamento por Colunas: Agrupa registros correspondentes para gerar relatórios consolidados.
- Arquivo Único: Junta todos os documentos gerados em um único arquivo .docx ao final.
- Conversão Segura para PDF: Utiliza automação COM em segundo plano do Microsoft Word para converter os documentos gerados sem erros de terminal.
- Máscaras de Formatação: Formatação automática de CPFs, CNPJs, telefones e valores monetários baseada em tags.

## Estrutura do Repositório

GeraDoc/
├── .gitignore
├── README.md
├── requirements.txt
├── GeraDoc.py
└── docs/
    ├── manual.pdf
    └── modelo-exemplo.docx

## Pré-requisitos

- Python 3.8 ou superior instalado.
- Microsoft Word instalado (necessário para a conversão de arquivos para PDF via win32com).

## Instalação e Execução

1. Clone o repositório ou baixe os arquivos:
git clone https://github.com/ElbioTeixeira/GeraDoc.git
cd GeraDoc

2. Instale as dependências necessárias:
pip install -r requirements.txt

3. Execute o programa:
python GeraDoc.py

## Compilação do Executável (PyInstaller)

Caso queira gerar o seu próprio arquivo executável (.exe) a partir do código-fonte:

1. Instale o PyInstaller:
pip install pyinstaller

2. Compile o script em modo de arquivo único (onefile) e com interface oculta de console (noconsole):
pyinstaller --noconsole --onefile GeraDoc.py

O executável final será gerado dentro da pasta "dist/".

## Documentação e Exemplos

- Manual de Utilização: Instruções detalhadas de formatação de tags e uso da interface estão disponíveis em docs/manual.pdf.
- Modelo de Exemplo: Utilize o arquivo docs/modelo-exemplo.docx como base para estruturar as suas tags de substituição.

## Licença

Este projeto é distribuído sob a licença MIT. Sinta-se à vontade para utilizar, modificar e contribuir.