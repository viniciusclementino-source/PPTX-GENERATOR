"TÍTULO","DESCRIÇÃO"
"🧩 Gerador de Apresentações PPTX v2","Aplicativo Python + Tkinter que gera apresentações PowerPoint (.pptx) automaticamente a partir de comandos de texto e imagens. Ideal para criação rápida de slides com layouts personalizados e controle de conteúdo via interface gráfica simples."

"---"
"🖼️ Screenshot do App"
"[COLE UM PRINT DA INTERFACE DO APP AQUI]"
"(Recomendação: tire um print, envie para o repositório e substitua este texto pelo link da imagem)"
"---"

"🚀 Tecnologias Utilizadas","Este projeto foi construído com as seguintes ferramentas:

Python 3: Linguagem principal do projeto

Tkinter: Para a interface gráfica (GUI)

python-pptx: Biblioteca para criar e manipular os arquivos .pptx [source: 8]

PyInstaller: Utilizado nos scripts de build para gerar o executável [source: 1, 6]"

"⚙️ Instalação (para Desenvolvimento)","Para rodar o projeto em sua máquina local:

Clone o repositório:
git clone https://github.com/viniciusclementino-source/PPTX-GENERATOR.git
cd PPTX-GENERATOR

Crie e ative um ambiente virtual (Recomendado):

No Windows

python -m venv venv
.\venv\Scripts\activate

No macOS/Linux

python3 -m venv venv
source venv/bin/activate

Instale as dependências do projeto:
pip install -r requirements.txt [source: 8]

Execute o aplicativo:
python main.py [source: 7]"

"📦 Gerando o Executável (.exe / Binário)","O repositório inclui scripts para compilar o app em um único executável (usando PyInstaller [source: 1, 6]), facilitando a distribuição para quem não tem Python.

No Windows:
Basta executar o script de build:
.\build.bat [source: 1]

No macOS/Linux:
Dê permissão de execução e rode o script:
chmod +x build.sh
./build.sh [source: 6]

O executável final aparecerá na pasta dist/."

"🧠 Como o app funciona","O aplicativo lê:

Um arquivo de tema (.pptx) — usado como modelo base.

Uma lista de imagens — aplicadas automaticamente conforme o layout.

Um campo de comandos de texto, onde você descreve cada slide.
A saída é uma nova apresentação PowerPoint gerada de forma automatizada."

"🖋️ Manual de Formatação dos Comandos","Os comandos devem ser inseridos um por linha, com campos separados por '|'.

layout | título | texto_ou_legendas"

"🧩 Estrutura geral","layout: Define o tipo de slide (ex: img left custom, img6)
título: Título do slide
texto_ou_legendas: Texto principal (ou legendas, dependendo do layout)"

"🧱 Layouts disponíveis","1. img left custom → Imagem à esquerda, texto à direita
   Exemplo: img left custom | Título | Texto principal

img top custom → Imagem no topo, texto abaixo
   Exemplo: img top custom | Introdução | Texto do conteúdo

img right custom → Imagem à direita, texto à esquerda
   Exemplo: img right custom | Tema | Texto explicativo

img2 → Duas imagens lado a lado abaixo do texto
   Exemplo: img2 | Título | Texto explicativo

img6 → Seis imagens (2x3) com legendas
   Exemplo: img6 | Título | Legenda 1 | Legenda 2 | ... | Legenda 6
   ⚠️ Requer 6 imagens e até 6 legendas"

"✍️ Separadores de texto","Use '///' para criar quebras de parágrafo.
Exemplo: Primeiro parágrafo /// Segundo parágrafo"

"🎨 Formatação avançada de texto (tags)","O app reconhece tags HTML-like:
<b>texto</b> → Negrito
<i>texto</i> → Itálico
<u>texto</u> → Sublinhado
<s>texto</s> → Tachado
<b:cor>texto</b:cor> → Negrito colorido

Exemplo completo:
img left custom | Formatação | <b:azul>Texto em azul</b:azul> /// <i>Texto em itálico</i>"

"🧭 Manual dos Botões","Selecionar... (Tema): Escolhe o arquivo base .pptx
Adicionar...: Adiciona uma ou mais imagens (.png, .jpg)
Remover: Exclui imagens selecionadas
Cima / Baixo: Move imagens na lista
Limpar Tudo: Limpa tema, imagens e comandos
Gerar Apresentação: Cria e salva o .pptx final"

"📂 Estrutura do Projeto","main.py → Código principal do app (lógica e GUI) [source: 7]
requirements.txt → Lista de dependências Python [source: 8]
build.bat → Script de build para Windows (gera .exe) [source: 1]
build.sh → Script de build para macOS/Linux [source: 6]
README.md → Este manual
assets/ → (opcional) pasta para temas e imagens"

"💡 Dicas de Uso","- A ordem das imagens define a sequência dos slides.

Se faltar imagem para um layout, o app exibe aviso.

Slides do tema são removidos automaticamente do resultado.

Use um tema .pptx limpo (layout branco no índice 6)."

"🧑‍💻 Autor","Desenvolvido por Vinícius Martins Clementino — Ferramenta para geração automatizada de apresentações didáticas em PowerPoint."

"📜 Licença","Projeto sob licença MIT. Livre para uso, modificação e redistribuição."
