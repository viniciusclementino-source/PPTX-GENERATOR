"TÍTULO","DESCRIÇÃO"
"🧩 Gerador de Apresentações PPTX v2","Aplicativo Python + Tkinter que gera apresentações PowerPoint (.pptx) automaticamente a partir de comandos de texto e imagens. Ideal para criação rápida de slides com layouts personalizados e controle de conteúdo via interface gráfica simples."

"⚙️ Instalação","1. Clone o repositório:
git clone https://github.com/seuusuario/gerador-pptx.git
cd gerador-pptx

2. Instale as dependências:
pip install python-pptx

3. Execute o aplicativo:
python main.py"

"🧠 Como o app funciona","O aplicativo lê:
- Um arquivo de tema (.pptx) — usado como modelo base.
- Uma lista de imagens — aplicadas automaticamente conforme o layout.
- Um campo de comandos de texto, onde você descreve cada slide.
A saída é uma nova apresentação PowerPoint gerada de forma automatizada."

"🖋️ Manual de Formatação dos Comandos","Os comandos devem ser inseridos um por linha, com campos separados por '|'.

layout | título | texto_ou_legendas"

"🧩 Estrutura geral","layout: Define o tipo de slide (ex: img left custom, img6)
título: Título do slide
texto_ou_legendas: Texto principal (ou legendas, dependendo do layout)"

"🧱 Layouts disponíveis","1. img left custom → Imagem à esquerda, texto à direita
   Exemplo: img left custom | Título | Texto principal

2. img top custom → Imagem no topo, texto abaixo
   Exemplo: img top custom | Introdução | Texto do conteúdo

3. img right custom → Imagem à direita, texto à esquerda
   Exemplo: img right custom | Tema | Texto explicativo

4. img2 → Duas imagens lado a lado abaixo do texto
   Exemplo: img2 | Título | Texto explicativo

5. img6 → Seis imagens (2x3) com legendas
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

"📦 Estrutura do Projeto","main.py → Código principal
README.md → Este manual
assets/ → (opcional) pasta para temas e imagens"

"💡 Dicas de Uso","- A ordem das imagens define a sequência dos slides.
- Se faltar imagem para um layout, o app exibe aviso.
- Slides do tema são removidos automaticamente do resultado.
- Use um tema .pptx limpo (layout branco no índice 6)."

"🧑‍💻 Autor","Desenvolvido por Vinícius Martins Clementino — Ferramenta para geração automatizada de apresentações didáticas em PowerPoint."

"📜 Licença","Projeto sob licença MIT. Livre para uso, modificação e redistribuição."
