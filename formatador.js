const officegen = require('officegen')
const fs = require('fs')

// Função para criar o documento
function criarDocumento() {
  // Criar um documento Word
  const docx = officegen({
    type: 'docx',
    creator: 'Nome do Criador',
    description: 'Descrição do Documento',
    title: 'Título do Documento',
    subject: 'Assunto do Documento',
    keywords: 'Palavras-chave do Documento',
    category: 'Categoria do Documento',
    pageMargins: {
      top: 1701, // 3 cm
      right: 1134, // 2 cm
      bottom: 1134, // 2 cm
      left: 1701, // 3 cm
    },
    pageSize: 'A4',
  })
  // Adicionar páginas ao documento
  adicionarCapa(docx)
  adicionarFolhaDeRostoIntegrante(docx)
  adicionarFolhaDeRostoTitulos(docx)
  adicionarFolhaDeRostoObjetivo(docx)
  adicionarFolhaDeRostoAnoCidade(docx)
  adicionarSumario(docx)
  adicionarConteudo(docx)
  adicionarPaginaFinal(docx)

  // Salvar o documento
  const nomeArquivo = 'template_abnt.docx'
  const outputStream = fs.createWriteStream(nomeArquivo)
  docx.generate(outputStream)
}

// Funções para adicionar elementos ao documento
function adicionarCapa(docx) {
  const capa = docx.createP({ align: 'center', spacing: { line: 360 } }) // Espaçamento de 1,5 linhas
  capa.addText('Nome da Instituição de Ensino', {
    font_face: 'Arial',
    font_size: 12,
  })
  capa.addLineBreak()
  capa.addText('Curso', { font_face: 'Arial', font_size: 12 })
  for (let i = 0; i < 7; i++) {
    capa.addLineBreak() // Espaço de 7(sete) Enter
  }
  capa.addText('Integrante 1', {
    font_face: 'Arial',
    font_size: 12,
  })
  capa.addText(', Integrante 2', {
    font_face: 'Arial',
    font_size: 12,
  })
  capa.addText(', Integrante 3', {
    font_face: 'Arial',
    font_size: 12,
  })
  for (let i = 0; i < 7; i++) {
    capa.addLineBreak() // Espaço de 7(sete) Enter
  }
  capa.addText('Título do Trabalho', {
    font_face: 'Arial',
    font_size: 12,
  })
  capa.addLineBreak()
  capa.addText('Subtítulo do Trabalho', { font_face: 'Arial', font_size: 12 })
  for (let i = 0; i < 9; i++) {
    capa.addLineBreak() // Espaço de 9(nove) Enter
  }
  capa.addText('Cidade - Estado', {
    font_face: 'Arial',
    font_size: 12,
  })
  capa.addLineBreak()
  capa.addText('Ano', { font_face: 'Arial', font_size: 12 })
  for (let i = 0; i < 11; i++) {
    capa.addLineBreak() // Espaço de 7(sete) Enter
  }
}

function adicionarFolhaDeRostoIntegrante(docx) {
  const folhaDeRostoIntegrante = docx.createP({
    align: 'center',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas

  folhaDeRostoIntegrante.addText('Integrante 1', {
    font_face: 'Arial',
    font_size: 12,
  })
  folhaDeRostoIntegrante.addText(', Integrante 2', {
    font_face: 'Arial',
    font_size: 12,
  })
  folhaDeRostoIntegrante.addText(', Integrante 3', {
    font_face: 'Arial',
    font_size: 12,
  })
  for (let i = 0; i < 7; i++) {
    folhaDeRostoIntegrante.addLineBreak() // Espaço de 7(sete) Enter
  }
}

function adicionarFolhaDeRostoTitulos(docx) {
  const folhaDeRostoTitulo = docx.createP({
    align: 'center',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas

  folhaDeRostoTitulo.addText('Título do Trabalho', {
    font_face: 'Arial',
    font_size: 12,
  })
  folhaDeRostoTitulo.addLineBreak()
  folhaDeRostoTitulo.addText('Subtítulo do Trabalho', {
    font_face: 'Arial',
    font_size: 12,
  })
  for (let i = 0; i < 7; i++) {
    folhaDeRostoTitulo.addLineBreak() // Espaço de 7(sete) Enter
  }
}

function adicionarFolhaDeRostoObjetivo(docx) {
  const folhaDeRostoObjetivo = docx.createP({
    align: 'right',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas

  folhaDeRostoObjetivo.addText('Texto do objetivo do trabalho...', {
    font_face: 'Arial',
    font_size: 10,
    indentLeft: 4536, // Recuo esquerdo de 8,00cm
  })
  folhaDeRostoObjetivo.addLineBreak()
  folhaDeRostoObjetivo.addText('Nome do professor solicitante do trabalho', {
    font_face: 'Arial',
    font_size: 10,
    indentLeft: 4536, // Recuo esquerdo de 8,00cm
  })

  folhaDeRostoObjetivo.addLineBreak() // Adiciona um espaço entre o objetivo do trabalho e a cidade - estado
}

function adicionarFolhaDeRostoAnoCidade(docx) {
  const folhaDeRostoAnoCidade = docx.createP({
    align: 'center',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas

  folhaDeRostoAnoCidade.addText('Cidade - Estado', {
    font_face: 'Arial',
    font_size: 12,
  })
  folhaDeRostoAnoCidade.addLineBreak()
  folhaDeRostoAnoCidade.addText('Ano', {
    font_face: 'Arial',
    font_size: 12,
  })
  for (let i = 0; i < 11; i++) {
    folhaDeRostoAnoCidade.addLineBreak() // Espaço de 11(onze) Enter
  }
}

function adicionarSumario(docx) {
  const sumario = docx.createP({ align: 'center', spacing: { line: 360 } }) // Espaçamento de 1,5 linhas
  sumario.addText('Adicionar Sumario ', {
    font_size: 12,

    color: '#000000',
  })
  for (let i = 0; i < 11; i++) {
    sumario.addLineBreak() // Espaço de 7(sete) Enter
  }
}

function adicionarConteudo(docx) {
  const conteudo = docx.createP({ align: 'justify', spacing: { line: 360 } }) // Espaçamento de 1,5 linhas
  conteudo.addText('Titulo do Texto  - Exemplo: Titulo 1', {
    font_size: 12,

    color: '#000000',
  })
  conteudo.addLineBreak()
  conteudo.addText('Conteúdo do Titulo 1', { font_size: 12 })
  conteudo.addLineBreak()
  conteudo.addText('Sub-Titulo do Texto - Exemplo: Sub-Titulo 1.1', {
    font_size: 12,
    color: '#000000',
  })
  conteudo.addLineBreak()
  conteudo.addText('Conteúdo do Sub-Titulo 1.1', { font_size: 12 })
  conteudo.addLineBreak()
  conteudo.addText('Sub-SubTitulo do Texto - Exemplo: Sub-SubTitulo 1.1.1', {
    font_size: 12,
    italic: true,
    color: '#000000',
  })
  conteudo.addLineBreak()
  conteudo.addText('Conteúdo do Sub-SubTitulo 1.1.1', { font_size: 12 })
  conteudo.addLineBreak()
}

function adicionarPaginaFinal(docx) {
  // Criar um parágrafo para a seção da bibliografia
  const bibliografia = docx.createP({
    align: 'justify',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas

  // Adicionar o título "Bibliografia" ao parágrafo
  bibliografia.addText('Bibliografia', {
    font_size: 12,

    color: '#000',
  })
  bibliografia.addLineBreak()

  // Definir uma função para adicionar entradas de livro à bibliografia
  const adicionarAutorLivro = (ultimoNome, inicialNome, inicialSobrenome) => {
    const entradaLivro = `${ultimoNome}, ${inicialNome}. ${inicialSobrenome}. ` // Montar a entrada do livro
    bibliografia.addText(entradaLivro, { font_size: 12 }) // Adicionar a entrada do livro ao parágrafo
  }
  // Exemplo de chamada da função para adicionar uma entrada de livro
  adicionarAutorLivro('Sobrenome', 'I', 'N')

  const adicionarTituloObra = (tituloObra) => {
    const entrataTituloObra = `${tituloObra}. `
    bibliografia.addText(entrataTituloObra, { font_size: 12 })
  }
  adicionarTituloObra('Titulo da Obra')

  const adicionarEdicaoDaObra = (edicaoDaObra) => {
    const entrataEdicaoDaObra = `${edicaoDaObra}. `
    bibliografia.addText(entrataEdicaoDaObra, { font_size: 12 })
  }
  adicionarEdicaoDaObra('Edição da Obra')

  const adicionarCidade = (cidade) => {
    const entrataCidade = `${cidade}: `
    bibliografia.addText(entrataCidade, { font_size: 12 })
  }
  adicionarCidade('Cidade onde o Livro foi lançado')

  const adicionarEditoraDaObra = (editoraDaObra) => {
    const entrataEditoraDaObra = `${editoraDaObra}, `
    bibliografia.addText(entrataEditoraDaObra, { font_size: 12 })
  }
  adicionarEditoraDaObra('Editora da Obra')

  const adicionarAno = (ano) => {
    const entrataAno = `${ano}. `
    bibliografia.addText(entrataAno, { font_size: 12 })
  }
  adicionarAno('Ano de Lançamento da Obra')
}

criarDocumento()
