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
    language: 'pt-BR',
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
  for (let i = 0; i < 15; i++) {
    capa.addLineBreak() // Espaço de 9(nove) Enter
  }
  capa.addText('Cidade - Estado', {
    font_face: 'Arial',
    font_size: 12,
  })
  capa.addLineBreak()
  capa.addText('Ano', { font_face: 'Arial', font_size: 12 })
  for (let i = 0; i < 1; i++) {
    capa.addLineBreak()
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

  for (let i = 0; i < 11; i++) {
    folhaDeRostoObjetivo.addLineBreak()
  }
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
  for (let i = 0; i < 1; i++) {
    folhaDeRostoAnoCidade.addLineBreak() // Espaço de 11(onze) Enter
  }
}

function adicionarSumario(docx) {
  const sumario = docx.createP({ align: 'center', spacing: { line: 360 } }) // Espaçamento de 1,5 linhas
  sumario.addText('Adicionar Sumario ', {
    font_size: 12,

    color: '#000000',
  })
  for (let i = 0; i < 33; i++) {
    sumario.addLineBreak()
  }
}

function adicionarConteudo(docx) {
  const titulo = docx.createP({ align: 'justify', spacing: { line: 360 } }) // Espaçamento de 1,5 linhas
  titulo.addText('Titulo do Texto  - Exemplo: Titulo 1', {
    font_face: 'Arial',
    font_size: 12,
    bold: true,
    color: '#000000',
    indentFirstLine: 0,
    numLevel: 0,
    align: 'left',
  })
  titulo.addLineBreak()
  // Conteúdo correspondente aos títulos
  const conteudoTitulo = docx.createP({
    align: 'justify',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas
  conteudoTitulo.addText('Conteúdo do Titulo 1', {
    font_face: 'Arial',
    font_size: 12,
  })
  conteudoTitulo.addText(
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vestibulum maximus aliquam tellus, vitae sollicitudin purus. Integer venenatis hendrerit libero id dictum. In eu dictum massa. Morbi maximus tincidunt ante ac rutrum. Integer feugiat eget justo eget pulvinar. Praesent accumsan libero felis. Mauris euismod luctus euismod. Fusce pulvinar, neque in sodales condimentum, ligula elit malesuada mi, et ornare sem libero eget dolor. Maecenas aliquam non urna congue dictum. Fusce ante orci, elementum sed dui at, fermentum mollis odio. Fusce iaculis egestas mi, in auctor sem maximus sagittis. Fusce finibus velit lorem, ut sodales erat lobortis quis.',
    {
      font_face: 'Arial',
      font_size: 12,
    },
  )
  conteudoTitulo.addLineBreak()

  const subtitulo = docx.createP({ align: 'justify', spacing: { line: 360 } }) // Espaçamento de 1,5 linhas
  subtitulo.addText('Sub-Titulo do Texto - Exemplo: Sub-Titulo 1.1', {
    font_face: 'Arial',
    font_size: 12,
    bold: true,
    color: '#000000',
    indentFirstLine: 0,
    numLevel: 1,
    align: 'left',
  })
  subtitulo.addLineBreak()
  const conteudoSubtitulo = docx.createP({
    align: 'justify',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas
  conteudoSubtitulo.addText('Conteúdo do Sub-Titulo 1.1', {
    font_face: 'Arial',
    font_size: 12,
  })

  conteudoSubtitulo.addText(
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vestibulum maximus aliquam tellus, vitae sollicitudin purus. Integer venenatis hendrerit libero id dictum. In eu dictum massa. Morbi maximus tincidunt ante ac rutrum. Integer feugiat eget justo eget pulvinar. Praesent accumsan libero felis. Mauris euismod luctus euismod. Fusce pulvinar, neque in sodales condimentum, ligula elit malesuada mi, et ornare sem libero eget dolor. Maecenas aliquam non urna congue dictum. Fusce ante orci, elementum sed dui at, fermentum mollis odio. Fusce iaculis egestas mi, in auctor sem maximus sagittis. Fusce finibus velit lorem, ut sodales erat lobortis quis.',
    {
      font_face: 'Arial',
      font_size: 12,
    },
  )
  conteudoSubtitulo.addLineBreak()

  const subSubtitulo = docx.createP({
    align: 'justify',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas
  subSubtitulo.addText(
    'Sub-SubTitulo do Texto - Exemplo: Sub-SubTitulo 1.1.1',
    {
      font_face: 'Arial',
      font_size: 12,
      italic: true,
      color: '#000000',
      indentFirstLine: 0,
      numLevel: 2,
      align: 'left',
    },
  )
  subSubtitulo.addLineBreak()

  const conteudoSubSubtitulo = docx.createP({
    align: 'justify',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas
  conteudoSubSubtitulo.addText('Conteúdo do Sub-SubTitulo 1.1.1', {
    font_face: 'Arial',
    font_size: 12,
  })

  conteudoSubSubtitulo.addText(
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vestibulum maximus aliquam tellus, vitae sollicitudin purus. Integer venenatis hendrerit libero id dictum. In eu dictum massa. Morbi maximus tincidunt ante ac rutrum. Integer feugiat eget justo eget pulvinar. Praesent accumsan libero felis. Mauris euismod luctus euismod. Fusce pulvinar, neque in sodales condimentum, ligula elit malesuada mi, et ornare sem libero eget dolor. Maecenas aliquam non urna congue dictum. Fusce ante orci, elementum sed dui at, fermentum mollis odio. Fusce iaculis egestas mi, in auctor sem maximus sagittis. Fusce finibus velit lorem, ut sodales erat lobortis quis.',
    {
      font_face: 'Arial',
      font_size: 12,
    },
  )
  conteudoSubSubtitulo.addLineBreak()

  // Adicionar espaçamento adicional
  for (let i = 0; i < 28; i++) {
    conteudoSubSubtitulo.addLineBreak()
  }
}

function adicionarPaginaFinal(docx) {
  const bibliografia = docx.createP({
    align: 'justify',
    spacing: { line: 360 },
  })

  bibliografia.addText('Bibliografia', {
    font_size: 12,
    color: '#000',
  })
  bibliografia.addLineBreak()

  const adicionarAutorLivro = (ultimoNome, inicialNome, inicialSobrenome) => {
    const entradaLivro = `${ultimoNome}, ${inicialNome}. ${inicialSobrenome}. `
    bibliografia.addText(entradaLivro, { font_size: 12, font_face: 'Arial' })
  }
  adicionarAutorLivro('Sobrenome', 'I', 'N')

  const adicionarTituloObra = (tituloObra) => {
    const entrataTituloObra = `${tituloObra}. `
    bibliografia.addText(entrataTituloObra, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarTituloObra('Titulo da Obra')

  const adicionarEdicaoDaObra = (edicaoDaObra) => {
    const entrataEdicaoDaObra = `${edicaoDaObra}. `
    bibliografia.addText(entrataEdicaoDaObra, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarEdicaoDaObra('Edição da Obra')

  const adicionarCidade = (cidade) => {
    const entrataCidade = `${cidade}: `
    bibliografia.addText(entrataCidade, { font_size: 12, font_face: 'Arial' })
  }
  adicionarCidade('Cidade onde o Livro foi lançado')

  const adicionarEditoraDaObra = (editoraDaObra) => {
    const entrataEditoraDaObra = `${editoraDaObra}, `
    bibliografia.addText(entrataEditoraDaObra, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarEditoraDaObra('Editora da Obra')

  const adicionarAno = (ano) => {
    const entrataAno = `${ano}. `
    bibliografia.addText(entrataAno, { font_size: 12, font_face: 'Arial' })
  }
  adicionarAno('Ano de Lançamento da Obra')

  for (let i = 0; i < 5; i++) {
    bibliografia.addLineBreak()
  }

  const bibliografiaSite = docx.createP({
    align: 'justify',
    spacing: { line: 360 },
  }) // Espaçamento de 1,5 linhas

  const adicionarAutorSite = (ultimoNome, inicialNome) => {
    const entradaSite = `${ultimoNome}, ${inicialNome}. `
    bibliografiaSite.addText(entradaSite, { font_size: 12, font_face: 'Arial' })
  }
  adicionarAutorSite('Sobrenome', 'I')

  const adicionarTituloSite = (tituloSite) => {
    const entrataTituloSite = `${tituloSite}, `
    bibliografiaSite.addText(entrataTituloSite, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarTituloSite('Titulo da Site')

  const adicionarTituloTrativaSite = (tituloTrativaSite) => {
    const entrataTituloTrativaSite = `${tituloTrativaSite}, `
    bibliografiaSite.addText(entrataTituloTrativaSite, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarTituloTrativaSite('Titulo do que se trata Site')

  const adicionarDisponibilidade = (disponibilidade) => {
    const entrataDisponibilidade = `${disponibilidade}: `
    bibliografiaSite.addText(entrataDisponibilidade, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarDisponibilidade('Disponibilidade do Site')

  const adicionarURL = (URL) => {
    const entrataURL = `<${URL}> `
    bibliografiaSite.addText(entrataURL, { font_size: 12, font_face: 'Arial' })
  }
  adicionarURL('URL do Site')

  const adicionarAnoSite = (anoSite) => {
    const entrataAnoSite = `${anoSite}. `
    bibliografiaSite.addText(entrataAnoSite, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarAnoSite('Ano de Lançamento do Site')

  const adicionarAcessoSite = (DD, MM, AAAA) => {
    const entradaAcessoSite = `Acesso em: ${DD} de ${MM} de ${AAAA} . `
    bibliografiaSite.addText(entradaAcessoSite, {
      font_size: 12,
      font_face: 'Arial',
    })
  }
  adicionarAcessoSite('Dia', 'Mês', 'Ano')
}

criarDocumento()
