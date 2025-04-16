// Função principal que gera os slides
function gerarSlides() {
    const PRESENTATION_ID = "SEU_ID_APRESENTACAO";
    const SHEET_ID = "SEU_ID_DO_SHEETS";
    const sheetName = "SHEET_NAME";
  
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    const presentation = SlidesApp.openById(PRESENTATION_ID);
  
    // Configuração das colunas no Sheets
    const HEADER_ROW = 0; // Linha do cabeçalho (indexada a 0)
    const colLogo = 0; // Coluna A
    const colTema = 1; // Coluna B
    const colTitulo = 2; // Coluna C
    const colResumo = 3; // Coluna D
    const colDestaques = 4; // Coluna E
    const colImagemVideo = 5; // Coluna F
    const colNoticiaLink = 6; // Coluna G
    const colNoticiaLink2 = 7; // Coluna H
    const colTituloNoticias = 8; // Coluna I
    const colFonteNoticias = 9; // Coluna J
    const colFonteNoticias2 = 10; // Coluna K
    const colResumoNoticias = 11; // Coluna L
    const colResumoNoticias2 = 12; // Coluna M
  
    // Iterar pelas linhas do Sheets
    for (let i = HEADER_ROW + 1; i < data.length; i++) {
      const logo = data[i][colLogo] || "";
      const tema = data[i][colTema] || "";
      const titulo = data[i][colTitulo] || "";
      const resumo = data[i][colResumo] || "";
      const destaques = data[i][colDestaques] || "";
      const imagemVideo = data[i][colImagemVideo] || "";
      const noticiaLink = data[i][colNoticiaLink] || "";
      const noticiaLink2 = data[i][colNoticiaLink2] || "";
      const titulonoticias = data[i][colTituloNoticias] || "";
      const fontenoticias = data[i][colFonteNoticias] || "";
      const fontenoticias2 = data[i][colFonteNoticias2] || "";
      const resumonoticias = data[i][colResumoNoticias] || "";
      const resumonoticias2 = data[i][colResumoNoticias2] || "";
  
      // Criar um novo slide
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      function adicionarRetangulo() {
    const slide = presentation.getSlides()[presentation.getSlides().length - 1]; // Pega o último slide criado
  
    // Definir as dimensões do retângulo em pontos
    const alturaCm = 14.30; // Altura em cm
    const larguraCm = 0.25; // Largura em cm
  
    // Convertendo de cm para pontos
    const alturaRetangulo = alturaCm * 28.35;  // 1 cm = 28.35 pontos
    const larguraRetangulo = larguraCm * 28.35;  // 1 cm = 28.35 pontos
  
    // Log para verificar as dimensões calculadas
    Logger.log('Altura em pontos: ' + alturaRetangulo);
    Logger.log('Largura em pontos: ' + larguraRetangulo);
  
    // Verificando se as dimensões são válidas
    if (alturaRetangulo <= 0 || larguraRetangulo <= 0) {
      Logger.log("As dimensões do retângulo são inválidas. Verifique as entradas.");
      return; // Não criar o retângulo se as dimensões forem inválidas
    }
  
    // Inserir o retângulo no canto superior esquerdo
    const retangulo = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, larguraRetangulo, alturaRetangulo);
  
    // Definir a cor do retângulo
    const corLogo = '#00FF00';  // Cor da logo ou outra cor
    retangulo.getFill().setSolidFill(corLogo);
  
  
  }
  
      // Inserir o tema no canto superior direito
      const temaWidth = 4.25 * 28.35;  // Convertendo de cm para pt (1 cm = 28.35 pt)
      const temaHeight = 1.00 * 28.35;  // Convertendo de cm para pt (1 cm = 28.35 pt)
  
      // Calcular a posição no canto superior direito (ajustando a margem direita)
      const slideWidth = presentation.getPageWidth();  // Largura da apresentação
      const temaXPosition = slideWidth - temaWidth - 14.17; // Ajustar para distância da borda direita e largura do tema
  
      // Inserir a caixa de texto para o tema
      const temaTextBox = slide.insertTextBox(tema, temaXPosition, 0, temaWidth, temaHeight);
  
      // Obter o texto inserido
      const temaText = temaTextBox.getText();
  
      // Estilizar o tema
      const temaStyle = temaText.getTextStyle();
      temaStyle.setBold(true)
        .setFontFamily("Proxima Nova")
        .setFontSize(10)
        .setForegroundColor('#000000'); // Cor do texto preta (ajustável conforme necessário)
  
      // Ajustar a posição do tema para a centralização manual
      const themeTextLength = temaTextBox.getText().asString().length * 7.5; // Aproximando o comprimento do texto (1 caractere ~ 7.5 pontos)
  
      // Definir o deslocamento para centralizar o texto
      const offset = (temaWidth - themeTextLength) / 2; // Deslocamento para centralização
  
      // Ajustar posição horizontal da caixa de texto
      temaTextBox.setLeft(temaXPosition + offset);
  
      // Definir o fundo cinza claro para o tema
      temaTextBox.getFill().setSolidFill('#D3D3D3');  // Cor cinza claro (#D3D3D3)
  
      // Inserir logo do concorrente mantendo proporção
      if (logo && logo.startsWith("http")) {
        try {
          const response = UrlFetchApp.fetch(logo);
          const blob = response.getBlob();
  
          const image = slide.insertImage(blob)
            .setLeft(14.17) // 0,5 cm da borda esquerda
            .setTop(7.09); // 0,25 cm da borda superior
  
          // Obter dimensões atuais da imagem
          const originalWidth = image.getWidth(); // Largura original em pontos
          const originalHeight = image.getHeight(); // Altura original em pontos
  
          // Definir dimensões máximas (3 cm largura e 1 cm altura)
          const maxWidth = 85.05; // 3 cm em pontos
          const maxHeight = 28.35; // 1 cm em pontos
  
          // Calcular escala mantendo proporção
          const widthScale = maxWidth / originalWidth;
          const heightScale = maxHeight / originalHeight;
          const scale = Math.min(widthScale, heightScale); // Escolher a menor escala para manter proporção
  
          // Aplicar novas dimensões
          image.setWidth(originalWidth * scale);
          image.setHeight(originalHeight * scale);
        } catch (e) {
          Logger.log("Erro ao inserir logo: " + e.message);
        }
      }
  
      // Inserir título abaixo da logo
      const tituloYPosition = 7.09 + 28.35 + 3.96; // 0,25 cm (top logo) + altura máxima da logo (1 cm ou 28.35 pt) + 0,14 cm (3.96 pt)
      const tituloXPosition = 14.17; // Mesma distância da logo em relação à esquerda
  
      const tituloTextBox = slide.insertTextBox(titulo, tituloXPosition, tituloYPosition, 400, 50);
      const tituloText = tituloTextBox.getText();
  
      // Configurar estilo do título
      const tituloStyle = tituloText.getTextStyle();
      tituloStyle.setBold(true)
        .setFontFamily("Proxima Nova")
        .setFontSize(11);
  
      // Inserir resumo abaixo do título
      const resumoYPosition = 56.7; // 2.00 cm em pontos
      const resumoXPosition = 14.17; // Distância da borda esquerda
      const resumoWidth = 361.46; // 12.75 cm em pontos
  
      const resumoTextBox = slide.insertTextBox(resumo, resumoXPosition, resumoYPosition, resumoWidth, 100);
      const resumoText = resumoTextBox.getText();
  
      // Configurar estilo do resumo
      const resumoStyle = resumoText.getTextStyle();
      resumoStyle.setFontFamily("Proxima Nova")
        .setFontSize(10); // Tamanho da fonte
  
      // Inserir os destaques com a formatação corrigida
      const destaqueYPosition = 233.2875; // 8.25 cm em pt (8.25 * 28.35 pt)
      const destaqueXPosition = 14.17; // Mesma distância da borda esquerda que o título
  
      const destaqueTextBox = slide.insertTextBox("Destaques:", destaqueXPosition, destaqueYPosition, resumoWidth, 200);
      const destaqueText = destaqueTextBox.getText();
  
      // Adicionar os pontos aos destaques
      const pontos = destaques.split("\n");
      pontos.forEach((ponto) => {
        destaqueText.appendParagraph(`• ${ponto}`);
      });
  
      // Configurar o estilo dos destaques
      const destaqueStyle = destaqueText.getTextStyle();
      destaqueStyle.setFontFamily("Proxima Nova")
        .setFontSize(10); // Tamanho da fonte 10 para os destaques
  
      // Inserir imagem ou vídeo mantendo proporção
  if (imagemVideo.startsWith("http")) {
    try {
      const response = UrlFetchApp.fetch(imagemVideo);
      const blob = response.getBlob();
  
      // Inserir a imagem no slide
      const image = slide.insertImage(blob)
        .setTop(63.79); // 2.25 cm em pontos (2.25 * 28.35)
  
      // Obter dimensões atuais da imagem
      const originalWidth = image.getWidth();
      const originalHeight = image.getHeight();
  
      // Definir a altura máxima da imagem em pontos (5 cm = 5 * 28.35 pt)
      const maxHeight = 5 * 28.35; // 5 cm em pontos
  
      // Calcular a escala mantendo a proporção
      const heightScale = maxHeight / originalHeight;
      const scale = heightScale; // Escolher a escala para altura
  
      // Aplicar novas dimensões mantendo a proporção
      image.setHeight(originalHeight * scale);
      image.setWidth(originalWidth * scale);
  
      // Definir o intervalo entre 13.50 cm e 24.75 cm
      const leftBoundary = 13.50 * 28.35; // 13.50 cm em pontos
      const rightBoundary = 24.75 * 28.35; // 24.75 cm em pontos
  
      // Calcular a posição central
      const centerPosition = (leftBoundary + rightBoundary - image.getWidth()) / 2;
  
      // Ajustar a posição horizontal (left) para centralizar a imagem
      image.setLeft(centerPosition);
  
    } catch (e) {
      Logger.log("Erro ao inserir imagem ou vídeo: " + e.message);
    }
  }
      const tituloNoticias = data[i][colTituloNoticias] || ""; // Aqui assumimos que "colTituloNoticias" é o índice da coluna correspondente ao título da notícia na planilha
  
      // Definir a posição do título (mesma altura dos destaques, a 13.50 cm da esquerda)
      const tituloNoticiasYPosition = destaqueYPosition; // Usando a mesma posição Y dos destaques
      const tituloNoticiasXPosition = 13.50 * 28.35; // 13.50 cm em pontos (13.50 * 28.35)
  
      // Inserir caixa de texto para o título
      const tituloNoticiasTextBox = slide.insertTextBox(tituloNoticias, tituloNoticiasXPosition, tituloNoticiasYPosition, 400, 50);
      const tituloNoticiasText = tituloNoticiasTextBox.getText();
  
      // Configurar estilo do título da notícia
      const tituloNoticiasStyle = tituloNoticiasText.getTextStyle();
      tituloNoticiasStyle.setBold(true)
        .setFontFamily("Proxima Nova")
        .setFontSize(10); // Tamanho 10, negrito, Proxima Nova
        
      // Obter a Fonte da Notícia da planilha
      const fonteNoticias = data[i][colFonteNoticias] || ""; // A coluna 9 (colFonteNoticias) armazena o campo Fonte da Notícia
  
      // Definir a posição da Fonte da Notícia (0.75 cm abaixo do Título da Notícia)
      const fonteNoticiasYPosition = tituloNoticiasYPosition + 0.75 * 28.35; // 0.40 cm abaixo do título (0.20 + 0.20)
      const fonteNoticiasXPosition = 13.50 * 28.35; // 13.50 cm da borda esquerda (mesmo que o título)
  
      // Inserir caixa de texto para a fonte da notícia
      const fonteNoticiasTextBox = slide.insertTextBox(fonteNoticias, fonteNoticiasXPosition, fonteNoticiasYPosition, 400, 50);
      const fonteNoticiasText = fonteNoticiasTextBox.getText();
  
      // Configurar estilo da Fonte da Notícia
      const fonteNoticiasStyle = fonteNoticiasText.getTextStyle();
      fonteNoticiasStyle.setFontFamily("Proxima Nova")
        .setFontSize(6) // Tamanho 6, conforme solicitado
        .setBold(true); // Negrito
  
      // Ajustar a altura da caixa de texto para 0.25 cm
      fonteNoticiasTextBox.setHeight(0.25 * 28.35); // 0.25 cm de altura em pontos
  
    // Caso o nome 'noticiaLink' já tenha sido utilizado, altere para outro nome, como 'noticiaLinkTexto'
      const noticiaLinkTexto = data[i][colNoticiaLink] || ""; // A coluna 6 (colNoticiaLink) armazena o Link da Notícia
  
      // Definir a posição do Link da Notícia (0.25 cm abaixo da Fonte da Notícia)
      const noticiaLinkYPosition = fonteNoticiasYPosition + 0.25 * 28.35; // 0.25 cm abaixo da Fonte da Notícia
      const noticiaLinkXPosition = 13.50 * 28.35; // 13.50 cm da borda esquerda (mesmo que o título)
  
      // Inserir caixa de texto para o Link da Notícia
      const noticiaLinkTextBox = slide.insertTextBox(noticiaLinkTexto, noticiaLinkXPosition, noticiaLinkYPosition, 11.25 * 28.35, 1.00 * 28.35);
      const noticiaLinkText = noticiaLinkTextBox.getText();
  
      // Configurar estilo do Link da Notícia
      const noticiaLinkStyle = noticiaLinkText.getTextStyle();
      noticiaLinkStyle.setFontFamily("Proxima Nova")
        .setFontSize(10) // Tamanho 10, conforme solicitado
        .setBold(true); // Negrito
  
      // Ajustar a altura da caixa de texto para 1.00 cm
      noticiaLinkTextBox.setHeight(1.00 * 28.35); // 1 cm de altura em pontos
  
      // Aplique o link clicável ao texto inserido na caixa de texto
      noticiaLinkText.getTextStyle().setLinkUrl(noticiaLinkTexto); // Aplica o link em todo o texto inserido
  
      // Resumo da Noticia 
      const resumonoticiasTexto = data[i][colResumoNoticias] || ""; // A coluna 6 (colresumonoticias) armazena o Link da Notícia
  
      // Definir a posição do Resumo da noticia (0.25 cm abaixo da Fonte da Notícia)
      const resumonoticiasYPosition = noticiaLinkYPosition + 0.50 * 28.35; // 0.25 cm abaixo da Fonte da Notícia
      const resumonoticiasXPosition = 13.50 * 28.35; // 13.50 cm da borda esquerda (mesmo que o título)
  
      // Inserir caixa de texto para o Resumo da noticia
      const resumonoticiasTextBox = slide.insertTextBox(resumonoticiasTexto, resumonoticiasXPosition, resumonoticiasYPosition,11.25 * 28.35, 1.00 * 28.35);
      const resumonoticiasText = resumonoticiasTextBox.getText();
  
      // Configurar estilo do Resumo da noticia
      const resumonoticiasStyle = resumonoticiasText.getTextStyle();
      resumonoticiasStyle.setFontFamily("Proxima Nova")
        .setFontSize(8) // Tamanho 10, conforme solicitado
        .setBold(false); // sem negrito
  
      // Ajustar a altura da caixa de texto para 1.00 cm
      resumonoticiasTextBox.setHeight(1.00 * 28.35); // 1 cm de altura em pontos
  
      // Obter a Fonte da Notícia da planilha 2
      const fonteNoticias2 = data[i][colFonteNoticias2] || ""; // A coluna 9 (colFonteNoticias2) armazena o campo Fonte da Notícia
  
      // Definir a posição da Fonte da Notícia 2 (0.75 cm abaixo do Título da Notícia)2
      const fonteNoticias2YPosition = resumonoticiasYPosition + 0.75 * 28.35; // 0.40 cm abaixo do título (0.20 + 0.20)
      const fonteNoticias2XPosition = 13.50 * 28.35; // 13.50 cm da borda esquerda (mesmo que o título)
  
      // Inserir caixa de texto para a fonte da notícia2
      const fonteNoticias2TextBox = slide.insertTextBox(fonteNoticias2, fonteNoticias2XPosition, fonteNoticias2YPosition, 400, 50);
      const fonteNoticias2Text = fonteNoticias2TextBox.getText();
  
      // Configurar estilo da Fonte da Notícia2
      const fonteNoticias2Style = fonteNoticias2Text.getTextStyle();
      fonteNoticias2Style.setFontFamily("Proxima Nova")
        .setFontSize(6) // Tamanho 6, conforme solicitado
        .setBold(true); // Negrito
  
      // Ajustar a altura da caixa de texto para 0.25 cm
      fonteNoticias2TextBox.setHeight(0.25 * 28.35); // 0.25 cm de altura em pontos
  
      // Obter o Link da Notícia 2 da planilha
      const noticiaLink2Texto = data[i][colNoticiaLink2] || ""; // A coluna 7 (colNoticiaLink2) armazena o Link da Notícia 2
  
      // Definir a posição do Link da Notícia 2 (0.25 cm abaixo do Título 2 da Notícia)
      const noticiaLink2YPosition = fonteNoticias2YPosition + 0.25 * 28.35; // 0.25 cm abaixo do Título 2 da Notícia
      const noticiaLink2XPosition = 13.50 * 28.35; // 13.50 cm da borda esquerda (mesmo que o título)
  
      // Inserir caixa de texto para o Link da Notícia 2
      const noticiaLink2TextBox = slide.insertTextBox(noticiaLink2Texto, noticiaLink2XPosition, noticiaLink2YPosition, 11.25 * 28.35, 1.00 * 28.35);
  
      // Configurar estilo do Link da Notícia 2
      const noticiaLink2Text = noticiaLink2TextBox.getText();
      const noticiaLink2Style = noticiaLink2Text.getTextStyle();
      noticiaLink2Style.setFontFamily("Proxima Nova")
        .setFontSize(10) // Tamanho de fonte 10
        .setBold(true); // Negrito
  
      // Aplicar o link clicável ao texto inserido na caixa de texto
      noticiaLink2Text.getTextStyle().setLinkUrl(noticiaLink2Texto); // Aplica o link em todo o texto inserido
  
      // Resumo da Notícia 2
      const resumonoticias2Texto = data[i][colResumoNoticias2] || ""; // A coluna 6 (colResumoNoticias2) armazena o Resumo da Notícia 2
  
      // Definir a posição do Resumo da notícia 2 (0.50 cm abaixo do Link da Notícia 2)
      const resumonoticias2TextoYPosition = noticiaLink2YPosition + 1.00 * 28.35; // 0.50 cm abaixo do Link da Notícia 2
      const resumonoticias2TextoXPosition = 13.50 * 28.35; // 13.50 cm da borda esquerda (mesmo que o título)
  
      // Inserir caixa de texto para o Resumo da notícia 2
      const resumonoticias2TextBox = slide.insertTextBox(resumonoticias2Texto, resumonoticias2TextoXPosition, resumonoticias2TextoYPosition, 11.25 * 28.35, 1.00 * 28.35);
  
      // Configurar estilo do Resumo da notícia 2
      const resumonoticias2TextoText = resumonoticias2TextBox.getText();
      const resumonoticias2Style = resumonoticias2TextoText.getTextStyle();
      resumonoticias2Style.setFontFamily("Proxima Nova")
        .setFontSize(8) // Tamanho 8, conforme solicitado
        .setBold(false); // sem negrito
  
      // Ajustar a altura da caixa de texto para 1.00 cm
      resumonoticias2TextBox.setHeight(1.00 * 28.35); // 1 cm de altura em pontos
  
      // Chama a função para adicionar o retângulo no slide
      adicionarRetangulo();
  
      Logger.log(`Slide ${i} gerado com sucesso!`);
    }
  
    Logger.log("Slides gerados com sucesso!");
  }
  