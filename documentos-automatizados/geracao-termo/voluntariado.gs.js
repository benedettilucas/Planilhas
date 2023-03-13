function voluntariado() {
  
  contarArquivos();
  //ATIVA AS APLICAÇÕES NECESSÁRIAS
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); //acessa a planilha ativa
  let doc = DocumentApp;
  let sheet = spreadsheet.getSheetByName("Termo de Voluntariado");//Acessando a aba Termo Voluntariado

  //DADOS VOLUNTARIO(A)
  let dadosPessoais = sheet.getRange("F3:F11").getValues();
  let dadosEndereco = sheet.getRange("F15:F19").getValues();
  Logger.log(dadosPessoais)
  //DADOS REPRESENTANTE
  let dadosRepresentante = sheet.getRange("C8:C11").getValues();  

  //DADOS PARA DOCUMENTO
  let numTermo = sheet.getRange("C2").getValues();
  numTermo = numeroDoTermo(numTermo);
  let data = [Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd"),Utilities.formatDate(new Date(), "America/Sao_Paulo", "MM"),Utilities.formatDate(new Date(), "America/Sao_Paulo", "YYYY")]
  let mes = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  let dia = ['1º','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31'];
  let dataHoje = dia[data[0]-1]+" de "+mes[data[1]-1]+" de "+data[2]
  
  //DADOS DE MÁSCARA PARA SUBSTITUIÇÃO
  let mascaraVolPes = ["{matricula}","{voluntario}","{curso}","{nacionalidade}","{estado civil}","​{profissao}","{CPF}","{RG}","{ps}"];
  let mascaraVolEnd = ["{endereco}","{bairro}","{cidade}","{UF}","{CEP}"];
  let mascaraRep = ["{cargo-rep}","{nome-rep}","{cpf-rep}","{rg-rep}"];
  let mascaraDoc = ["{termo-n}","{data-firmamento}","{ano}"];

  //CONFIGURAÇÕES DE DOCUMENTOS
  let pasta = DriveApp.getFolderById("url-pasta-destino");
  let arquivo_modelo = DriveApp.getFileById("url-arquivo-modelo");
  let nome_arquivo = "[Sem assinatura] - Termo de Voluntariado " + numTermo + "/" + data[2] + " - " + dadosPessoais[1];
  let nome_arquivo2 = "[Sem assinatura] - Termo de Voluntariado " + numTermo + "/" + data[2] + " - " + dadosPessoais[1]+ ".pdf";
  let novo_termo = arquivo_modelo.makeCopy(nome_arquivo, pasta);
  let termoVoluntariado = doc.openById(novo_termo.getId());
  let body = termoVoluntariado.getBody();
  

  for(i=0; i<mascaraVolPes.length; i++){
    body.replaceText(mascaraVolPes[i], dadosPessoais[i]);
  }
  
  for(i=0; i<mascaraVolEnd.length; i++){
    body.replaceText(mascaraVolEnd[i], dadosEndereco[i]);
  }

  for(i=0; i<mascaraRep.length; i++){
    body.replaceText(mascaraRep[i], dadosRepresentante[i]);
  }
  
  body.replaceText(mascaraDoc[0],numTermo);
  body.replaceText(mascaraDoc[1],dataHoje); //data manual ex: body.replaceText(mascaraDoc[1],"24/04/2018");
  body.replaceText(mascaraDoc[2],data[2]);

  termoVoluntariado.saveAndClose();

  //Gera o PDF
  const termoVoluntariadoPDF = termoVoluntariado.getAs(MimeType.PDF);
  pasta.createFile(termoVoluntariadoPDF);

  //Apaga o arquivo cópia
  pasta.removeFile(novo_termo)

  //Limpa os formulários
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Termo de Voluntariado").getRange("C8").setValue("Selecionar");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Termo de Voluntariado").getRange("C4").setValue("");

  //Logger.log(nome_arquivo)
  contarArquivos();

  let files = DriveApp.getFilesByName(nome_arquivo2);
  while (files.hasNext()) {
    var file = files.next();

  }

  let htmlString = '<!DOCTYPE html>'
    + '<html> <head> <base target="_top"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>'
    + '<div class="spinner-border" role="status">'
    + '<span class="visually-hidden"></span>'
    + '</div>'
    + '<script>'
    + 'window.open(##URL## , ##TYPE##, width=0, height=0); google.script.host.close();'
    + '</script>'
    + '</html>';

    // Change the parameters inside the window.open method 
  htmlString = htmlString.replace("##URL##", "'" + file.getUrl() + "'");
  htmlString = htmlString.replace("##TYPE##", "'_blank'");

  // Create the output window 
  let html = HtmlService.createHtmlOutput(htmlString)
      .setWidth(200)
      .setHeight(40);
  
  // Show the Window
  SpreadsheetApp.getUi() 
      .showModalDialog(html,"Carregando...");
  //window.open(pdfFile, "_blank", "width=600,height=800");
  //return file.getUrl()
  limparDados();
}

function numeroDoTermo(numTermo){

    if (numTermo < 10) {
    numTermo = '000'+numTermo;
  }

  else if (numTermo < 99 && numTermo >= 10){
    numTermo = '00'+numTermo;
  }

  else{
    numTermo = '0'+numTermo;
  }
  return numTermo
}
function limparDados(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Termo de Voluntariado").getRange("C8").setValue("Selecionar");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Termo de Voluntariado").getRange("C4").setValue("");
}
function contarArquivos() {
  
  let pasta = DriveApp.getFolderById("url-pasta-destino");
  let arquivos = pasta.getFiles();
  let numeroTermo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Termo de Voluntariado").getRange("C2");
  let contador = [];
  

  while (arquivos.hasNext()) {contador.push(arquivos.next());}
  
  numeroTermo.setValue(contador.length+1);

}

// Generate a custom HTML to open the PDF
function showPdfFile(){
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  let paginaDeVendas = HtmlService.createTemplateFromFile('pag.html').evaluate();
  paginaDeVendas.setWidth(400).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(paginaDeVendas,"RH | Sistema de Gestão");
}