function certificado() {
  
  contarArquivos3();
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let main_sheet = spreadsheet.getSheetByName('Certificado') //Define a aba principal
  let doc_template = DriveApp.getFileById('url-arquivo-modelo') //Localiza o arquivo modelo
  let folder = DriveApp.getFolderById('url-pasta-destino') //Localiza a pasta de destino
  let temp_file = doc_template.makeCopy(folder) //Cria uma cópia do arquivo modelo na pasta de destino
  let temp_doc_file = DocumentApp.openById(temp_file.getId()) //Abre a cópia do arquivo modelo
  let replace = ['{matricula}','{nome}','{cpf}','{inicio}','{fim}','{dia}','{mes}','{ano}','{carga-horaria}','{curso}','{representante}','{cargo-rep}','{n-certi}','{orientacao}','{siape}'] //Define a lista de termos a serem substituídos
  let data = [Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd"),Utilities.formatDate(new Date(), "America/Sao_Paulo", "MM"),Utilities.formatDate(new Date(), "America/Sao_Paulo", "YYYY")] // Cria uma lista com os valores separados da data (dia, mês e ano)
  let mes = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'] //lista com os nomes dos meses em português, será utilizada na etapa de substituição para converter o valor numérico do mês no seu nome
  let dia = ['1º','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31'];
  let numTermo = main_sheet.getRange("C2").getValues();
  numTermo = numeroDoTermo(numTermo);
  let footer = DocumentApp.openById(temp_file.getId()).getFooter();
  let first_footer = footer.getParent().getChild(5);
  
  //Faz as substituições no arquivo cópia

  //Dados membros
  temp_doc_file.getBody().replaceText(replace[0],main_sheet.getRange(3,6).getValue()); //Adiciona a matricula do membro ao arquivo
  temp_doc_file.getBody().replaceText(replace[1],main_sheet.getRange(4,6).getValue()); //Adiciona o nome do membro ao arquivo
  temp_doc_file.getBody().replaceText(replace[2],main_sheet.getRange(6,6).getValue()); //Adiciona o CPF do membro ao arquivo
  temp_doc_file.getBody().replaceText(replace[3],main_sheet.getRange(7,6).getDisplayValue()); //Adiciona a data de admissão do membro ao arquivo
  temp_doc_file.getBody().replaceText(replace[4],main_sheet.getRange(8,6).getDisplayValue()); //Adiciona a data de desligamento do membro ao arquivo
  temp_doc_file.getBody().replaceText(replace[8],main_sheet.getRange(9,6).getDisplayValue()); //Adiciona a carga horária do membro ao arquivo
  temp_doc_file.getBody().replaceText(replace[9],main_sheet.getRange(5,6).getDisplayValue()); //Adiciona o curso do membro ao arquivo

  //Dados orientacao
  temp_doc_file.getBody().replaceText(replace[13],main_sheet.getRange(11,3).getDisplayValue()); 
  temp_doc_file.getBody().replaceText(replace[14],main_sheet.getRange(12,3).getDisplayValue()); 

  //Dados representante
  temp_doc_file.getBody().replaceText(replace[10],main_sheet.getRange(9,3).getDisplayValue());
  temp_doc_file.getBody().replaceText(replace[11],main_sheet.getRange(8,3).getDisplayValue()); 

  //Dados doc


  first_footer.replaceText(replace[12],numTermo) //Adiciona numero do certificado
  temp_doc_file.getBody().replaceText(replace[5],dia[data[0]-1]) //Adiciona o dia atual ao arquivo
  temp_doc_file.getBody().replaceText(replace[6],mes[data[1]-1]) // Converte o valor numérico do mês em seu respectivo nome e o adiciona ao arquivo
  temp_doc_file.getBody().replaceText(replace[7],data[2]) //Adiciona o ano atualao arquivo

  temp_doc_file.saveAndClose(); //Salva e fecha o arquivo cópia
  
  folder.createFile(temp_file.getAs(MimeType.PDF)).setName("[Sem assinatura] | Certificado " + numTermo + "/" + data[2] + " - " + main_sheet.getRange(4,6).getValue()+".pdf"); //Cria o arquivo PDF
  let nome_arquivo3 = "[Sem assinatura] | Certificado " + numTermo + "/" + data[2] + " - " + main_sheet.getRange(4,6).getValue()+".pdf";
  folder.removeFile(temp_file)//Apaga o arquivo cópia

  //Limpa os formulários
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Certificado").getRange("C8").setValue("Selecionar");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Certificado").getRange("C4").setValue("");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Certificado").getRange("C11").setValue("Selecionar");
    
  contarArquivos3();

  let files = DriveApp.getFilesByName(nome_arquivo3);
  while (files.hasNext()) {
    var file = files.next();
  }

  let htmlString = '<!DOCTYPE html>'
    + '<html> <head> <base target="_top"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>'
    + '<div class="spinner-border" role="status">'
    + '<span class="visually-hidden"></span>'
    + '</div>'
    + '<script>'
    + 'window.open(##URL## , ##TYPE##, width=0, height=0);google.script.host.close();'
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

  limparDados3()
}

function contarArquivos3() {
  
  let pasta = DriveApp.getFolderById("url-pasta-destino");
  let arquivos = pasta.getFiles();
  let numeroTermo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Certificado").getRange("C2");
  let contador = [];
  

  while (arquivos.hasNext()) {contador.push(arquivos.next());}
  
  numeroTermo.setValue(contador.length+1);

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

function limparDados3(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Certificado").getRange("C8").setValue("Selecionar");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Certificado").getRange("C4").setValue("");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Certificado").getRange("C11").setValue("Selecionar");
}

function showPdfFile3(){
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  let paginaDeVendas = HtmlService.createTemplateFromFile('pag3.html').evaluate();
  paginaDeVendas.setWidth(400).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(paginaDeVendas,"RH | Sistema de Gestão");
}