function cadastroVenda(){
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  let paginaDeVendas = HtmlService.createTemplateFromFile('vendas-pagina.html').evaluate();
  paginaDeVendas.setWidth(800);
  SpreadsheetApp.getUi().showModalDialog(paginaDeVendas,"Cadastro de venda");}

function cadastroCompra(){
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  let paginaDeProdutos = HtmlService.createTemplateFromFile('compras-pagina.html').evaluate();
  paginaDeProdutos.setWidth(900);
  SpreadsheetApp.getUi().showModalDialog(paginaDeProdutos,"Cadastro de compra");
}

function cadastroProduto(){}
///////////////////////////////////////////////

//INSERSÃO NA PLANILHAS E VALIDAÇÃO DOS FORMULÁRIOS
function insereVenda(venda) { 
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  let menu = spreadsheet.getSheetByName('Vendas');
  let ultimaLinha = menu.getLastRow()+1;
    
  menu.protect().remove();
  
  menu.getRange('A'+ultimaLinha).setValue(Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy - HH:mm:ss"));
  menu.getRange('B'+ultimaLinha).setValue([venda.Venda]);
  menu.getRange('C'+ultimaLinha).setValue([venda.Resp]);
  menu.getRange('D'+ultimaLinha).setValue([venda.Ped]);
  menu.getRange('E'+ultimaLinha).setValue([venda.CPF]);
  menu.getRange('F'+ultimaLinha).setValue([venda.Pgto]);
  menu.getRange('G'+ultimaLinha).setValue([venda.Valor]);

  menu.protect();
  ui.alert('Venda cadastrada com sucesso!');   
}

function validakkao(venda){
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  
  ui.alert("preencha todos os campos!");
}

function desprotege(){
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  let menu = spreadsheet.getSheetByName('Vendas');
  
  menu.protect().remove();
}

function protege(){
  let app = SpreadsheetApp;
  let spreadsheet = app.getActiveSpreadsheet();
  let ui = app.getUi();
  let menu = spreadsheet.getSheetByName('Vendas');
  
  menu.protect();
}
///////////////////////////////////////////////