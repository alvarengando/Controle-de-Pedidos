function ggff() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C32').activate()
  .setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
  spreadsheet.getRange('D32').activate()
  .setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
};