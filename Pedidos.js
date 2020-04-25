/* ********************************************  Inicio Novo Pedido ******************************************* */
//Modo Salvar Pedido
function modoNovoPedido() {
    var spreadsheet = SpreadsheetApp.getActive(); 
    
    spreadsheet.getRange('AL3').setValue(1);    
    spreadsheet.getRange('D1').setValue("Novo");

    spreadsheet.getRangeList(['G7','C13','C16','K12:K15','G15','M7']).clear({contentsOnly: true, skipFilteredRows: true});

    spreadsheet.getRange('D4').setFormula('=IF(G7="";"";MAX(\'Pedidos Dados\'!A2:A)+1)');
    spreadsheet.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
    spreadsheet.getRange('D6').setFormula('=IF(K15="";"";NOW())');

    spreadsheet.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
    spreadsheet.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
    spreadsheet.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
    spreadsheet.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
    spreadsheet.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
    spreadsheet.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
    spreadsheet.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');

    spreadsheet.getRange('J7').setFormula('=IF(C16="";"";C16)');

    spreadsheet.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');

    spreadsheet.getRange('M7').setFormula('=IF(D4="";"";"Pendente")');

    spreadsheet.getRange('C16').activate();
    
  };