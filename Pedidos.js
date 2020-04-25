/* ********************************************  Inicio Novo Pedido ******************************************* */

//Formular Pedido

function formularPedido(){

  var spreadsheet = SpreadsheetApp.getActive(); 

  spreadsheet.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  spreadsheet.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  spreadsheet.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  spreadsheet.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  spreadsheet.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  spreadsheet.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  spreadsheet.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  spreadsheet.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');          
  spreadsheet.getRange('M7').setFormula('=IF(D4="";"";"Pendente")');
  spreadsheet.getRange('J7').setFormula('=IF(C16="";"";C16)');

};

//Modo Salvar Pedido
function modoNovoPedido() {
    var spreadsheet = SpreadsheetApp.getActive(); 
    
    spreadsheet.getRange('AL3').setValue(1);    
    spreadsheet.getRange('D1').setValue("Novo");

    //spreadsheet.getRangeList(['G7','C13','C16','J7','J10','K12:K15','G15','M7']).clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRangeList(['J10','K12:K15']).clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('D4').setFormula('=IF(G7="";"";MAX(\'Pedidos Dados\'!A2:A)+1)');
    spreadsheet.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
    spreadsheet.getRange('D6').setFormula('=IF(K15="";"";NOW())');

    formularPedido();

    spreadsheet.getRange('C16').activate();
    
  };

  /* ************************ Salvar Pedido ******************** */

function SalvarPedido() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Pedidos');
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }
  
  else{
    
    // Sal  var na Página Pedidos Dados
                                           
        var values = [[Form.getRange('D4').getValue(),    // ID Pedido
                       Form.getRange('D6').getValue(),    // Data Pedido
                       Form.getRange('D5').getValue(),    // ID Pedido
                       Form.getRange('G7').getValue(),    // Pedido
                       Form.getRange('G9').getValue(),    // Logradouro
                       Form.getRange('H10').getValue(),    // Complemento
                       Form.getRange('H11').getValue(),   // Município
                       Form.getRange('H12').getValue(),   // Bairro
                       Form.getRange('H13').getValue(),   // Telefone Cadastrado
                       Form.getRange('G15').getValue(),   // Referência
                       Form.getRange('J7').getValue(),   // Telefone Utilizado
                       Form.getRange('J10').getValue(),   // Motorista
                       Form.getRange('K12').getValue(),   // Engregue
                       Form.getRange('K13').getValue(),   // Produto
                       Form.getRange('K14').getValue(),   // Quantidade
                       Form.getRange('K15').getValue(),   // Preço
                       Form.getRange('K16').getValue(),   // Total
                       Form.getRange('M7').getValue(),   // Status
                     ]];
       
          PedidosDados.getRange(PedidosDados.getLastRow()+1,1,1,18).setValues(values);
          spreadsheet.getRangeList(['J10','K12:K15']).clear({contentsOnly: true, skipFilteredRows: true});

          formularPedido();
          Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK);
          spreadsheet.getRange('C16').activate(); 
                   
       } 
         
};


//******************    Finalizador   ******************************************************************


function FinalizadorPedido(){

  var spreadsheet = SpreadsheetApp.getActive();

  if(spreadsheet.getRange('AL3').getValue() == 1)
  {
      SalvarPedido();
  }
                      
  else if(spreadsheet.getRange('AL3').getValue() == 2)
   {
         editarPedido();
   }
                      
   else
   {
    deletarPedido(); 
   } 


};