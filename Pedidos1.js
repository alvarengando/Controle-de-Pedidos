/* ********************************************  Inicio Novo Pedido ******************************************* */

//Formular Pedido

function formularPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');

  pedidos1.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
  pedidos1.getRange('D6').setFormula('=IF(K15="";"";NOW())');
  pedidos1.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  pedidos1.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  pedidos1.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  pedidos1.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  pedidos1.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  pedidos1.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  pedidos1.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  pedidos1.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');
  pedidos1.getRange('M7').setFormula('=IF(D4="";"";"Pendente")');
  pedidos1.getRange('J7').setFormula('=IF(C16="";H13;C16)');
  pedidos1.getRange('J9').setFormula('=IF(J7="";"";"Telegás")');

};

//Modo Salvar Pedido
function modoNovoPedido1() {
  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');

  pedidos1.getRange('AL3').setValue(1);
  pedidos1.getRange('D1').setValue("Novo");
  pedidos1.getRange('AN3').setFormula('=IF(AND(C13="";C16="");"";IF(C16<>"";QUERY(\'Pedidos Dados\'!A:K;"SELECT * WHERE "&C16&" = I ORDER BY A DESC LIMIT 1");QUERY(\'Pedidos Dados\'!A:K;"SELECT * WHERE \'"&C13&"\' = D ORDER BY A DESC LIMIT 1")))');

  pedidos1.getRangeList(['C13', 'C16', 'J11', 'K12:K16', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
  pedidos1.getRange('D4').setBackground('#134f5c').setFontColor('#ffffff').clearDataValidations().setFormula('=IF(G7="";"";MAX(\'Pedidos Dados\'!A2:A)+1)');
  pedidos1.getRange('D5').setFormula('=IF(AND(C16="";G7="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');


  formularPedido1();

  pedidos1.getRange('C16').activate();

};

/* ************************ Salvar Pedido ******************** */

function SalvarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');

  if (pedidos1.getRange('AK3').getValue() > 0) {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }

  else {

    // Sal  var na Página Pedidos Dados

    var values = [[pedidos1.getRange('D4').getValue(),    // ID Pedido
    pedidos1.getRange('D6').getValue(),    // Data Pedido
    pedidos1.getRange('D5').getValue(),    // ID Cliente
    pedidos1.getRange('G7').getValue(),    // Cliente
    pedidos1.getRange('G9').getValue(),    // Logradouro
    pedidos1.getRange('H10').getValue(),   // Complemento
    pedidos1.getRange('H11').getValue(),   // Município
    pedidos1.getRange('H12').getValue(),   // Bairro
    pedidos1.getRange('H13').getValue(),   // Telefone Cadastrado
    pedidos1.getRange('G15').getValue(),   // Referência
    pedidos1.getRange('J7').getValue(),    // Telefone Utilizado
    pedidos1.getRange('J11').getValue(),   // Motorista
    pedidos1.getRange('K12').getValue(),   // Produto
    pedidos1.getRange('K13').getValue(),   // Quantidade
    pedidos1.getRange('K14').getValue(),   // Preço
    pedidos1.getRange('K15').getValue(),   // Total
    pedidos1.getRange('M7').getValue(),    // Status
    pedidos1.getRange('J9').getValue(),    // Canal de Venda
      "",                                  // Justificativa de cancelamento  
    pedidos1.getRange('K16').getValue()    // Forma de Pagamento
    ]];

    PedidosDados.getRange(PedidosDados.getLastRow() + 1, 1, 1, 20).setValues(values);
    pedidos1.getRangeList(['C13', 'C16', 'J11', 'K12:K16']).clear({ contentsOnly: true, skipFilteredRows: true });

    formularPedido();
    Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK);
    pedidos1.getRange('C16').activate();

  }

};

//Formular editar Pedido
function formularEditarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');

  pedidos1.getRange('D5').setFormula('=IF(AND(C13="";C16="");"";IF(AP4="";COUNTA(\'Pedidos Dados\'!C2:C)+1;AP4))');
  pedidos1.getRange('D6').setFormula('=IF(D4="";"";AO4)');
  pedidos1.getRange('G7').setFormula('=IF(AND(C13="";C16="");"";AQ4)');
  pedidos1.getRange('G9').setFormula('=IF(AND(C13="";C16="");"";AR4)');
  pedidos1.getRange('H10').setFormula('=IF(AND(C13="";C16="");"";AS4)');
  pedidos1.getRange('H11').setFormula('=IF(AND(C13="";C16="");"";AT4)');
  pedidos1.getRange('H12').setFormula('=IF(AND(C13="";C16="");"";AU4)');
  pedidos1.getRange('H13').setFormula('=IF(AND(C13="";C16="");"";AV4)');
  pedidos1.getRange('G15').setFormula('=IF(AND(C13="";C16="");"";AW4)');
  pedidos1.getRange('J7').setFormula('=IF(D4="";"";AX4)');
  pedidos1.getRange('J11').setFormula('=IF(D4="";"";AY4)');
  pedidos1.getRange('J9').setFormula('=IF(D4="";"";BE4)');
  pedidos1.getRange('K12').setFormula('=IF(D4="";"";AZ4)');
  pedidos1.getRange('K13').setFormula('=IF(D4="";"";BA4)');
  pedidos1.getRange('K14').setFormula('=IF(D4="";"";BB4)');
  pedidos1.getRange('K15').setFormula('=IF(K14="";"";K13*K14)');
  pedidos1.getRange('K16').setFormula('=IF(D4="";"";BG4)');
  pedidos1.getRange('M7').setFormula('=IF(D4="";"";BD4)');

};

//Modo Editar Pedido
function modoEditarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');

  pedidos1.getRange('AL3').setValue(2);
  pedidos1.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');

  pedidos1.getRangeList(['D4', 'C13', 'C16']).clear({ contentsOnly: true, skipFilteredRows: true });

  pedidos1.getRange('D1').setValue("Editar");
  //ID Pedido
  pedidos1.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Pedidos 1\'!$BH$4:$BH'), true).build());

  formularEditarPedido1();

  pedidos1.getRange('C16').activate();


};

//Salvar alteração
function editarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  var linhaPedido = pedidos1.getRange('AJ3').getValue(); //linha correspondente em Pedidos dados
  var contVazioPedi = pedidos1.getRange('AK3').getValue();
  var status = pedidos1.getRange('AI3').getValue();

  if (contVazioPedi > 0 || status == 1) {
    if (contVazioPedi > 0) {
      Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
    } else {
      Browser.msgBox("Erro", "Necessário preencher a justificativa do Cancelamento!", Browser.Buttons.OK);
    }


  } else {

    // Salvar na Página Pedidos Dados

    var values = [[pedidos1.getRange('D6').getValue(),    // Data Pedido
    pedidos1.getRange('D5').getValue(),    // ID Cliente
    pedidos1.getRange('G7').getValue(),    // Cliente
    pedidos1.getRange('G9').getValue(),    // Logradouro
    pedidos1.getRange('H10').getValue(),    // Complemento
    pedidos1.getRange('H11').getValue(),   // Município
    pedidos1.getRange('H12').getValue(),   // Bairro
    pedidos1.getRange('H13').getValue(),   // Telefone Cadastrado
    pedidos1.getRange('G15').getValue(),   // Referência
    pedidos1.getRange('J7').getValue(),    // Telefone Utilizado
    pedidos1.getRange('J11').getValue(),   // Motorista
    pedidos1.getRange('K12').getValue(),   // Produto
    pedidos1.getRange('K13').getValue(),   // Quantidade
    pedidos1.getRange('K14').getValue(),   // Preço
    pedidos1.getRange('K15').getValue(),   // Total
    pedidos1.getRange('M7').getValue(),    // Status
    pedidos1.getRange('J9').getValue(),   // Canal de Venda
    pedidos1.getRange('M10').getValue(),   // Justificativa
    pedidos1.getRange('K16').getValue()    // Forma de Pagamento
    ]];

    PedidosDados.getRange(linhaPedido, 2, 1, 19).setValues(values);

    Browser.msgBox("Informativo", "Registro alterado com sucesso!", Browser.Buttons.OK);
    pedidos1.getRangeList(['D4', 'C13', 'C16', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
    formularEditarPedido1();
    pedidos1.getRange('C16').activate();

  }
};

/**    * ********************************* */
//Modo Excluir Pedido
function modoDeletarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');

  pedidos1.getRange('AL3').setValue(3);
  pedidos1.getRange('AN3').setFormula('=IF(AM3="";""; IF(AM3="16";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I");IF(AM3="164";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE "&C16&" = I AND "&D4&" = A");IF(AM3="13";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D");IF(AM3="134";QUERY(\'Pedidos Dados\'!A:T;"SELECT * WHERE \'"&C13&"\' = D AND "&D4&" = A"))))))');

  pedidos1.getRangeList(['D4', 'C13', 'C16']).clear({ contentsOnly: true, skipFilteredRows: true });

  pedidos1.getRange('D1').setValue("Deletar");
  //ID Pedido
  pedidos1.getRange('D4').setBackground('#ffffff').setFontColor('#000000').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Pedidos 1\'!$BH$4:$BH'), true).build());

  formularEditarPedido1();

  pedidos1.getRange('C16').activate();

};

function deletarPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');
  var PedidosDados = spreadsheet.getSheetByName('Pedidos Dados');
  var linhaPedido = pedidos1.getRange('AJ3').getValue(); //linha correspondente em Pedidos dados


  if (pedidos1.getRange('AK3').getValue() > 0) {
    Browser.msgBox("Erro", "Necessário preencher todos os campos essenciais!", Browser.Buttons.OK);
  }

  else {

    PedidosDados.deleteRow(linhaPedido);
    pedidos1.getRangeList(['D4', 'C13', 'C16', 'J11', 'M10']).clear({ contentsOnly: true, skipFilteredRows: true });
    Browser.msgBox("Informativo", "Registro Deletado com sucesso!", Browser.Buttons.OK);

    pedidos1.getRange('C16').activate();

  }

};


//******************    Finalizador   ******************************************************************


function finalizadorPedido1() {

  var spreadsheet = SpreadsheetApp.getActive();
  var pedidos1 = spreadsheet.getSheetByName('Pedidos 1');

  if (pedidos1.getRange('AL3').getValue() == 1) {
    SalvarPedido1();
  }

  else if (pedidos1.getRange('AL3').getValue() == 2) {
    editarPedido1();
  }

  else {
    deletarPedido1();
  }


};