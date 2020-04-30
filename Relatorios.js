function relatoriosPedidosDialog() {
  
    var url = 'https://datastudio.google.com/embed/reporting/9619e806-d091-4b3c-8d35-555ce5b7d744/page/Xn1JB';
    var name = 'Pedidos por Status';
    var url2 = 'https://datastudio.google.com/embed/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB';
    var name2 = 'Cadastro de Clientes';  
    var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a> <br><br/><a href="'+url2+'" target="blank" onclick="google.script.host.close()">'+name2+'</a></body></html>';
    var ui = HtmlService.createHtmlOutput(html)
    SpreadsheetApp.getUi().showModelessDialog(ui,"Relatórios de Pedidos");
  }