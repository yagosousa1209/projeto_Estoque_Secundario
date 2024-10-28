const filtros_Mov = ativar_Planilha.getSheetByName("ENTRADAS_SAIDA")

function filtrar_Entradas() {
  filtros_Mov.getRange('D4').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenNumberGreaterThanOrEqualTo(0)
  .build();
  filtros_Mov.getFilter().setColumnFilterCriteria(4, criteria);
};

function limpar_Filtros() {
  filtros_Mov.getRange('D4').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  filtros_Mov.getFilter().setColumnFilterCriteria(4, criteria);
};

function filtrar_Saidas() { 
  filtros_Mov.getRange('D4').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenNumberLessThan(0)
  .build();
  filtros_Mov.getFilter().setColumnFilterCriteria(4, criteria);
};
