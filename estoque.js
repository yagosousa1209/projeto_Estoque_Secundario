function cadastro_Caixote() {
  let ativar_Estoque = ativar_Planilha.getSheetByName("ESTOQUE")

  let valor_D3 = ativar_Estoque.getRange('D3').getValue()
  if (valor_D3 <= 9) {
    var valor_D3_Conf = "CX 0" + valor_D3
  } else {
    var valor_D3_Conf = "CX " + valor_D3
  }
  
  if 
  (
    verificacao_Dados_Celulas("B3") == true &
    verificacao_Dados_Celulas('D3') == true &
    ativar_Estoque.getRange('E3').getValue() !== "" == true
  ) 
  {
    let ultima_Linha = ativar_Planilha.getLastRow() + 1
    ativar_Estoque.getRange('A' + ultima_Linha).activate();
    ativar_Estoque.getRange('A1').copyTo(ativar_Estoque.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    ativar_Estoque.getRange('B' + ultima_Linha).activate();
    ativar_Estoque.getRange('B3').copyTo(ativar_Estoque.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    ativar_Estoque.getRange('G' + ultima_Linha).activate();
    ativar_Estoque.getRange('E3').copyTo(ativar_Estoque.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    ativar_Estoque.getRange('F' + ultima_Linha).setValue(valor_D3_Conf)

    apagar_Dados_Celulas(['B3', 'D3:E3'], 'E3') 
    return
  }
  SpreadsheetApp.getUi().alert("Alguma celula está em branco ou com valores não numéricos")
  
};

function editar_Pesquisar_Caixote() {
  let ativar_Estoque = ativar_Planilha.getSheetByName("ESTOQUE")

  let valor_B5 = ativar_Estoque.getRange('B5').getValue()
  let valor_D5 = ativar_Estoque.getRange('D5').getValue()

  if (valor_D5 <= 9) {
    var valor_D5_conf = "CX 0" + valor_D5
  } else {
    var valor_D5_conf = "CX " + valor_D5
  }

  limpar_Filtros_Estoque()

  if 
  (
    verificacao_Dados_Celulas("B5") == true &
    verificacao_Dados_Celulas('D5') == false
  ) 
  {
    ativar_Estoque.getRange('B10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenNumberEqualTo(valor_B5)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(2, criteria);

  } if 
  (
    verificacao_Dados_Celulas("B5") == false &
    verificacao_Dados_Celulas('D5') == true 
  ) 
  {
    ativar_Estoque.getRange('F10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(valor_D5_conf)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(6, criteria);
 
  } if 
  (
    verificacao_Dados_Celulas("B5") == true &
    verificacao_Dados_Celulas('D5') == true 
  )
  {
    ativar_Estoque.getRange('B10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenNumberEqualTo(valor_B5)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(2, criteria);

    ativar_Estoque.getRange('F10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(valor_D5_conf)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(6, criteria);

  } if (
    verificacao_Dados_Celulas("B5") == false &
    verificacao_Dados_Celulas('D5') == false
  ) {
    SpreadsheetApp.getUi().alert("Por favor digite algum valor ou digite apenas números")
    return
  }
  apagar_Dados_Celulas(["B5", "D5"], "D5")
}

function apagar_Caixote() {
  let ativar_Estoque = ativar_Planilha.getSheetByName("ESTOQUE")

  let valor_B7 = ativar_Estoque.getRange('B7').getValue()
  let valor_D7 = ativar_Estoque.getRange('D7').getValue()

  if (valor_D7 <= 9) {
    var valor_D7_conf = "CX 0" + valor_D7
  } else {
    var valor_D7_conf = "CX " + valor_D7
  }

  if 
  (
    verificacao_Dados_Celulas("B7") == true &
    verificacao_Dados_Celulas('D7') == true 
  )
  {
    ativar_Estoque.getRange('B10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenNumberEqualTo(valor_B7)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(2, criteria);

    ativar_Estoque.getRange('F10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(valor_D7_conf)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(6, criteria);

    ativar_Estoque.getRange('F10').activate();
    let desc_Linha = ativar_Estoque.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate().getRow();
    ativar_Estoque.deleteRow(desc_Linha)

    limpar_Filtros_Estoque()

    return
  }
  SpreadsheetApp.getUi().alert("Alguma celula está em branco ou com valores não numéricos")
}


function apagar_Dados_Celulas(arg_1, arg_2) {
  let apagar_Celulas = ativar_Planilha.getSheetByName("ESTOQUE");
  apagar_Celulas.getRangeList(arg_1).activate();
  apagar_Celulas.setCurrentCell(apagar_Celulas.getRange(arg_2));
  apagar_Celulas.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  apagar_Celulas.getRange('A1').activate();
};


function verificacao_Dados_Celulas (celula_Analizada) {
  let verificacao_Celula = ativar_Planilha.getSheetByName("ESTOQUE")
    
  // Selecione a célula que deseja verificar
  let celula = verificacao_Celula.getRange(celula_Analizada);
    
  // Verifique se a célula contém apenas números e não contém espaços
  if (/^\d+$/.test(celula.getValue())) {
    return true
  } else {
    return false
  }

}

function limpar_Filtros_Estoque() {
  let limpar_Filtros = ativar_Planilha.getSheetByName("ESTOQUE")
  limpar_Filtros.getRange('B10').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  limpar_Filtros.getFilter().setColumnFilterCriteria(2, criteria);

  limpar_Filtros.getRange('F10').activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  limpar_Filtros.getFilter().setColumnFilterCriteria(6, criteria);
  limpar_Filtros.getRange('A11').activate()
};
