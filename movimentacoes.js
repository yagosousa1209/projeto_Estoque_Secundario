
function cadastro_mov () {
  let cadastro_mov = ativar_Planilha.getSheetByName("ENTRADAS_SAIDA");

  if 
  (
    verificacao_Dados_Celulas_mov("B3") == true &
    verificacao_Dados_Celulas_mov('E3') == true &
    cadastro_mov.getRange('G3').getValue() !== "" == true
  ) 
  {
    let ultima_Linha = ativar_Planilha.getLastRow() + 1
    // set data
    cadastro_mov.getRange('A' + ultima_Linha).activate()
    cadastro_mov.getRange('A1').copyTo(cadastro_mov.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    // set cod
    cadastro_mov.getRange('B' + ultima_Linha).activate()
    cadastro_mov.getRange('B3').copyTo(cadastro_mov.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    // set qtd
    cadastro_mov.getRange('D' + ultima_Linha).activate()
    cadastro_mov.getRange('E3').copyTo(cadastro_mov.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    // set cx
    cadastro_mov.getRange('F' + ultima_Linha).activate()
    cadastro_mov.getRange('D3').copyTo(cadastro_mov.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    // set responsavel
    cadastro_mov.getRange('G' + ultima_Linha).activate()
    cadastro_mov.getRange('G3').copyTo(cadastro_mov.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    //set destino
    cadastro_mov.getRange('H' + ultima_Linha).activate()
    cadastro_mov.getRange('F3').copyTo(cadastro_mov.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    apagar_Dados_Celulas_mov(['B3', 'D3:G3'], 'E3') 
    return
  }
  SpreadsheetApp.getUi().alert("Alguma celula está em branco ou com valores não numéricos")
  
}


function editar_Pesquisar_mov() {
  let ativar_Estoque = ativar_Planilha.getSheetByName("ENTRADAS_SAIDA")

  let valor_B5 = ativar_Estoque.getRange('B5').getValue()
  let valor_D5 = ativar_Estoque.getRange('D5').getValue()

  limpar_Filtros_Estoque_mov()

  if 
  (
    verificacao_Dados_Celulas_mov("B5") == true &
    valor_D5 === ""
  ) 
  {
    ativar_Estoque.getRange('B10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenNumberEqualTo(valor_B5)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(2, criteria);

  } if 
  (
    verificacao_Dados_Celulas_mov("B5") == false &
    valor_D5 !== ""
  ) 
  {
    ativar_Estoque.getRange('F10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(valor_D5)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(6, criteria);
 
  } if 
  (
    verificacao_Dados_Celulas_mov("B5") == true &
    valor_D5 !== ""
  )
  {
    ativar_Estoque.getRange('B10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenNumberEqualTo(valor_B5)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(2, criteria);

    ativar_Estoque.getRange('F10').activate();
    var criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(valor_D5)
    .build();
    ativar_Estoque.getFilter().setColumnFilterCriteria(6, criteria);

  } if (
    verificacao_Dados_Celulas_mov("B5") == false &
    valor_D5 === ""
  ) {
    SpreadsheetApp.getUi().alert("Por favor digite algum valor ou digite apenas números.")
    return
  }
  apagar_Dados_Celulas_mov(["B5", "D5"], "B5")
}

function apagar_mov () {
  let ativar_Estoque = ativar_Planilha.getSheetByName("ENTRADAS_SAIDA")
  let linha = ativar_Estoque.getActiveRange().getRow()

  let ui  = SpreadsheetApp.getUi()
  let caixa_pergunta = ui.alert("Deseja realmente apagar a movimentação selecionada?", ui.ButtonSet.YES_NO)

  if (caixa_pergunta == ui.Button.YES) {
    ativar_Estoque.deleteRow(linha)
  } else {
    return
  }
}

function apagar_Dados_Celulas_mov(arg_1, arg_2) {
  let apagar_Celulas = ativar_Planilha.getSheetByName("ENTRADAS_SAIDA");
  apagar_Celulas.getRangeList(arg_1).activate()
  apagar_Celulas.setCurrentCell(apagar_Celulas.getRange(arg_2));
  apagar_Celulas.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  apagar_Celulas.getRange('A1').activate()
};


function verificacao_Dados_Celulas_mov (celula_Analizada) {
  let verificacao_Celula = ativar_Planilha.getSheetByName("ENTRADAS_SAIDA")
    
  // Selecione a célula que deseja verificar
  let celula = verificacao_Celula.getRange(celula_Analizada);
    
  // Verifique se a célula contém apenas números e não contém espaços
  if (/^\d+$/.test(celula.getValue())) {
    return true
  } else {
    return false
  }

}

function limpar_Filtros_Estoque_mov() {
  let limpar_Filtros = ativar_Planilha.getSheetByName("ENTRADAS_SAIDA")
  limpar_Filtros.getRange('B10')
  var criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  limpar_Filtros.getFilter().setColumnFilterCriteria(2, criteria);

  limpar_Filtros.getRange('F10')
  criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  limpar_Filtros.getFilter().setColumnFilterCriteria(6, criteria);
  limpar_Filtros.getRange('A11')
};