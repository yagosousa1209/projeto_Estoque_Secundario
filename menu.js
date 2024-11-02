function menu_Estoque() {
  ativar_Planilha.getSheetByName('ESTOQUE').activate()
}

function menu_Entradas_Saidas () {
  ativar_Planilha.getSheetByName('ENTRADAS_SAIDA').activate()
}

function menu_BD_Itens () {
  ativar_Planilha.getSheetByName('BD ITENS').activate()
}

function menu_Lista_Suspensa () {
  ativar_Planilha.getSheetByName('LISTA SUSPENSA').activate()
}

function voltar_Topo () {
  ativar_Planilha.getSheetByName("ESTOQUE").getRange('A11').activate()
}

function voltar_Topo_mov () {
  ativar_Planilha.getSheetByName("ENTRADAS_SAIDA").getRange('A11').activate()
}