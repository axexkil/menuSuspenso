function menuSuspenso(e) { // função que é executada quando uma célula é editada

  // obtém a planilha ativa
  var sheet = e.source.getActiveSheet();

  // obtém a célula editada
  var editedCell = e.range;

  // define a linha inicial e o número de linhas a serem consideradas
  var startRow = 2;
  var numRows = sheet.getLastRow() - startRow + 1;

  // define a coluna que será usada para verificar que é a coluna editada ("Sistema Afetado") e a coluna para exibir o menu suspenso ("Erros")
  var systemColumn = 5; // coluna "Sistema Afetado"
  var errorColumn = 7; // coluna "Erros"
  
  // verifica se a célula editada está na coluna correta e dentro do intervalo de linhas especificado
  if (editedCell.getColumn() == systemColumn && editedCell.getRow() >= startRow && editedCell.getRow() < startRow + numRows) {

    // obtém o valor da célula editada
    var systemValue = editedCell.getValue();

    // obtém a planilha "DADOS" que contém as opções para o menu suspenso
    var dadosSheet = e.source.getSheetByName("DADOS");

    // filtra as opções na coluna "D" com base no valor da célula editada ("Sistema Afetado")
    var options = dadosSheet.getRange("D2:D").getValues().filter(function(row) {
      return row[0].toString().startsWith(systemValue);
    });

    // cria uma nova regra de validação de dados que exige que a entrada seja uma das opções fornecidas
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(options.map(function(row) {
      return row[0];
    }), true).build();

    // aplica a regra de validação à célula ("Erros") 
    editedCell.offset(0, errorColumn - editedCell.getColumn()).setDataValidation(rule);
  }
}
