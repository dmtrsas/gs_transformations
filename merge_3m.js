function merge_3m() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var badQualitySheet = ss.getSheetByName("Bad quality");
  var exportSheet = ss.getSheetByName("Выгрузка");

  if (!badQualitySheet || !exportSheet) {
    Logger.log("Один из листов не найден.");
    return;
  }

  var exportData = exportSheet.getDataRange().getValues();
  var badQualityData = badQualitySheet.getDataRange().getValues();

  if (exportData.length === 0 || badQualityData.length === 0) return;

  // Определяем индексы
  var exportHeader = exportData[0];
  var idIndexExport = exportHeader.indexOf("Split 1");
  var amountIndexExport = exportHeader.indexOf("14D Deposit OAS");

  var badHeader = badQualityData[0];
  var idIndexBad = badHeader.indexOf("ID");
  var outputIndex = badHeader.indexOf("14D Deposit/TC - last 3 months");
  var statusIndex = badHeader.indexOf("Fin status");

  if ([idIndexExport, amountIndexExport, idIndexBad, outputIndex, statusIndex].includes(-1)) {
    Logger.log("Не найдены нужные столбцы.");
    return;
  }

  // Группируем значения по ID
  var dataMap = {};
  for (var i = 1; i < exportData.length; i++) {
    var id = exportData[i][idIndexExport];
    var amountRaw = exportData[i][amountIndexExport];

    if (!id || !amountRaw) continue;

    var val = parseFloat(amountRaw.toString().replace(",", ".").replace("%", ""));
    if (isNaN(val)) continue;

    if (!dataMap[id]) dataMap[id] = [];
    dataMap[id].push(val);
  }

  // Обновляем колонку в Bad quality
  for (var r = 1; r < badQualityData.length; r++) {

    if (badQualityData[r][statusIndex] == 'Paid') continue;
    var id = badQualityData[r][idIndexBad];
    if (!id || !dataMap[id]) continue;

    var values = dataMap[id];
    var sum = values.reduce((a, b) => a + b, 0);
    var avg = String(Math.round((sum / values.length).toFixed(2) * 100)) + "%";

    badQualitySheet.getRange(r + 1, outputIndex + 1).setValue(avg);
  }

  Logger.log("Готово: колонка '14D Deposit/TC - last 3 months' обновлена.");
}
