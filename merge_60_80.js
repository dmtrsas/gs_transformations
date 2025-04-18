function merge_60_80() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Открываем активную таблицу
  var data1Sheet = ss.getSheetByName("Bad quality 2025_new"); // Получаем лист "Bad quality"
  var data2Sheet = ss.getSheetByName("Выгрузка"); // Получаем лист "Выгрузка"

  if (!data1Sheet || !data2Sheet) {
    Logger.log("Один из листов не найден.");
    return; // Прерываем выполнение, если один из листов не найден
  }

  var data1 = data1Sheet.getDataRange().getDisplayValues(); // Получаем все данные из листа "Bad quality"
  var data2 = data2Sheet.getDataRange().getDisplayValues(); // Получаем все данные из листа "Выгрузка"

  if (data1.length === 0 || data2.length === 0) {
    Logger.log("Один из листов пуст.");
    return; // Прерываем выполнение, если один из листов пуст
  }

  // Получаем индексы нужных столбцов
  var header1 = data1[0];
  var header2 = data2[0];

  var idIndex1 = header1.indexOf("ID");
  var periodIndex = header1.indexOf("Period");
  var amount1Index1 = header1.indexOf("Commission to pay under KPI 60%");
  var amount2Index1 = header1.indexOf("Commission to pay under KPI 80%");
  var statusIndex = header1.indexOf("Fin status");
  var fiatIndex = header1.indexOf("Fiat amount, EUR");
  var rsIndex = header1.indexOf("RS")


  var idIndex2 = header2.indexOf("Split 1");
  var dateIndex = header2.indexOf("Cohort Date");
  var amount1Index2 = header2.indexOf("14D Deposit S");
  var amount2Index2 = header2.indexOf("1M Deposit S");
  var marketspendIndex = header2.indexOf("Marketing Spend")

  // Проверка: все ли нужные столбцы найдены
  if ([idIndex1, periodIndex, amount1Index1, amount2Index1, statusIndex, fiatIndex, rsIndex,
       idIndex2, dateIndex, amount1Index2, amount2Index2, marketspendIndex].includes(-1)) {
    Logger.log("Ошибка: не найдены все нужные столбцы.");
    return;
  }

  // Сопоставление номеров месяцев с названиями
  var monthMap = {
    '01': "January", '02': "February", '03': "March", '04': "April", '05': "May", '06': "June",
    '07': "July", '08': "August", '09': "September", '10': "October", '11': "November", '12': "December"
  };

  var monthOrder = Object.values(monthMap); // Для корректной сортировки месяцев

  // Сборка структуры dataMap: { ID: { Месяц: [14D, 1M] } }
  var dataMap = {};
  for (var i = 1; i < data2.length; i++) {
    var id = data2[i][idIndex2];
    var date = data2[i][dateIndex];
    var month = monthMap[date.substring(3, 5)]; // Извлекаем месяц из Cohort Date

    var raw1 = data2[i][amount1Index2].replace(",", ".").replace(" ", "");
    var raw2 = data2[i][amount2Index2].replace(",", ".").replace(" ", "");
    var raw3 = data2[i][marketspendIndex].replace(",",".").replace(" ", "");

    var amount1 = raw1 ? parseFloat(raw1.replace(/\s/g, "")) / 0.6 : ""; // Преобразуем
    var amount2 = raw2 ? parseFloat(raw2.replace(/\s/g, "")) / 0.8 : "";
    var spend = raw3 ? parseFloat(raw3.replace(/\s/g, "")) : "";

    if (!dataMap[id]) dataMap[id] = {}; // Если ID ещё нет, создаём
    dataMap[id][month] = [amount1, amount2, spend]; // Сохраняем данные по ID и месяцу
  }

  // Справочник из трёхмесячных периодов — для удобства и надёжности
  var generatedCompositePeriods = {}; // Чтобы не считать один и тот же ID+период дважды
  var threeMonthsMap = {
    "January - March": ["January", "February", "March"],
    "February - April": ["February", "March", "April"],
    "March - May": ["March", "April", "May"],
    "April - June": ["April", "May", "June"],
    "May - July": ["May", "June", "July"],
    "June - August": ["June", "July", "August"],
    "July - September": ["July", "August", "September"],
    "August - October": ["August", "September", "October"],
    "September - November": ["September", "October", "November"],
    "October - December": ["October", "November", "December"],
    "November - January": ["November", "December", "January"],
    "December - February": ["December", "January", "February"]
  };

  // Основной цикл по строкам листа "Bad quality"
  for (var i = 1; i < data1.length; i++) {

    if (data1[i][statusIndex] == 'Paid') continue; // Пропускаем уже оплаченные

    var id = data1[i][idIndex1];
    var period = data1[i][periodIndex];
    var fiatAmount = parseFloat(data1[i][fiatIndex].replace(",",".").replace(/\s/g, ""));

    if (!id || !period) continue; // Если нет ID или периода — пропускаем

    var months = period.split("-").map(m => m.trim()).filter(m => m); // Разделяем период по "-"

    if (months.length === 1) {
      // Одномесячный период
      var month = months[0];
      if (dataMap[id] && dataMap[id][month] && data1[i][rsIndex] === "RS") { 
        // Обновляем значения, если есть данные
        data1Sheet.getRange(i + 1, amount1Index1 + 1).setValue(
          fiatAmount - dataMap[id][month][2] + dataMap[id][month][0]); // Fiat amount - Marketing spend + 0.6_Deposit
        data1Sheet.getRange(i + 1, amount2Index1 + 1).setValue(
          fiatAmount - dataMap[id][month][2] + dataMap[id][month][1]); // Fiat amount - Marketing spend + 0.8_Deposit
      } 
      else if (dataMap[id] && dataMap[id][month] && data1[i][rsIndex] != "RS") {
        // Обновляем значения, если есть данные
        data1Sheet.getRange(i + 1, amount1Index1 + 1).setValue(dataMap[id][month][0]);
        data1Sheet.getRange(i + 1, amount2Index1 + 1).setValue(dataMap[id][month][1]);
      }
    } else {
      // Многомесячный период
      var sortedMonths = months.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b)); // Сортируем
      var compositeKey = sortedMonths.join(" - "); // Создаём ключ, например: "June - August"

      if (!generatedCompositePeriods[id]) generatedCompositePeriods[id] = {};
      if (generatedCompositePeriods[id][compositeKey]) continue; // Пропускаем, если уже считан этот период

      var allsortedMonths;
      if (compositeKey in threeMonthsMap) {
        allsortedMonths = threeMonthsMap[compositeKey]; // Если есть в справочнике — берём оттуда
      } else {
        allsortedMonths = sortedMonths; // Иначе используем как есть
      }

      // Eсть ли все месяцы по этому ID
      var allMonthsPresent = allsortedMonths.every(mon => dataMap[id] && dataMap[id][mon]);
      if (!allMonthsPresent) continue; // Пропускаем, если хотя бы одного месяца нет

      // Считаем значения по всем месяцам
      var sum1 = 0, sum2 = 0, sum3 = 0, count1 = 0, count2 = 0, count3 = 0;
      for (var m = 0; m < allsortedMonths.length; m++) {
        var mon = allsortedMonths[m];
        var val1 = parseFloat(dataMap[id][mon][0]);
        var val2 = parseFloat(dataMap[id][mon][1]);
        var val3 = parseFloat(dataMap[id][mon][2]);
        if (!isNaN(val1)) { sum1 += val1; count1++; }
        if (!isNaN(val2)) { sum2 += val2; count2++; }
        if (!isNaN(val3)) { sum3 += val3; count3++; }
      }

      if ((count1 || count2 || count3) && data1[i][rsIndex] === "RS")  {

        sum1 = ((Math.round(sum1 * 100)/100) + (fiatAmount - sum3)).toFixed(2); // SUM of 0.6_Deposits + (Fiat amount - Sum of Marketing spends) 
        sum2 = ((Math.round(sum2 * 100)/100) + (fiatAmount - sum3)).toFixed(2); // SUM of 0.8_Deposits + (Fiat amount - Sum of Marketing spends)

        data1Sheet.getRange(i + 1, amount1Index1 + 1).setValue(sum1);
        data1Sheet.getRange(i + 1, amount2Index1 + 1).setValue(sum2);

        generatedCompositePeriods[id][compositeKey] = [sum1, sum2]; // Сохраняем, чтобы не пересчитывать
      }
      else if ((count1 || count2 || count3) && data1[i][rsIndex] != "RS") {
        sum1 = Math.round(sum1 * 100)/100;
        sum2 = Math.round(sum2 * 100)/100;

        data1Sheet.getRange(i + 1, amount1Index1 + 1).setValue(sum1);
        data1Sheet.getRange(i + 1, amount2Index1 + 1).setValue(sum2);
        generatedCompositePeriods[id][compositeKey] = [sum1, sum2]; // Сохраняем, чтобы не пересчитывать
      }
    }
  }

  // Формируем вспомогательный лист "Лист3" для вывода всех значений
  var sheet3 = ss.getSheetByName("Лист3") || ss.insertSheet("Лист3");
  sheet3.clear(); // Очищаем

  var outputData = [["Split 1", "Cohort Date", "14D Deposit S", "1M Deposit S"]]; // Заголовок

  // Добавляем данные из "Выгрузки"
  for (var id in dataMap) {
    for (var month in dataMap[id]) {
      outputData.push([id, month, dataMap[id][month][0], dataMap[id][month][1]]);
    }
  }

  // Добавляем рассчитанные трёхмесячные значения
  for (var id in generatedCompositePeriods) {
    for (var period in generatedCompositePeriods[id]) {
      var values = generatedCompositePeriods[id][period];
      outputData.push([id, period, values[0], values[1]]);
    }
  }

  // Записываем в "Лист3"
  sheet3.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
  Logger.log("Готово: Bad quality обновлён и Лист3 создан.");
}
