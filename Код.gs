let protectionName = `защищаемый-диапазон`; // имя основного защищаемого диапазона
let admin = SpreadsheetApp.getActive().getOwner().getEmail(); // имя админа, владельца файла, на чьё имя будет защищён диапазон
let sheetName = `Лист1`; // имя листа, где будут защищаться диапазоны
let flagColNumber = 16; // номер столбца, в котором проставляется признак - блокировать или нет диапазон
let firstColToProtect = 1; // номер первого столбца, который будет защищён напротив ячейки с признаком
let lastColToProtect = 15; // номер последнего столбца, который будет защищён напротив ячейки с признаком
let numColsToProtect = 15; // число столбцов, которые будут защищены напротив ячейки с признаком


/**
 * Функция защищает диапазон от всех кроме админа.
 * Функция нужна, т.к. стандартный метод защиты добавляет в редакторы всех, у кого есть доступ к таблице
 */
function protectRange(range, targetUser, forAdminOnly) {
  let a1TargetRange = range.getA1Notation();
  Logger.log(`защищаемый диапазон = ${a1TargetRange}`);
  let protection = range.protect();
  let editors = protection.getEditors();
  editors = editors.filter(x => x.getEmail() !== admin);
  if (forAdminOnly) {} else { // в защиту диапазона добавляем и текущего пользователя
    editors = editors.filter(x => x.getEmail() !== targetUser);
  }
  Logger.log(`диапазон защищён от всех\nкроме ${admin} и ${targetUser}`);
  protection.removeEditors(editors);
}


/**
 * 1. Функция запускается, если изменён целевой диапазон на целевом листе
 * 2. Если в 3-м столбце стоит флаг "защитить", защищает соответствующие ячейки диапазон с 1 по 2 столбец
 * 3. Перед тем, как защитить диапазон, проверяет, можно ли слить его с соседними защищёнными диапазонами
 * 4. Если находит соседние, удаляет их и создаёт новый единый защищённый диапазон, куда входят и целевые ячейки.
 */
function onEdit(e) {
  // let sp = SpreadsheetApp.getActive();
  // let sh = sp.getSheetByName(sheetName);
  // let shName = sheetName;
  // let triggerRange = sh.getRange(`C2:C3`);

  let sh = e.range.getSheet();
  let shName = sh.getName();
  let triggerRange = e.range;
  let rgCol1 = triggerRange.getColumn();
  let rgCol2 = triggerRange.getLastColumn();
  let triggerA1Range = triggerRange.getA1Notation();

  let firstCol1 = firstColToProtect;
  let lastCol1 = lastColToProtect;
  let firstRow1 = triggerRange.getRow();
  let lastRow1 = triggerRange.getLastRow();
  let numRows = triggerRange.getNumRows();
  let vl = triggerRange.getValues();
  let numCols = lastCol1 - firstCol1 + 1;
  let rg = sh.getRange(firstRow1, firstCol1, numRows, numCols);
  let a1Rg = rg.getA1Notation(); 
  let firstACol1 = a1Rg.match(/^[A-Z]+/);
  let lastACol1 = a1Rg.match(/:([A-Z]+)/)[1];

  if (shName == sheetName) {
    Logger.log(`Изменён целевой лист ${shName}`);
    if (rgCol1 == flagColNumber && rgCol2 == flagColNumber) {
      Logger.log(`Изменён целевой диапазон ${a1Rg}`);
      if (vl.every(x => x !== ``)) {
        Logger.log(`Поступила команда защитить диапазон. Проверим, нет ли возможности слить диапазон с соседними...`);
        let rangesToRemove = [];
        let rangeToCreate = [a1Rg, firstRow1, lastRow1, firstCol1, lastCol1];
        Logger.log(`Имеем начальный диапазон ${rangeToCreate[0]}`);
        let protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        for (let i=0;i<protections.length;i++) {
          let otherRange = protections[i].getRange();
          let otherA1Range = otherRange.getA1Notation();
          let otherRangeFirstCol = otherRange.getColumn();
          let otherRangeLastCol = otherRange.getLastColumn();
          // убираем из protections диапазоны, столбцы которых не соответствуют
          if (otherRangeFirstCol == firstCol1 && otherRangeLastCol == lastCol1) {
            Logger.log(`Сравниваем с защищённым диапазоном ${otherA1Range}`);
            let firstRow2 = otherRange.getRow();
            let lastRow2 = otherRange.getLastRow();
            let firstDiff = rangeToCreate[1]-lastRow2 > 1; // n1 > k2 на 2 и более
            let secondDiff = rangeToCreate[2]-firstRow2 < -1; // k1 < n2 на 2 и более
            if (!firstDiff && !secondDiff && !rangesToRemove.includes(otherA1Range)) {// диапазон ещё не записан в список на удаление
              Logger.log(`Диапазоны ${rangeToCreate[0]} и ${otherA1Range} смежны. Сливаем...`);
              // находим максимальные значения для границ единого диапазона
              let rangeLimits = [rangeToCreate[1], rangeToCreate[2], firstRow2, lastRow2];
              let topLimit = Math.min(...rangeLimits);
              let bottomLimit = Math.max(...rangeLimits);
              numRows = bottomLimit - topLimit + 1;
              numCols = lastCol1 - firstCol1 + 1;
              // перезаписываем контуры единого диапазона
              rangeToCreate = [`${firstACol1}${topLimit}:${lastACol1}${bottomLimit}`, topLimit, bottomLimit, firstCol1, lastCol1];
              rangesToRemove.push(otherA1Range);
              // удаляем диапазон
              Logger.log(`Удаляем защищённый диапазон ${otherA1Range}, т.к. он войдёт в единый`);
              protections[i].remove();
              SpreadsheetApp.flush();
              // переводим индекс в начало, т.к. после слияния прежние диапазоны могли стать смежными
              i = -1;
            } else {
              Logger.log(`Диапазоны ${rangeToCreate[0]} и ${otherA1Range} далеки либо уже слиты в единый. Ничего не делаем...`);
            }
          } else {Logger.log(`Защищённый диапазон не является целевым. Пропускаем его`)}
        }
        Logger.log(`Защищаем единый диапазон ${rangeToCreate[0]}`);
        rg = sh.getRange(rangeToCreate[0]);
        protectRange(rg, admin, true);
        // выполняем задание на удаление прежних диапазонов и создание единого
      } else {Logger.log(`Поступившая команда ${vl} не требует защищать диапазон. Ничего не делаем...`);}
    } else {Logger.log(`Изменённый диапазон ${triggerA1Range} не является целевым`);}
  } else {Logger.log(`Изменённый лист ${shName} не является целевым`);}
}
