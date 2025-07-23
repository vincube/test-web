/**
 * @OnlyCurrentDoc
 */
// Updated: 2025-07-24
// ... (константы и начало кода без изменений) ...
const SPREADSHEET_ID = "18Hx_V3lPBQYp4xcpopQPjQ5dR7cCtaWpESb1EM92gqI"; // ID ВАШЕЙ GOOGLE ТАБЛИЦЫ
const SHEET_NAME = "АвтоматическийОтчет";
const VENDISTA_API_BASE_URL = "https://api.vendista.ru:99";
const VENDISTA_API_TOKEN = "fe38f8470367451f81228617"; // ВАШ API ТОКЕН VENDISTA (используется для удаленного отчета)

const MACHINES_PER_TRIGGER_RUN = 5;
const DELAY_AFTER_ZAPROS_MS = 5000;
const DELAY_BETWEEN_MACHINES_MS = 500;
const DELAY_BETWEEN_TRIGGERS_MINUTES = 1;
const NUM_SLOTS_CONST = 12;
const NUM_COLUMNS_MAIN_TABLE = 8; 


// --- Функции для веб-интерфейса ---
// doGet, include, updateData, fetchPacketlog, parseCustomPacketToFields, loadMachineList, loadProductMatrix, loadMachineIngredients, sendCommand, formatVendistaDateTime, fetchSalesList - БЕЗ ИЗМЕНЕНИЙ
function showRemote(e) {
  return HtmlService.createTemplateFromFile("Index").evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Управление товарами и журнал пакетов");
}


function updateData(params) {
  if (!params || !params.token || !params.terminalId) {
    Logger.log("updateData: Отсутствует токен или terminalId в params: %s", JSON.stringify(params));
    return { found: false, error: "Токен и terminalId обязательны (серверная проверка).", prices: [], errors: [], allowed_errors: null, uart_speed: null, timestamp: null, source: 'api' };
  }
  try {
    var jsonData = fetchPacketlog(params);
    if (jsonData.error) {
        return { found: false, error: jsonData.error, prices: [], errors: [], allowed_errors: null, uart_speed: null, timestamp: null, source: 'api' };
    }
    var items = jsonData.items || [];
    var result = { found: false, prices: [], errors: [], allowed_errors: null, uart_speed: null, timestamp: null, source: 'api' };
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var desc = item.description || "";
      if (desc.startsWith("CustomPacketToExternalServer:")) {
        var parsedPacket = parseCustomPacketToFields(desc);
        if (parsedPacket) {
          result.prices = parsedPacket.prices; result.errors = parsedPacket.errors;
          result.allowed_errors = parsedPacket.allowed_errors; result.uart_speed = parsedPacket.uart_speed;
          result.timestamp = item.time; result.found = true;
        } else { result.error = "Не удалось разобрать данные пакета."; }
        break;
      }
    }
    if (!result.found && items.length > 0 && !result.error) result.error = "Нет данных журнала пакетов для терминала/фильтра.";
    else if (items.length === 0 && !result.error) result.error = "Нет данных журнала пакетов для терминала/фильтра.";
    return result;
  } catch (error) {
    Logger.log("updateData Exception: %s", error.toString());
    return { found: false, error: `Ошибка обновления данных: ${error.message}`, prices: [], errors: [], allowed_errors: null, uart_speed: null, timestamp: null, source: 'api' };
  }
}

function fetchPacketlog(params) {
  var url = `${VENDISTA_API_BASE_URL}/packetlog/${encodeURIComponent(params.terminalId)}?FilterText=CustomPacketToExternalServer&PageNumber=1&ItemsOnPage=10&OrderDesc=true&token=${encodeURIComponent(params.token)}`;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseCode = resp.getResponseCode(); var contentText = resp.getContentText();
    if (responseCode !== 200) {
      Logger.log("fetchPacketlog API Error: Code %s, Response: %s", responseCode, contentText);
      return { error: `Ошибка API (packetlog): ${responseCode}. ${contentText}`, items: [] };
    }
    return JSON.parse(contentText || "{}");
  } catch (e) {
    Logger.log("fetchPacketlog Exception: %s", e.toString());
    return { error: `Исключение (packetlog): ${e.message}`, items: [] };
  }
}

function parseCustomPacketToFields(str) {
  var rawValues = str.replace("CustomPacketToExternalServer:", "").trim().split(/\s+/);
  var prices = []; var errors = []; var offset = 0;
  if (rawValues.length < 52) { 
    Logger.log("parseCustomPacketToFields: rawValues length %s is less than 52. String: %s", rawValues.length, str);
    return null; 
  }
  try {
    for (var i = 0; i < NUM_SLOTS_CONST; i++) { 
      offset++; 
      prices.push((parseInt(rawValues[offset++])||0)*100 + (parseInt(rawValues[offset++])||0));
    }
    for (var e = 0; e < NUM_SLOTS_CONST; e++) {
      errors.push(parseInt(rawValues[offset++]) || 0);
    }
    var allowed_errors = parseInt(rawValues[offset++]) || 0;
    var speed_part1 = parseInt(rawValues[offset++]) || 0; 
    var speed_part2 = parseInt(rawValues[offset++]) || 0; 
    var speed_part3 = parseInt(rawValues[offset++]) || 0;
    var uart_speed = speed_part1*10000 + speed_part2*100 + speed_part3;
    return { prices, errors, allowed_errors, uart_speed };
  } catch (ex) {
    Logger.log("parseCustomPacketToFields Exception during parsing: %s. String: %s. RawValues: %s", ex.toString(), str, JSON.stringify(rawValues));
    return null;
  }
}

function loadMachineList(token) {
  if (!token) {
    Logger.log("loadMachineList: Токен не предоставлен.");
    return { error: "Token обязателен (machines).", items: [] };
  }
  var url = `${VENDISTA_API_BASE_URL}/machines?ItemsOnPage=1000&token=${encodeURIComponent(token)}`;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseCode = resp.getResponseCode(); var contentText = resp.getContentText();
    if (responseCode !== 200) {
      Logger.log("loadMachineList API Error: Code %s, Response: %s", responseCode, contentText);
      return { error: `Ошибка API (machines): ${responseCode}. ${contentText}`, items: [] };
    }
    var jsonData = JSON.parse(contentText || "{}");
    const machines = (jsonData.items || []).map(m => ({ id: m.id, name: m.name, address: m.address, product_matrix_id: m.product_matrix_id }));
    return machines;
  } catch (e) {
    Logger.log("loadMachineList Exception: %s", e.toString());
    return { error: `Исключение (machines): ${e.message}`, items: [] };
  }
}

function loadProductMatrix(productMatrixId, token) {
  if (!productMatrixId || !token) {
    Logger.log("loadProductMatrix: Отсутствует productMatrixId или токен. ID: %s, Token present: %s", productMatrixId, !!token);
    return { error: "ID матрицы и токен обязательны (productmatrix).", products: [] };
  }
  var url = `${VENDISTA_API_BASE_URL}/productmatrix/${productMatrixId}?token=${encodeURIComponent(token)}`;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseCode = resp.getResponseCode(); var contentText = resp.getContentText();
    if (responseCode !== 200) {
      Logger.log("loadProductMatrix API Error: Code %s, Response: %s", responseCode, contentText);
      return { error: `Ошибка API (productmatrix): ${responseCode}. ${contentText}`, products: [] };
    }
    var jsonData = JSON.parse(contentText || "{}");
    return (jsonData.item && jsonData.item.products) ? jsonData.item.products : [];
  } catch (e) {
    Logger.log("loadProductMatrix Exception: %s", e.toString());
    return { error: `Исключение (productmatrix): ${e.message}`, products: [] };
  }
}

function loadMachineIngredients(machineId, token) {
  if (!machineId || !token) {
    Logger.log("loadMachineIngredients: Отсутствует machineId или токен. MachineID: %s, Token present: %s", machineId, !!token);
    return { error: "ID машины и токен обязательны для загрузки ингредиентов.", items: [] };
  }
  var url = `${VENDISTA_API_BASE_URL}/machines/${encodeURIComponent(machineId)}/ingredients?token=${encodeURIComponent(token)}`;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseCode = resp.getResponseCode();
    var contentText = resp.getContentText();
    if (responseCode !== 200) {
      Logger.log("loadMachineIngredients API Error: Code %s, Response: %s", responseCode, contentText);
      return { error: `Ошибка API (ingredients): Статус ${responseCode}. Ответ: ${contentText}`, items: [] };
    }
    var jsonData = JSON.parse(contentText || "{}");
    if (jsonData && jsonData.items !== undefined && Array.isArray(jsonData.items)) {
        return jsonData.items; 
    } else {
        if (Array.isArray(jsonData)) return jsonData; 
        Logger.log("loadMachineIngredients: Некорректная структура ответа API. Response: %s", contentText);
        return { error: "Некорректная структура ответа от API ингредиентов.", items: [] };
    }
  } catch (e) {
    Logger.log("loadMachineIngredients Exception: %s", e.toString());
    return { error: `Исключение при запросе ингредиентов: ${e.message}`, items: [] };
  }
}

function sendCommand(params) {
  if (!params || !params.token || !params.terminalId || params.str_parameter1 === undefined) {
    Logger.log("sendCommand: Неполные параметры: %s", JSON.stringify(params));
    return { error: "Token, terminalId, и str_parameter1 обязательны (sendCommand, серверная проверка)." };
  }
  var url = `${VENDISTA_API_BASE_URL}/terminals/${encodeURIComponent(params.terminalId)}/commands?token=${encodeURIComponent(params.token)}`;
  var payload = { command_id: 73, parameter1: 1, str_parameter1: params.str_parameter1 };
  var options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  try {
    var resp = UrlFetchApp.fetch(url, options);
    return { status: resp.getResponseCode(), body: resp.getContentText() };
  } catch (e) {
    Logger.log("sendCommand Exception: %s", e.toString());
    return { error: `Исключение (sendCommand): ${e.message}` };
  }
}

function formatVendistaDateTime(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  const hours = date.getHours().toString().padStart(2, '0');
  const minutes = date.getMinutes().toString().padStart(2, '0');
  const seconds = date.getSeconds().toString().padStart(2, '0');
  const milliseconds = date.getMilliseconds().toString().padStart(3, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}.${milliseconds}`;
}

function fetchSalesList(params) { 
  if (!params || !params.token || !params.machineId || !params.timestampBeforeCommand) {
    Logger.log("fetchSalesList: Отсутствует токен, machineId или timestampBeforeCommand: %s", JSON.stringify(params));
    return { error: "Токен, ID автомата и время перед командой обязательны.", items: [] };
  }

  const dateFrom = new Date(new Date(params.timestampBeforeCommand).getTime() - 2 * 60 * 1000); 
  const dateTo = new Date(Date.now() + 10 * 60 * 1000); 

  const formattedDateFrom = formatVendistaDateTime(dateFrom);
  const formattedDateTo = formatVendistaDateTime(dateTo);

  let url = `${VENDISTA_API_BASE_URL}/sales/list?MachineId=${encodeURIComponent(params.machineId)}&DateFrom=${encodeURIComponent(formattedDateFrom)}&DateTo=${encodeURIComponent(formattedDateTo)}&ItemsOnPage=10&OrderDesc=true`;
  url += "&SellTypes=1&SellTypes=2&SellTypes=3&SellTypes=4"; 
  url += `&token=${encodeURIComponent(params.token)}`;

  Logger.log("GS: Запрос списка продаж URL: " + url);

  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseCode = resp.getResponseCode();
    var contentText = resp.getContentText();

    Logger.log("GS: fetchSalesList - Код ответа: " + responseCode);
    Logger.log("GS: fetchSalesList - Текст ответа (начало): " + contentText.substring(0, 1000));

    if (responseCode !== 200) {
      Logger.log("fetchSalesList API Error: Code %s, Response: %s", responseCode, contentText);
      return { error: `Ошибка API (sales/list): ${responseCode}. ${contentText}`, items: [] };
    }
    var jsonData = JSON.parse(contentText || "{}");
    Logger.log("GS: fetchSalesList - Распарсенные продажи (все полученные): " + JSON.stringify(jsonData.items || []));
    return jsonData.items || []; 
  } catch (e) {
    Logger.log("fetchSalesList Exception: %s", e.toString());
    return { error: `Исключение (sales/list): ${e.message}`, items: [] };
  }
}


// --- Функции для удаленного отчета ---
function startRemoteReportGeneration() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('lastProcessedMachineIndex');
  userProperties.deleteProperty('remoteReportMachineList');
  userProperties.deleteProperty('remoteReportStatus');
  userProperties.deleteProperty('totalMachinesForReport');
  deleteAllTriggersByName_('continueRemoteReport');

  console.log("Инициация удаленного отчета. Загрузка списка автоматов (используется VENDISTA_API_TOKEN из констант).");
  const machineListResponse = loadMachineList(VENDISTA_API_TOKEN); 

  if (machineListResponse.error || !machineListResponse.length) {
    const errorMsg = "Удаленный отчет: Не удалось загрузить список автоматов или список пуст: " + (machineListResponse.error || "Список пуст. Проверьте VENDISTA_API_TOKEN в Code.gs");
    console.error(errorMsg);
    try {
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME).getRange(1,1).setValue(errorMsg).setFontWeight("bold").setFontColor("red");
    } catch(e){ console.error("Не удалось записать ошибку в таблицу: " + e.toString()); }
    userProperties.setProperty('remoteReportStatus', 'Ошибка: ' + errorMsg);
    return "Ошибка загрузки списка автоматов: " + (machineListResponse.error || "Список пуст. Проверьте VENDISTA_API_TOKEN в Code.gs");
  }

  const totalMachines = machineListResponse.length;
  userProperties.setProperty('remoteReportMachineList', JSON.stringify(machineListResponse));
  userProperties.setProperty('lastProcessedMachineIndex', '0');
  userProperties.setProperty('totalMachinesForReport', totalMachines.toString());
  const initialStatus = `Инициализация отчета для ${totalMachines} автоматов... Ожидайте первого обновления.`;
  userProperties.setProperty('remoteReportStatus', initialStatus);

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.clearContents(); 
    sheet.getRange(1, 1).setValue("Генерация удаленного отчета начата: " + new Date().toLocaleString()).setFontWeight("bold");
    sheet.getRange(2, 1).setValue(initialStatus).setBackground("#fff3cd"); 
    SpreadsheetApp.flush();
  } catch (e) {
    console.error("Ошибка очистки/записи в таблицу для удаленного отчета: " + e.toString());
    userProperties.setProperty('remoteReportStatus', 'Ошибка таблицы: ' + e.toString());
    return "Ошибка работы с таблицей: " + e.toString();
  }

  console.log("Список автоматов загружен. Запускаем первый триггер продолжения.");
  ScriptApp.newTrigger('continueRemoteReport')
      .timeBased()
      .after(10 * 1000) 
      .create();
  
  return "Процесс генерации удаленного отчета запущен. Данные будут появляться в таблице постепенно.";
}

function continueRemoteReport() {
  const userProperties = PropertiesService.getUserProperties();
  // ... (получение machineListJson, lastProcessedIndex, totalMachines - без изменений) ...
  const machineListJson = userProperties.getProperty('remoteReportMachineList');
  let lastProcessedIndex = parseInt(userProperties.getProperty('lastProcessedMachineIndex') || '0');
  const totalMachines = parseInt(userProperties.getProperty('totalMachinesForReport') || '0');

  let sheet;
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) { 
        Logger.log("continueRemoteReport: Лист '" + SHEET_NAME + "' не найден. Остановка.");
        deleteAllTriggersByName_('continueRemoteReport');
        userProperties.setProperty('remoteReportStatus', 'Ошибка: Лист отчета не найден.');
        return;
    }
  } catch (e) {
     Logger.log("continueRemoteReport: Ошибка доступа к таблице. Остановка. " + e.toString());
     deleteAllTriggersByName_('continueRemoteReport');
     userProperties.setProperty('remoteReportStatus', 'Критическая ошибка доступа к таблице.');
     return;
  }

  if (!machineListJson) { 
    const errorMsg = "continueRemoteReport: Отсутствует список автоматов в UserProperties. Процесс остановлен.";
    Logger.log(errorMsg);
    try { sheet.getRange(2,1).setValue(errorMsg).setFontWeight("bold").setFontColor("red").setBackground(null); } catch(e){}
    deleteAllTriggersByName_('continueRemoteReport');
    userProperties.setProperty('remoteReportStatus', errorMsg);
    return;
  }
  
  let machineList;
  try {
    machineList = JSON.parse(machineListJson);
  } catch (e) { 
    const errorMsg = "continueRemoteReport: Ошибка парсинга JSON списка автоматов: " + e.message;
    Logger.log(errorMsg);
    try { sheet.getRange(2,1).setValue(errorMsg).setFontWeight("bold").setFontColor("red").setBackground(null); } catch(e){}
    deleteAllTriggersByName_('continueRemoteReport');
    userProperties.deleteProperty('lastProcessedMachineIndex');
    userProperties.deleteProperty('remoteReportMachineList');
    userProperties.deleteProperty('totalMachinesForReport');
    userProperties.setProperty('remoteReportStatus', errorMsg);
    return;
  }

  if (lastProcessedIndex >= machineList.length) { 
    const completionMsg = "Отчет полностью сгенерирован: " + new Date().toLocaleString();
    Logger.log("continueRemoteReport: Все автоматы обработаны. Завершение.");
    try { sheet.getRange(2,1).setValue(completionMsg).setFontWeight("bold").setBackground("#d4edda"); } catch(e){}
    deleteAllTriggersByName_('continueRemoteReport');
    userProperties.deleteProperty('lastProcessedMachineIndex');
    userProperties.deleteProperty('remoteReportMachineList');
    userProperties.deleteProperty('totalMachinesForReport');
    userProperties.setProperty('remoteReportStatus', 'Завершено: ' + completionMsg);
    return;
  }

  let currentStatus = `Обработка автоматов: ${lastProcessedIndex + 1}-${Math.min(lastProcessedIndex + MACHINES_PER_TRIGGER_RUN, machineList.length)} из ${totalMachines}. Пожалуйста, подождите...`;
  try { sheet.getRange(2,1).setValue(currentStatus).setFontWeight("normal").setBackground("#fff3cd"); SpreadsheetApp.flush(); } catch(e){}
  userProperties.setProperty('remoteReportStatus', currentStatus);

  let dataForBatchWrite = []; 
  let formattingRanges = []; 

  let tempCurrentRowOffset = 0; // Смещение строк внутри текущего батча dataForBatchWrite

  for (let i = 0; i < MACHINES_PER_TRIGGER_RUN; i++) {
    if (lastProcessedIndex >= machineList.length) break;

    const machine = machineList[lastProcessedIndex];
    const machineAddressForTable = machine.address || "";
    const machineHeaderString = machine.name + (machineAddressForTable ? ` (${machineAddressForTable})` : "");
    
    Logger.log(`Обработка автомата (удаленный отчет): ${machineHeaderString}`);
    
    // 1. Строка заголовка автомата
    let machineHeaderRowArray = [machineHeaderString];
    for(let j=1; j < NUM_COLUMNS_MAIN_TABLE; j++) machineHeaderRowArray.push("");
    dataForBatchWrite.push(machineHeaderRowArray);
    formattingRanges.push({type: "machineHeader", rowOffset: tempCurrentRowOffset });
    tempCurrentRowOffset++;
    
    // 2. Строка заголовка таблицы товаров
    dataForBatchWrite.push(["Отсек", "Наименование", "Цена (коп)", "Ошибка", "Остаток", "Вместимость", "% Остатка", "Адрес"]);
    formattingRanges.push({type: "tableHeader", rowOffset: tempCurrentRowOffset });
    tempCurrentRowOffset++;
    
    let machineDataLog = "";
    try {
      // ... (логика получения packetResp, products, ingredients - БЕЗ ИЗМЕНЕНИЙ) ...
      const zaprosParams = { token: VENDISTA_API_TOKEN, terminalId: machine.name, str_parameter1: "{zapros}" };
      const sendCmdResp = sendCommand(zaprosParams);
      if (sendCmdResp.error || (sendCmdResp.status && sendCmdResp.status !== 200 && sendCmdResp.status !== 202)){
         machineDataLog += `Ошибка отправки {zapros}: ${sendCmdResp.error || sendCmdResp.body}; `;
      }
      Utilities.sleep(DELAY_AFTER_ZAPROS_MS);

      const packetResp = updateData({ token: VENDISTA_API_TOKEN, terminalId: machine.name });
      
      let products = [];
      if (machine.product_matrix_id) {
        const matrixResp = loadProductMatrix(machine.product_matrix_id, VENDISTA_API_TOKEN);
        if (matrixResp.error) machineDataLog += `Ошибка матрицы: ${matrixResp.error}; `;
        else products = matrixResp || [];
      } else {
        machineDataLog += `Матрица не задана; `;
      }

      let ingredients = [];
      if (machine.id) { 
        const ingredResp = loadMachineIngredients(machine.id, VENDISTA_API_TOKEN);
        if (ingredResp.error) machineDataLog += `Ошибка остатков: ${ingredResp.error}; `;
        else ingredients = ingredResp || [];
      } else {
        machineDataLog += `Числовой ID для остатков (machine.id) не найден; `;
      }

      if (packetResp.found) {
        const mergedData = mergeAllDataServer(packetResp, products, ingredients);
        const filteredMergedData = mergedData.filter(dataRow => !(dataRow.productName && /^Отсек \d+$/.test(dataRow.productName.trim())));

        if (filteredMergedData.length > 0) {
            filteredMergedData.forEach(dataRow => {
              let percentageText = 'N/A';
              // ... (расчет percentageText) ...
              const loadingNum = parseFloat(dataRow.loading);
              const capacityNum = parseFloat(dataRow.capacity);
              if (!isNaN(loadingNum) && !isNaN(capacityNum) && capacityNum > 0) {
                  percentageText = Math.round((loadingNum / capacityNum) * 100) + '%';
              } else if (!isNaN(loadingNum) && !isNaN(capacityNum) && capacityNum === 0 && loadingNum === 0) {
                  percentageText = '0%';
              }
              dataForBatchWrite.push([
                dataRow.slot, dataRow.productName, dataRow.price, dataRow.error,
                dataRow.loading, dataRow.capacity, percentageText, machineAddressForTable
              ]);
              tempCurrentRowOffset++;
            });
        } else { 
            let emptyProductRow = ["-", "Нет активных товаров", "-", "-", "-", "-", "-", machineAddressForTable];
            dataForBatchWrite.push(emptyProductRow);
            tempCurrentRowOffset++;
        }

        let paramsInfoText = `Доп. ошибки: ${packetResp.allowed_errors !== null ? packetResp.allowed_errors : 'N/A'}, UART: ${packetResp.uart_speed !== null ? packetResp.uart_speed : 'N/A'}, Время данных: ${formatTimeServer(packetResp.timestamp)} | ${machineDataLog.trim()}`;
        let paramsRowArray = [paramsInfoText];
        for(let j=1; j < NUM_COLUMNS_MAIN_TABLE -1; j++) paramsRowArray.push("");
        paramsRowArray.push(machineAddressForTable); 
        dataForBatchWrite.push(paramsRowArray);
        formattingRanges.push({type: "paramsInfo", rowOffset: tempCurrentRowOffset });
        tempCurrentRowOffset++;

      } else {
        let errorText = `Данные пакета не найдены или ошибка: ${packetResp.error || 'нет данных'}. ${machineDataLog.trim()}`;
        let errorRowArray = [errorText];
        for(let j=1; j < NUM_COLUMNS_MAIN_TABLE -1; j++) errorRowArray.push("");
        errorRowArray.push(machineAddressForTable);
        dataForBatchWrite.push(errorRowArray);
        formattingRanges.push({type: "packetError", rowOffset: tempCurrentRowOffset });
        tempCurrentRowOffset++;
      }
    } catch (e) {
      let criticalErrorText = `Критическая ошибка при обработке ${machine.name}: ${e.toString()} ${machineDataLog.trim()}`;
      let criticalErrorRowArray = [criticalErrorText];
      for(let j=1; j < NUM_COLUMNS_MAIN_TABLE -1; j++) criticalErrorRowArray.push("");
      criticalErrorRowArray.push(machineAddressForTable);
      dataForBatchWrite.push(criticalErrorRowArray);
      formattingRanges.push({type: "criticalError", rowOffset: tempCurrentRowOffset });
      tempCurrentRowOffset++;
      console.error(`Критическая ошибка для ${machine.name}: ${e.toString()}`);
    }
    
    let emptySeparatorRowArray = [];
    for(let j=0; j < NUM_COLUMNS_MAIN_TABLE; j++) emptySeparatorRowArray.push("");
    dataForBatchWrite.push(emptySeparatorRowArray); 
    tempCurrentRowOffset++;

    lastProcessedIndex++;
    userProperties.setProperty('lastProcessedMachineIndex', lastProcessedIndex.toString());
    if (i < MACHINES_PER_TRIGGER_RUN - 1 && lastProcessedIndex < machineList.length) {
        Utilities.sleep(DELAY_BETWEEN_MACHINES_MS);
    }
  }

  if (dataForBatchWrite.length > 0) {
    try {
      let startRowForWrite = sheet.getLastRow() + 1;
      if (startRowForWrite <= 2 && sheet.getRange(1,1).getValue().toString().toLowerCase().includes("генерация удаленного отчета")) {
          startRowForWrite = 3;
      }
      
      sheet.getRange(startRowForWrite, 1, dataForBatchWrite.length, NUM_COLUMNS_MAIN_TABLE).setValues(dataForBatchWrite);
      
      formattingRanges.forEach(fmt => {
          const actualRowInSheet = startRowForWrite + fmt.rowOffset;
          try {
            if (fmt.type === "machineHeader") {
                sheet.getRange(actualRowInSheet, 1, 1, NUM_COLUMNS_MAIN_TABLE - 1).merge().setFontWeight("bold").setBackground("#f0f0f0").setHorizontalAlignment("left");
                sheet.getRange(actualRowInSheet, NUM_COLUMNS_MAIN_TABLE).clearFormat().clearContent(); // Очищаем ячейку адреса в строке заголовка
            } else if (fmt.type === "tableHeader") {
                sheet.getRange(actualRowInSheet, 1, 1, NUM_COLUMNS_MAIN_TABLE).setFontWeight("bold");
            } else if (fmt.type === "paramsInfo" || fmt.type === "packetError" || fmt.type === "criticalError") {
                sheet.getRange(actualRowInSheet, 1, 1, NUM_COLUMNS_MAIN_TABLE - 1).merge().setHorizontalAlignment("left");
                if (fmt.type === "packetError") sheet.getRange(actualRowInSheet, 1).setFontColor("orange");
                if (fmt.type === "criticalError") sheet.getRange(actualRowInSheet, 1).setFontColor("red");
            }
          } catch (eFormat) {
            Logger.log("Ошибка форматирования строки " + actualRowInSheet + " тип " + fmt.type + ": " + eFormat.toString());
          }
      });

    } catch (e) {
      Logger.log("Ошибка при записи dataForBatchWrite в лист: " + e.toString() + " Данные (первые 1000 символов): " + JSON.stringify(dataForBatchWrite).substring(0,1000));
    }
  }
  
  // ... (остальная часть continueRemoteReport с планированием следующего триггера или завершением - БЕЗ ИЗМЕНЕНИЙ) ...
  if (lastProcessedIndex < machineList.length) {
    const nextRunTime = new Date(Date.now() + DELAY_BETWEEN_TRIGGERS_MINUTES * 60 * 1000);
    const statusToSave = `Обработано ${lastProcessedIndex} из ${totalMachines}. Следующий запуск: ${Utilities.formatDate(nextRunTime, Session.getScriptTimeZone(), "HH:mm dd.MM.yyyy")}`;
    userProperties.setProperty('remoteReportStatus', statusToSave);
    try { sheet.getRange(2, 1).setValue(statusToSave).setFontWeight("normal").setBackground(null); } catch(e){}
    console.log(`Запланировано продолжение обработки. Следующий индекс: ${lastProcessedIndex}. Статус: ${statusToSave}`);
    
    deleteAllTriggersByName_('continueRemoteReport');
    ScriptApp.newTrigger('continueRemoteReport')
        .timeBased()
        .after(DELAY_BETWEEN_TRIGGERS_MINUTES * 60 * 1000)
        .create();
  } else {
    const completionMsg = "Отчет полностью сгенерирован: " + new Date().toLocaleString();
    console.log("Все автоматы обработаны в этом цикле триггеров.");
    try { sheet.getRange(2,1).setValue(completionMsg).setFontWeight("bold").setBackground("#d4edda"); } catch(e){}
    deleteAllTriggersByName_('continueRemoteReport');
    userProperties.deleteProperty('lastProcessedMachineIndex');
    userProperties.deleteProperty('remoteReportMachineList');
    userProperties.deleteProperty('totalMachinesForReport');
    userProperties.setProperty('remoteReportStatus', 'Завершено: ' + completionMsg);
  }
  try { SpreadsheetApp.flush(); } catch(e){}
}


// ... (mergeAllDataServer, formatTimeServer, deleteAllTriggersByName_, setupDailyTrigger, startRemoteReportGenerationDaily, getRemoteReportStatus - БЕЗ ИЗМЕНЕНИЙ)
function mergeAllDataServer(packetData, productMatrix, ingredientsData) {
  const tableRows = [];
  for (let i = 0; i < NUM_SLOTS_CONST; i++) {
    const slotNumber = i + 1;
    const product = productMatrix.find(p => p.item_id === slotNumber);
    let row = {
      slot: slotNumber,
      productName: product ? product.product_name : `Отсек ${slotNumber}`,
      price: (packetData.prices && packetData.prices[i] !== undefined) ? packetData.prices[i] : 0,
      error: (packetData.errors && packetData.errors[i] !== undefined) ? packetData.errors[i] : 0,
      loading: 'N/A', capacity: 'N/A', item_id: slotNumber
    };
    if (product && Array.isArray(ingredientsData) && ingredientsData.length > 0) {
      const matchedIngredient = ingredientsData.find(ing =>
        ing.ingredient_name && product.product_name &&
        ing.ingredient_name.trim().toLowerCase() === product.product_name.trim().toLowerCase()
      );
      if (matchedIngredient) {
        row.loading = matchedIngredient.loading !== undefined ? matchedIngredient.loading : 'N/A';
        row.capacity = matchedIngredient.capacity !== undefined ? matchedIngredient.capacity : 'N/A';
      }
    }
    tableRows.push(row);
  }
  return tableRows;
}

function formatTimeServer(timeStr) {
  if (!timeStr) return "неизвестно";
  try {
    const d = new Date(timeStr);
    if (isNaN(d.getTime()) || d.getTime() === 0 || (d.getUTCFullYear() === 1970 && d.getUTCMonth() === 0 && d.getUTCDate() === 1)) {
        if (timeStr.includes("0001-01-01") || timeStr.includes("1970-01-01T00:00:00")) return "неизвестно";
    }
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  } catch (e) { return timeStr; }
}

function deleteAllTriggersByName_(handlerFunctionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerFunctionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function setupDailyTrigger() {
  deleteAllTriggersByName_('startRemoteReportGenerationDaily');
  const triggers = ScriptApp.getProjectTriggers();
  let triggerExists = false;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'startRemoteReportGenerationDaily') {
      triggerExists = true;
      break;
    }
  }
  if (!triggerExists) {
    ScriptApp.newTrigger('startRemoteReportGenerationDaily')
        .timeBased()
        .atHour(3) 
        .everyDays(1)
        .create();
    console.log("Ежедневный триггер на 3:00 успешно установлен для startRemoteReportGenerationDaily.");
  } else {
    console.log("Ежедневный триггер для startRemoteReportGenerationDaily уже существует.");
  }
}

function startRemoteReportGenerationDaily() {
  startRemoteReportGeneration();
}

function getRemoteReportStatus() {
  return PropertiesService.getUserProperties().getProperty('remoteReportStatus') || "Статус удаленного отчета не определен.";
}

function loadDataFromSheet(params) {
  const { terminalName, token, productMatrixId } = params; 
  if (!terminalName) { /* ... */ return { found: false, error: "Имя терминала обязательно для загрузки из таблицы." }; }

  let sheet;
  try { sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) { /* ... */ return { found: false, error: `Лист "${SHEET_NAME}" не найден.` }; }
  } catch (e) { /* ... */ return { found: false, error: "Ошибка доступа к таблице: " + e.message }; }

  const data = sheet.getDataRange().getValues();
  let machineBlockStartIndex = -1;
  let actualProductRowsCountForMachine = 0;

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] && typeof data[i][0] === 'string') {
       const cellContent = data[i][0].trim(); 
       const nameInCell = cellContent.includes("(") ? cellContent.substring(0, cellContent.indexOf("(")).trim() : cellContent;
       if (nameInCell === terminalName.trim()) {
           if (i + 1 < data.length && data[i+1][0] && data[i+1][0].toString().trim().toLowerCase() === "отсек") {
             machineBlockStartIndex = i;
             for (let k = i + 2; k < data.length; k++) {
                 if (!data[k][0] || data[k][0].toString().trim() === "" || 
                     data[k][0].toString().toLowerCase().startsWith("доп. ошибки") || 
                     data[k][0].toString().toLowerCase().startsWith("данные пакета не найдены") || 
                     data[k][0].toString().toLowerCase().startsWith("критическая ошибка")) {
                     break;
                 }
                 actualProductRowsCountForMachine++;
             }
             break;
           }
       }
    }
  }

  if (machineBlockStartIndex === -1) { /* ... */ return { found: false, error: `Данные для автомата "${terminalName}" не найдены в таблице.` }; }
  
  const headerRowIndex = machineBlockStartIndex + 1;
  if (headerRowIndex >= data.length ||
      !(data[headerRowIndex][0] && data[headerRowIndex][0].toString().trim().toLowerCase() === "отсек") ||
      !(data[headerRowIndex][2] && data[headerRowIndex][2].toString().trim().toLowerCase().startsWith("цена"))) {
    Logger.log("loadDataFromSheet: Не найдена таблица товаров для '%s' после его заголовка. Ожидался заголовок на строке %s.", terminalName, headerRowIndex + 1);
    return { found: false, error: `Не найдена таблица товаров для "${terminalName}" после его заголовка.` };
  }

  const prices = []; const errors = []; const sheetRowData = [];
  const dataStartIndex = headerRowIndex + 1;

  for (let i = 0; i < NUM_SLOTS_CONST; i++) {
    let rowValues = { slot: i + 1, productName: `Отсек ${i+1}`, price: 0, error: 0, loading: 'N/A', capacity: 'N/A', item_id: i + 1 };
    if (i < actualProductRowsCountForMachine) {
        const rowIndex = dataStartIndex + i;
        if (rowIndex < data.length && data[rowIndex] && data[rowIndex][0] && data[rowIndex][0].toString().trim() !== "") {
            const row = data[rowIndex];
             if (parseInt(row[0]) === (i+1) || actualProductRowsCountForMachine < NUM_SLOTS_CONST) { // Если слоты идут по порядку ИЛИ если товаров меньше 12 (значит это реальный товар)
                rowValues.slot = parseInt(row[0]) || (i+1); // Берем номер из таблицы или i+1 если там пусто
                rowValues.productName = row[1] ? row[1].toString() : `Отсек ${rowValues.slot}`; 
                rowValues.price = parseInt(row[2]) || 0;
                rowValues.error = parseInt(row[3]) || 0;
                rowValues.loading = row[4] !== undefined && row[4] !== null ? row[4].toString() : 'N/A';
                rowValues.capacity = row[5] !== undefined && row[5] !== null ? row[5].toString() : 'N/A';
                rowValues.item_id = rowValues.slot;
             }
        }
    }
    prices.push(rowValues.price); 
    errors.push(rowValues.error); 
    sheetRowData.push(rowValues); 
  }

  let allowed_errors = null; let uart_speed = null; let timestamp = null;
  const paramsRowIndexExpected = dataStartIndex + actualProductRowsCountForMachine; 
  
  if (paramsRowIndexExpected < data.length && data[paramsRowIndexExpected] && data[paramsRowIndexExpected][0] && typeof data[paramsRowIndexExpected][0] === 'string') {
      const paramsString = data[paramsRowIndexExpected][0].toString().trim(); 
      Logger.log(`loadDataFromSheet: [${terminalName}] Parsing paramsString: '${paramsString}' at index ${paramsRowIndexExpected}`);

      const allowedErrorsMatch = paramsString.match(/Доп\. ошибки:\s*([^,]+)/i);
      if (allowedErrorsMatch && allowedErrorsMatch[1] && allowedErrorsMatch[1].trim().toLowerCase() !== 'n/a') {
        allowed_errors = parseInt(allowedErrorsMatch[1].trim());
      }
      
      const uartSpeedMatch = paramsString.match(/UART:\s*([^,]+)/i);
      if (uartSpeedMatch && uartSpeedMatch[1] && uartSpeedMatch[1].trim().toLowerCase() !== 'n/a') {
        uart_speed = parseInt(uartSpeedMatch[1].trim());
      }
      
      const timeMatch = paramsString.match(/Время данных:\s*([\d.:\s]+)/i);
      if (timeMatch && timeMatch[1] && timeMatch[1].trim().toLowerCase() !== 'неизвестно') {
          const timeStringFromSheet = timeMatch[1].trim();
          const parts = timeStringFromSheet.match(/(\d{2})\.(\d{2})\.(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
          if (parts) {
            // Создаем дату, считая, что компоненты из строки УЖЕ в часовом поясе скрипта.
            // new Date(year, monthIndex, day, hours, minutes, seconds)
            var scriptLocalDate = new Date(
                parseInt(parts[3]),      // year
                parseInt(parts[2]) - 1,  // month (0-11)
                parseInt(parts[1]),      // day
                parseInt(parts[4]),      // hours
                parseInt(parts[5]),      // minutes
                parseInt(parts[6])       // seconds
            );
            // Важно! Если часовой пояс скрипта не UTC, то toISOString() преобразует это локальное время в UTC.
            // Если часовой пояс скрипта UTC, то toISOString() вернет то же время с 'Z'.
            timestamp = scriptLocalDate.toISOString(); 
            Logger.log(`loadDataFromSheet: [${terminalName}] Time from sheet: "${timeStringFromSheet}", Script TimeZone: ${Session.getScriptTimeZone()}, Created Date obj: ${scriptLocalDate}, ISO Timestamp: ${timestamp}`);
          } else {
            timestamp = timeStringFromSheet; // Если не распарсили, передаем как есть
             Logger.log(`loadDataFromSheet: [${terminalName}] Time string "${timeStringFromSheet}" did not match dd.MM.yyyy HH:mm:ss format.`);
          }
      }
      Logger.log(`loadDataFromSheet: [${terminalName}] Parsed - AE: ${allowed_errors}, UART: ${uart_speed}, TS: ${timestamp}`);
  } else {
    Logger.log(`loadDataFromSheet: [${terminalName}] Строка с доп. параметрами не найдена на ${paramsRowIndexExpected + 1}. Факт. строк товаров: ${actualProductRowsCountForMachine}`);
  }
  
  if (productMatrixId && token) {
      const matrixProducts = loadProductMatrix(productMatrixId, token);
      if (matrixProducts && !matrixProducts.error) {
          sheetRowData.forEach(sRow => {
              const productFromMatrix = matrixProducts.find(p => p.item_id === sRow.slot);
              if (productFromMatrix && sRow.productName.startsWith("Отсек ")) { 
                  sRow.productName = productFromMatrix.product_name; 
              }
          });
      } else {
          Logger.log(`loadDataFromSheet: Не удалось загрузить матрицу ${productMatrixId} для ${terminalName} при загрузке из таблицы для обновления имен: %s`, matrixProducts.error);
      }
  }

  return {
    found: true, source: 'sheet', prices: prices, errors: errors,
    allowed_errors: allowed_errors, uart_speed: uart_speed, timestamp: timestamp,
    sheetRowData: sheetRowData, 
  };
}

function loadAllMachinesDataFromSheet() { // ИСПРАВЛЕН ПАРСИНГ ДОП. ПАРАМЕТРОВ
  let sheet;
  try { sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) { /* ... */ return { error: `Лист "${SHEET_NAME}" не найден.` }; }
  } catch (e) { /* ... */ return { error: "Ошибка доступа к таблице: " + e.message }; }

  const allData = sheet.getDataRange().getValues();
  const machinesReport = [];
  let currentMachineData = null;
  let productRowsReadForCurrentMachine = 0;

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const headerCell = row[0] ? row[0].toString().trim() : ""; 

    if (i < 2 && (headerCell.toLowerCase().includes("генерация удаленного отчета начата") ||
                   headerCell.toLowerCase().includes("инициализация отчета для") ||
                   headerCell.toLowerCase().includes("обработка автоматов:") ||
                   headerCell.toLowerCase().includes("отчет полностью сгенерирован") )) {
        continue;
    }
    
    if (headerCell && headerCell.trim() !== "" &&
        headerCell.toLowerCase() !== "отсек" &&
        !headerCell.toLowerCase().startsWith("доп. ошибки:") &&
        !headerCell.toLowerCase().startsWith("данные пакета не найдены") &&
        !headerCell.toLowerCase().startsWith("критическая ошибка при обработке") &&
        (i + 1 < allData.length && allData[i+1][0] && allData[i+1][0].toString().trim().toLowerCase() === "отсек")
       ) {

      if (currentMachineData) {
        machinesReport.push(currentMachineData);
      }
      
      let machineName = headerCell;
      let machineAddress = "";
      const addressMatch = headerCell.match(/\(([^)]+)\)$/);
      if (addressMatch && addressMatch[1]) {
        machineAddress = addressMatch[1];
        machineName = headerCell.substring(0, headerCell.lastIndexOf('(')).trim();
      }

      currentMachineData = {
        name: machineName,
        address: machineAddress,
        rows: [],
        allowed_errors: null,
        uart_speed: null,
        timestamp: null,
        parseError: null
      };
      productRowsReadForCurrentMachine = 0; 
      i++; 
      continue; 
    }
    
    if (currentMachineData && 
        row[0] && row[0].toString().trim() !== "" && // Убедимся, что есть номер отсека
        !(row[0].toString().toLowerCase().startsWith("доп. ошибки")) &&
        !(row[0].toString().toLowerCase().startsWith("данные пакета не найдены")) &&
        !(row[0].toString().toLowerCase().startsWith("критическая ошибка"))
        ) {
        
        if (row[1] && row[1].toString().trim() !== "Нет активных товаров") {
            let rowValues = { 
                slot: parseInt(row[0]) || (productRowsReadForCurrentMachine + 1), 
                productName: row[1] ? row[1].toString() : `Отсек ${parseInt(row[0]) || (productRowsReadForCurrentMachine + 1)}`, 
                price: parseInt(row[2]) || 0,
                error: parseInt(row[3]) || 0, 
                loading: row[4] !== undefined && row[4] !== null ? row[4].toString() : 'N/A',
                capacity: row[5] !== undefined && row[5] !== null ? row[5].toString() : 'N/A',
                item_id: parseInt(row[0]) || (productRowsReadForCurrentMachine + 1)
            };
            currentMachineData.rows.push(rowValues);
        }
        productRowsReadForCurrentMachine++;

        // Проверяем, не достигли ли мы конца блока товаров для этого автомата
        let nextRowIsParamOrEnd = false;
        if (i + 1 >= allData.length) { // Конец таблицы
            nextRowIsParamOrEnd = true;
        } else {
            const nextRowFirstCell = allData[i+1][0] ? allData[i+1][0].toString().toLowerCase() : "";
            if (nextRowFirstCell.startsWith("доп. ошибки") || 
                nextRowFirstCell.startsWith("данные пакета не найдены") || 
                nextRowFirstCell.startsWith("критическая ошибка") || 
                nextRowFirstCell.trim() === "" // Пустая строка-разделитель
            ) {
                nextRowIsParamOrEnd = true;
            }
        }
        // Если это последняя возможная строка товара (12-я) ИЛИ следующая строка - это параметры/ошибка/пустая
        if (productRowsReadForCurrentMachine >= NUM_SLOTS_CONST || nextRowIsParamOrEnd) {
          const paramsRowIndex = i + 1; 
          if (paramsRowIndex < allData.length && allData[paramsRowIndex] && allData[paramsRowIndex][0]) {
            const paramsString = allData[paramsRowIndex][0].toString().trim();
            Logger.log(`loadAllMachinesDataFromSheet: [${currentMachineData.name}] Parsing paramsString: '${paramsString}' at sheet row ${paramsRowIndex + 1}`);

            if (paramsString.toLowerCase().startsWith("данные пакета не найдены") || paramsString.toLowerCase().startsWith("критическая ошибка при обработке")) {
                currentMachineData.parseError = paramsString;
            } else if (paramsString.toLowerCase().startsWith("доп. ошибки:")) {
                const allowedErrorsMatch = paramsString.match(/Доп\. ошибки:\s*([^,]+)/i);
                if (allowedErrorsMatch && allowedErrorsMatch[1] && allowedErrorsMatch[1].trim().toLowerCase() !== 'n/a') {
                  currentMachineData.allowed_errors = parseInt(allowedErrorsMatch[1].trim());
                }
                const uartSpeedMatch = paramsString.match(/UART:\s*([^,]+)/i);
                if (uartSpeedMatch && uartSpeedMatch[1] && uartSpeedMatch[1].trim().toLowerCase() !== 'n/a') {
                  currentMachineData.uart_speed = parseInt(uartSpeedMatch[1].trim());
                }
                const timeMatch = paramsString.match(/Время данных:\s*([\d.:\s]+)/i);
                if (timeMatch && timeMatch[1] && timeMatch[1].trim().toLowerCase() !== 'неизвестно') {
                  const timeStringFromSheet = timeMatch[1].trim();
                  const parts = timeStringFromSheet.match(/(\d{2})\.(\d{2})\.(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
                  if (parts) {
                    var scriptLocalDate = new Date(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1]), parseInt(parts[4]), parseInt(parts[5]), parseInt(parts[6]));
                    currentMachineData.timestamp = scriptLocalDate.toISOString();
                     Logger.log(`loadAllMachinesDataFromSheet: [${currentMachineData.name}] Time from sheet: "${timeStringFromSheet}", ScriptTZ: ${Session.getScriptTimeZone()}, Created Date: ${scriptLocalDate}, ISO TS: ${currentMachineData.timestamp}`);
                  } else {
                    currentMachineData.timestamp = timeStringFromSheet; 
                    Logger.log(`loadAllMachinesDataFromSheet: [${currentMachineData.name}] Time string "${timeStringFromSheet}" did not match dd.MM.yyyy HH:mm:ss format.`);
                  }
                }
            }
            i++; 
          }
          if (i + 1 < allData.length && allData[i+1] && allData[i+1][0] === "") { 
            i++;
          }
          // Сбрасываем счетчик, т.к. закончили с этим автоматом (или он будет сброшен при нахождении нового заголовка)
          productRowsReadForCurrentMachine = NUM_SLOTS_CONST; // Чтобы не войти в этот if снова для этого автомата
        }
        continue;
    }
  }

  if (currentMachineData) { 
    machinesReport.push(currentMachineData);
  }

  if (machinesReport.length === 0 && allData.length > 2) {
      return { error: "Не удалось найти данные автоматов в таблице. Проверьте формат листа или запустите удаленный отчет.", items: [] };
  }
  Logger.log("loadAllMachinesDataFromSheet: Возвращается " + machinesReport.length + " автоматов. Пример последнего: " + JSON.stringify(machinesReport[machinesReport.length-1]));
  return { items: machinesReport };
}
