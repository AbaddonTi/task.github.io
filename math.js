
/********** STYLE FUNCTIONS **********/
function setCellStyle(ri, ci, property, value, sheetIndex = 0) {
    const { styles, rows } = s.datas[sheetIndex];
    const cell = rows.getCellOrNew(ri, ci);
    let cstyle = cell.style !== undefined ? { ...styles[cell.style] } : {};
    if (property.startsWith('font')) {
        const [fontProp] = property.split('-').slice(1);
        cstyle.font = { ...(cstyle.font || {}), [fontProp]: value };
    } else {
        cstyle[property] = value;
    }
    cell.style = s.datas[sheetIndex].addStyle(cstyle);
}

const setCellBorder = (ri, ci, borderStyle, sheetIndex = 0) => setCellStyle(ri, ci, 'border', borderStyle, sheetIndex);
const setCellBackgroundColor = (row, col, color, sheetIndex = 0) => setCellStyle(row, col, 'bgcolor', color, sheetIndex);
const enableTextWrap = (row, col, sheetIndex = 0) => setCellStyle(row, col, 'textwrap', true, sheetIndex);

/********** DIMENSION FUNCTIONS **********/
function setColumnWidth(colIndex, width, sheetIndex = 0) {
    s.datas[sheetIndex].cols.setWidth(colIndex, width);
}

function setRowHeight(rowIndex, height, sheetIndex = 0) {
    s.datas[sheetIndex].rows.setHeight(rowIndex, height);
}

/********** TEXT FUNCTIONS **********/
function setCellText(row, col, text, sheetIndex = 0) {
    const cell = s.datas[sheetIndex].rows.getCellOrNew(row, col);
    cell.text = text;
}

/********** INITIALIZATION FUNCTIONS **********/
const defaultBorderStyle = { style: 'thin', color: '#000000' };
const getBorderStyle = style => ({
    top: [style, defaultBorderStyle.color],
    bottom: [style, defaultBorderStyle.color],
    right: [style, defaultBorderStyle.color],
    left: [style, defaultBorderStyle.color]
});

function initializeSpreadsheetDB() {
    const thickBorderStyle = getBorderStyle('thick');
    
    [[0, 0], [0, 1]].forEach(([row, col]) => setCellBorder(row, col, thickBorderStyle));
    [[0, 0, 'Статья'], [0, 1, 'Процессинг']].forEach(([row, col, text]) => setCellText(row, col, text, 0));
    [[0, 150], [1, 150]].forEach(([col, width]) => setColumnWidth(col, width, 0));
    setRowHeight(0, 30, 0);
    [[0, 0], [0, 1]].forEach(([row, col]) => enableTextWrap(row, col, 0));
    [
        { row: 6, text: 'Цена закуп. карты' },
        { row: 7, text: 'Кол-во карт' },
        { row: 8, text: 'Профит с 1 карты' },
        { row: 10, text: 'Выручка наша:' },
        { row: 12, text: '% проблемных' },
        { row: 13, text: 'Сумма замороженных денег в проблемных картах' },
        { row: 14, text: 'Кол-во карт проблемных' },
        { row: 16, text: '% блока' },
        { row: 17, text: 'Сумма в блоки (заморозка)' },
        { row: 18, text: 'Кол-во карт в блок' },
        { row: 20, text: 'закуп карточки' },
        { row: 21, text: '% парней с выручки на руки' },
        { row: 22, text: 'ФОТ операторы' },
        { row: 24, text: 'Итого: переменные косты на подподнаправления' },
        { row: 25, text: 'DB 1' }
    ].forEach(({ row, text }) => setCellText(row, 0, text, 0));
}

/********** DYNAMIC RENDERING FUNCTIONS **********/
function renderUniqueBankValues(filteredRecordsDB1, filteredRecordsDB2) {
    const bankData = processBankRecords(filteredRecordsDB1);
    const methodColIndexes = renderBankData(bankData);
    renderOperationTypes(filteredRecordsDB2, methodColIndexes);
    s.reRender();
}

function processBankRecords(filteredRecordsDB1) {
    const bankMethodMap = {},
        bankPricesMap = {},
        bankCountsMap = {},
        bankProfitsMap = {},
        bankProblemCountsMap = {},
        bankBlockCountsMap = {};

    let totalPrices = 0,
        totalPricesCount = 0,
        totalBanksCount = 0,
        totalProfits = 0,
        totalProfitsCount = 0,
        totalProblemCount = 0,
        totalBlockCount = 0;

    filteredRecordsDB1.forEach(record => {
        const bank = Array.isArray(record['Банк']) ? record['Банк'][0] : record['Банк'];
        const method = Array.isArray(record['Метод']) ? record['Метод'][0] : record['Метод'];
        const price = parseFloat(record['Цена']);
        const profit = parseFloat(record['Профит']);
        const status = Array.isArray(record['Статус']) ? record['Статус'] : [record['Статус']];

        if (bank && method) {
            bankMethodMap[method] = bankMethodMap[method] || new Set();
            bankMethodMap[method].add(bank);

            bankPricesMap[method] = bankPricesMap[method] || {};
            bankPricesMap[method][bank] = bankPricesMap[method][bank] || [];
            if (!isNaN(price)) {
                bankPricesMap[method][bank].push(price);
                totalPrices += price;
                totalPricesCount += 1;
            }

            bankCountsMap[method] = bankCountsMap[method] || {};
            bankCountsMap[method][bank] = (bankCountsMap[method][bank] || 0) + 1;
            totalBanksCount += 1;

            bankProfitsMap[method] = bankProfitsMap[method] || {};
            bankProfitsMap[method][bank] = bankProfitsMap[method][bank] || [];
            if (!isNaN(profit)) {
                bankProfitsMap[method][bank].push(profit);
                totalProfits += profit;
                totalProfitsCount += 1;
            }

            if (status.includes('проблема')) {
                bankProblemCountsMap[method] = bankProblemCountsMap[method] || {};
                bankProblemCountsMap[method][bank] = (bankProblemCountsMap[method][bank] || 0) + 1;
                totalProblemCount += 1;
            }

            if (status.includes('блокировка')) {
                bankBlockCountsMap[method] = bankBlockCountsMap[method] || {};
                bankBlockCountsMap[method][bank] = (bankBlockCountsMap[method][bank] || 0) + 1;
                totalBlockCount += 1;
            }
        }
    });

    return {
        bankMethodMap,
        bankPricesMap,
        bankCountsMap,
        bankProfitsMap,
        bankProblemCountsMap,
        bankBlockCountsMap,
        totalPrices,
        totalPricesCount,
        totalBanksCount,
        totalProfits,
        totalProfitsCount,
        totalProblemCount,
        totalBlockCount
    };
}

function renderBankData(bankData) {
    const {
        bankMethodMap,
        bankPricesMap,
        bankCountsMap,
        bankProfitsMap,
        bankProblemCountsMap,
        bankBlockCountsMap,
        totalPrices,
        totalPricesCount,
        totalBanksCount,
        totalProfits,
        totalProfitsCount,
        totalProblemCount,
        totalBlockCount
    } = bankData;

    const averagePrice = totalPricesCount > 0 ? (totalPrices / totalPricesCount).toFixed(2) : '0.00';
    const averageProfit = totalProfitsCount > 0 ? (totalProfits / totalProfitsCount).toFixed(2) : '0.00';
    const sumProfit = totalProfits.toFixed(2);

    let colIndex = 1;
    const methodColIndexes = {};

    Object.keys(bankMethodMap).forEach(method => {
        const banks = Array.from(bankMethodMap[method]);
        const endColIndex = colIndex + banks.length - 1;

        methodColIndexes[method] = colIndex; // Store column index for each method

        setCellText(2, colIndex, method, 0);
        setCellBorder(2, colIndex, getBorderStyle('thick'));
        for (let i = colIndex; i <= endColIndex; i++) {
            setCellBorder(3, i, getBorderStyle('thin'));
        }
        setCellBackgroundColor(0, colIndex, '#d9ead3', 0);
        for (let i = colIndex + 1; i <= endColIndex; i++) {
            setCellBackgroundColor(0, i, '#d9ead3', 0);
        }

        banks.forEach((bank, bankIndex) => {
            const bankColIndex = colIndex + bankIndex;
            const colLetter = String.fromCharCode(66 + bankColIndex - 1);
            setCellText(3, bankColIndex, bank, 0);

            const prices = bankPricesMap[method][bank];
            if (prices && prices.length) {
                const averagePrice = (prices.reduce((a, b) => a + b, 0) / prices.length).toFixed(2);
                setCellText(6, bankColIndex, averagePrice, 0);
            }

            setCellText(7, bankColIndex, bankCountsMap[method][bank], 0);

            const profits = bankProfitsMap[method][bank];
            if (profits && profits.length) {
                const averageProfit = (profits.reduce((a, b) => a + b, 0) / profits.length).toFixed(2);
                setCellText(8, bankColIndex, averageProfit, 0);
            }

            setCellText(10, bankColIndex, `=${colLetter}8*${colLetter}9`, 0);
            setCellText(12, bankColIndex, `=${colLetter}15/${colLetter}8*100`, 0);
            setCellText(13, bankColIndex, `=${colLetter}7*${colLetter}15`, 0);
            setCellText(14, bankColIndex, bankProblemCountsMap[method]?.[bank] || 0, 0);
            setCellText(16, bankColIndex, `=${colLetter}19/${colLetter}8*100`, 0);
            setCellText(17, bankColIndex, `=${colLetter}7*${colLetter}19`, 0);
            setCellText(18, bankColIndex, bankBlockCountsMap[method]?.[bank] || 0, 0);
            setCellText(20, bankColIndex, `=${colLetter}7*${colLetter}8`, 0);
            setCellText(21, bankColIndex, '0.01', 0);
            setCellText(22, bankColIndex, `=(${colLetter}11-${colLetter}18)*${colLetter}22`, 0);
            setCellText(24, bankColIndex, `=${colLetter}14+${colLetter}18+${colLetter}21+${colLetter}23`, 0);
            setCellText(25, bankColIndex, `=${colLetter}11-${colLetter}25`, 0);
        });
        colIndex = endColIndex + 1;
    });

    setCellText(0, colIndex, 'итого Процессинг', 0);
    setCellBorder(0, colIndex, getBorderStyle('thick'), 0);
    setCellBackgroundColor(0, colIndex, '#d9ead3', 0);

    setCellText(6, colIndex, averagePrice, 0);
    setCellText(7, colIndex, totalBanksCount, 0);
    setCellText(8, colIndex, averageProfit, 0);
    setCellText(10, colIndex, sumProfit, 0);
    setCellText(12, colIndex, `=${String.fromCharCode(66 + colIndex - 1)}15/${String.fromCharCode(66 + colIndex - 1)}8*100`, 0);
    setCellText(13, colIndex, `=${String.fromCharCode(66 + colIndex - 1)}7*${String.fromCharCode(66 + colIndex - 1)}15`, 0);
    setCellText(14, colIndex, totalProblemCount, 0);
    setCellText(16, colIndex, `=${String.fromCharCode(66 + colIndex - 1)}19/${String.fromCharCode(66 + colIndex - 1)}8*100`, 0);
    setCellText(17, colIndex, `=${String.fromCharCode(66 + colIndex - 1)}7*${String.fromCharCode(66 + colIndex - 1)}19`, 0);
    setCellText(18, colIndex, totalBlockCount, 0);
    setCellText(20, colIndex, `=${String.fromCharCode(66 + colIndex - 1)}7*${String.fromCharCode(66 + colIndex - 1)}8`, 0);
    setCellText(21, colIndex, '0.01', 0);
    setCellText(22, colIndex, `=(${String.fromCharCode(66 + colIndex - 1)}11-${String.fromCharCode(66 + colIndex - 1)}18)*${String.fromCharCode(66 + colIndex - 1)}22`, 0);
    setCellText(24, colIndex, `=${String.fromCharCode(66 + colIndex - 1)}14+${String.fromCharCode(66 + colIndex - 1)}18+${String.fromCharCode(66 + colIndex - 1)}21+${String.fromCharCode(66 + colIndex - 1)}23`, 0);
    setCellText(25, colIndex, `=${String.fromCharCode(66 + colIndex - 1)}11-${String.fromCharCode(66 + colIndex - 1)}25`, 0);

    for (let i = 1; i <= colIndex; i++) {
        setCellBackgroundColor(0, i, '#d9ead3', 0);
    }

    return { colIndex, methodColIndexes };
}

// Функция для получения текущего курса USD/RUB из всех записей
function getExchangeRate(records) {
    // Проходим по всем записям, пока не найдем первую, в которой указан курс
    for (let record of records) {
        if (record['Бухгалтерия_Курс_Usd_Rub']) {
            return parseFloat(record['Бухгалтерия_Курс_Usd_Rub']);
        }
    }
    return null; // Возвращаем null, если курс не найден
}

// Обновленная функция renderOperationTypes
function renderOperationTypes(filteredRecordsDB2, { colIndex, methodColIndexes }) {
    const excludedOperations = new Set([
        "Пересчёт кассы",
        "Перевод (получаем на счёт)",
        "Внесение Слава",
        "Внесение Patek",
        "Внесение Роман",
        "Вывод Роман",
        "Вывод Слава",
        "Вывод Patek",
        "Выплаты Consulting",
        "Перевод (снимаем с счёта)"
    ]);
    const uniqueOperations = new Set();
    const operationSums = {};

    // Получаем значение курса USD/RUB без фильтрации
    const exchangeRate = getExchangeRate(mappedRecords);

    let rowIndex = 28;
    const totalSums = { 'Процессинг//Переводы': 0, 'Процессинг//Наличка': 0, 'Итого': 0 };

    // Записываем все уникальные операции и суммы
    filteredRecordsDB2.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sum = parseFloat(record['Бухгалтерия_Сумма']);

        if (operation && !excludedOperations.has(operation) && (project === 'Процессинг//Переводы' || project === 'Процессинг//Наличка')) {
            uniqueOperations.add(operation);
            if (!isNaN(sum)) {
                operationSums[operation] = operationSums[operation] || { 'Процессинг//Переводы': 0, 'Процессинг//Наличка': 0 };
                operationSums[operation][project] += sum;
            }
        }
    });

    // Перебор и запись уникальных операций и их сумм
    uniqueOperations.forEach(operation => {
        setCellText(rowIndex, 0, operation, 0);
        if (operationSums[operation]) {
            const переводSum = operationSums[operation]['Процессинг//Переводы'];
            const наличкаSum = operationSums[operation]['Процессинг//Наличка'];

            const totalSum = (переводSum + наличкаSum).toFixed(2);
            setCellText(rowIndex, colIndex, totalSum, 0); // Колонка "итого Процессинг"

            if (methodColIndexes['Переводы']) {
                setCellText(rowIndex, methodColIndexes['Переводы'], переводSum.toFixed(2), 0);
                totalSums['Процессинг//Переводы'] += переводSum;
            }
            if (methodColIndexes['Наличка']) {
                setCellText(rowIndex, methodColIndexes['Наличка'], наличкаSum.toFixed(2), 0);
                totalSums['Процессинг//Наличка'] += наличкаSum;
            }

            totalSums['Итого'] += parseFloat(totalSum);
        }
        rowIndex++;
    });

    // Отступ одной строки
    rowIndex++;

    // Строка для текущего курса USD/RUB
    setCellText(rowIndex, 0, "Текущий курс USD/RUB:", 0);
    if (exchangeRate !== null) {
        setCellText(rowIndex, 1, exchangeRate.toFixed(2), 0);
    }

    // Переход к следующей строке для итогов в долларах
    rowIndex++;

    // Итоговая строка для фиксированных костов в долларах
    setCellText(rowIndex, 0, "$ Итого: фикс косты на направление", 0);
    if (methodColIndexes['Переводы']) {
        setCellText(rowIndex, methodColIndexes['Переводы'], totalSums['Процессинг//Переводы'].toFixed(2), 0);
    }
    if (methodColIndexes['Наличка']) {
        setCellText(rowIndex, methodColIndexes['Наличка'], totalSums['Процессинг//Наличка'].toFixed(2), 0);
    }
    setCellText(rowIndex, colIndex, totalSums['Итого'].toFixed(2), 0); // Колонка "итого Процессинг"

    // Следующая строка для фиксированных костов в рублях
    rowIndex++;
    setCellText(rowIndex, 0, "₽ Итого: фикс косты на направление", 0);
    if (exchangeRate !== null) {
        if (methodColIndexes['Переводы']) {
            const totalInRublesПереводы = (totalSums['Процессинг//Переводы'] * exchangeRate).toFixed(2);
            setCellText(rowIndex, methodColIndexes['Переводы'], totalInRublesПереводы, 0);
        }
        if (methodColIndexes['Наличка']) {
            const totalInRublesНаличка = (totalSums['Процессинг//Наличка'] * exchangeRate).toFixed(2);
            setCellText(rowIndex, methodColIndexes['Наличка'], totalInRublesНаличка, 0);
        }
        const totalInRubles = (totalSums['Итого'] * exchangeRate).toFixed(2);
        setCellText(rowIndex, colIndex, totalInRubles, 0); // Колонка "итого Процессинг"
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }

    // Сохраняем строку с "₽ Итого: фикс косты на направление" для дальнейшего использования в формуле
    const rubleTotalRowIndex = rowIndex + 1;

    // Следующая строка для "₽ DB 2"
    rowIndex++;
    setCellText(rowIndex, 0, "₽ DB 2", 0);

    // Определяем букву колонки для текущей колонки
    const currentColLetter = String.fromCharCode(66 + colIndex - 1);

    // Формула для расчета "₽ DB 2" как разница между значениями в строке 26 и строкой с "₽ Итого: фикс косты на направление"
    const db2Formula = `=${currentColLetter}26 - ${currentColLetter}${rubleTotalRowIndex}`;
    setCellText(rowIndex, colIndex, db2Formula, 0); // Колонка "итого Процессинг" с формулой

    // Отступ двух строк
    rowIndex += 2;

    // Работа с проектами "Процессинг//" (но не 'Процессинг//Наличка' и не 'Процессинг//Переводы')
    const processingOperations = {};
    filteredRecordsDB2.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sum = parseFloat(record['Бухгалтерия_Сумма']);

        if (operation && !excludedOperations.has(operation) && project.startsWith('Процессинг//') && project !== 'Процессинг//Переводы' && project !== 'Процессинг//Наличка') {
            processingOperations[operation] = processingOperations[operation] || 0;
            if (!isNaN(sum)) {
                processingOperations[operation] += sum;
            }
        }
    });

    // Перечисляем операции для "Процессинг//"
    let processingTotalSum = 0;
    Object.keys(processingOperations).forEach(operation => {
        setCellText(rowIndex, 0, operation, 0);
        const operationSum = processingOperations[operation].toFixed(2);
        setCellText(rowIndex, colIndex, operationSum, 0); // Колонка "итого Процессинг"
        processingTotalSum += parseFloat(operationSum);
        rowIndex++;
    });

    // Отступ двух строк
    rowIndex += 2;

    // Итоговая строка для фиксированных костов в долларах для "Процессинг//"
    setCellText(rowIndex, 0, "$ Итого: фикс косты на направление", 0);
    setCellText(rowIndex, colIndex, processingTotalSum.toFixed(2), 0); // Колонка "итого Процессинг"

    // Следующая строка для фиксированных костов в рублях для "Процессинг//"
    rowIndex++;
    setCellText(rowIndex, 0, "₽ Итого: фикс косты на направление", 0);
    if (exchangeRate !== null) {
        const totalInRublesProcessing = (processingTotalSum * exchangeRate).toFixed(2);
        setCellText(rowIndex, colIndex, totalInRublesProcessing, 0); // Колонка "итого Процессинг"
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }

    // Следующая строка для "₽ DB 3"
    rowIndex++;
    setCellText(rowIndex, 0, "₽ DB 3", 0);

    // Формула для расчета "₽ DB 3" как разница между значениями в строке "₽ DB 2" и строкой с "₽ Итого: фикс косты на направление"
    const db3Formula = `=${currentColLetter}${rubleTotalRowIndex + 1} - ${currentColLetter}${rowIndex}`;
    setCellText(rowIndex, colIndex, db3Formula, 0); // Колонка "итого Процессинг" с формулой
}
