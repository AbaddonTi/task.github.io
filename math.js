/********** STYLE FUNCTIONS **********/
function setCellStyle(ri, ci, property, value, sheetIndex = 0) {
    const { styles, rows } = s.datas[sheetIndex];
    const cell = rows.getCellOrNew(ri, ci);
    let cstyle = cell.style !== undefined ? { ...styles[cell.style] } : {};
    if (property.startsWith('font')) {
        const [fontProp] = property.split('-').slice(1);
        cstyle.font = { ...(cstyle.font || {}), [fontProp]: value };
    } else if (property === 'border') {
        cstyle.border = { ...(cstyle.border || {}), ...value };
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
const thickBorderStyle = { style: 'thick', color: '#000000' }; // Толстая граница

const getBorderStyle = style => ({
    top: [style, defaultBorderStyle.color],
    bottom: [style, defaultBorderStyle.color],
    right: [style, defaultBorderStyle.color],
    left: [style, defaultBorderStyle.color]
});

function initializeSpreadsheetDB() {
    // Установка толстых вертикальных границ для первых двух ячеек
    [[0, 0], [0, 1]].forEach(([row, col]) => setCellBorder(row, col, getBorderStyle('thick')));

    // Установка заголовков "Статья" и "Процессинг"
    [[0, 0, 'Статья'], [0, 1, 'Процессинг']].forEach(([row, col, text]) => setCellText(row, col, text, 0));

    // Установка ширины колонок
    [[0, 150], [1, 150]].forEach(([col, width]) => setColumnWidth(col, width, 0));

    // Установка высоты первой строки
    setRowHeight(0, 30, 0);

    // Включение переноса текста для первых двух ячеек
    [[0, 0], [0, 1]].forEach(([row, col]) => enableTextWrap(row, col, 0));

    // Установка текста для различных строк
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
        const methodStartColIndex = colIndex;
        const methodEndColIndex = colIndex + banks.length;

        methodColIndexes[method] = { start: methodStartColIndex, end: methodEndColIndex };

        setCellText(2, methodStartColIndex, method, 0);
        setCellBorder(2, methodStartColIndex, getBorderStyle('thick'));
        for (let i = methodStartColIndex; i <= methodEndColIndex; i++) {
            setCellBorder(3, i, getBorderStyle('thin'));
        }
        for (let i = methodStartColIndex; i <= methodEndColIndex; i++) {
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
            setCellText(21, bankColIndex, '0', 0);
            setCellText(22, bankColIndex, `=(${colLetter}11-${colLetter}18)*${colLetter}22`, 0);
            setCellText(24, bankColIndex, `=${colLetter}14+${colLetter}18+${colLetter}21+${colLetter}23`, 0);
            setCellText(25, bankColIndex, `=${colLetter}11-${colLetter}25`, 0);
        });

        const totalColIndex = colIndex + banks.length;
        const totalColLetter = String.fromCharCode(66 + totalColIndex - 1);
        setCellText(2, totalColIndex, `Итого ${method}`, 0);
        setCellBorder(2, totalColIndex, getBorderStyle('thick'), 0);

        const bankColLetters = banks.map((_, index) => String.fromCharCode(66 + (colIndex + index) -1));

        setCellText(6, totalColIndex, `=AVERAGE(${bankColLetters.map(l => l + '7').join(',')})`, 0);
        setCellText(7, totalColIndex, `=SUM(${bankColLetters.map(l => l + '8').join(',')})`, 0);
        setCellText(8, totalColIndex, `=AVERAGE(${bankColLetters.map(l => l + '9').join(',')})`, 0);
        setCellText(10, totalColIndex, `=SUM(${bankColLetters.map(l => l + '11').join(',')})`, 0);
        setCellText(12, totalColIndex, `=${totalColLetter}15/${totalColLetter}8*100`, 0);
        setCellText(13, totalColIndex, `=SUM(${bankColLetters.map(l => l + '14').join(',')})`, 0);
        setCellText(14, totalColIndex, `=SUM(${bankColLetters.map(l => l + '15').join(',')})`, 0);
        setCellText(16, totalColIndex, `=${totalColLetter}19/${totalColLetter}8*100`, 0);
        setCellText(17, totalColIndex, `=SUM(${bankColLetters.map(l => l + '18').join(',')})`, 0);
        setCellText(18, totalColIndex, `=SUM(${bankColLetters.map(l => l + '19').join(',')})`, 0);
        setCellText(20, totalColIndex, `=SUM(${bankColLetters.map(l => l + '21').join(',')})`, 0);
        setCellText(21, totalColIndex, `=AVERAGE(${bankColLetters.map(l => l + '22').join(',')})`, 0);
        setCellText(22, totalColIndex, `=SUM(${bankColLetters.map(l => l + '23').join(',')})`, 0);
        setCellText(24, totalColIndex, `=SUM(${bankColLetters.map(l => l + '25').join(',')})`, 0);
        setCellText(25, totalColIndex, `=${totalColLetter}11-${totalColLetter}25`, 0);

        colIndex = totalColIndex + 1;
    });

    setCellText(0, colIndex, 'итого Процессинг', 0);
    setCellBorder(0, colIndex, getBorderStyle('thick'), 0);
    setCellBackgroundColor(0, colIndex, '#d9ead3', 0);

    const totalColLetter = String.fromCharCode(66 + colIndex - 1);

    setCellText(6, colIndex, averagePrice, 0);
    setCellText(7, colIndex, totalBanksCount, 0);
    setCellText(8, colIndex, averageProfit, 0);
    setCellText(10, colIndex, sumProfit, 0);
    setCellText(12, colIndex, `=${totalColLetter}15/${totalColLetter}8*100`, 0);
    setCellText(13, colIndex, `=${totalColLetter}7*${totalColLetter}15`, 0);
    setCellText(14, colIndex, totalProblemCount, 0);
    setCellText(16, colIndex, `=${totalColLetter}19/${totalColLetter}8*100`, 0);
    setCellText(17, colIndex, `=${totalColLetter}7*${totalColLetter}19`, 0);
    setCellText(18, colIndex, totalBlockCount, 0);
    setCellText(20, colIndex, `=${totalColLetter}7*${totalColLetter}8`, 0);
    setCellText(21, colIndex, '0', 0);
    setCellText(22, colIndex, `=(${totalColLetter}11-${totalColLetter}18)*${totalColLetter}22`, 0);
    setCellText(24, colIndex, `=${totalColLetter}14+${totalColLetter}18+${totalColLetter}21+${totalColLetter}23`, 0);
    setCellText(25, colIndex, `=${totalColLetter}11-${totalColLetter}25`, 0);

    for (let i = 1; i <= colIndex; i++) {
        setCellBackgroundColor(0, i, '#d9ead3', 0);
    }

    return { colIndex, methodColIndexes };
}

function getExchangeRate(records) {
    for (let record of records) {
        if (record['Бухгалтерия_Курс_Usd_Rub']) {
            return parseFloat(record['Бухгалтерия_Курс_Usd_Rub']);
        }
    }
    return null;
}

function setRightBorders(colIndexes, startRow, endRow, sheetIndex = 0, isThick = false) {
    const borderStyle = isThick ? getBorderStyle('thick') : getBorderStyle('thin');
    colIndexes.forEach(colIndex => {
        for (let ri = startRow; ri <= endRow; ri++) {
            setCellBorder(ri, colIndex, {
                right: [borderStyle.right[0], borderStyle.right[1]]
            }, sheetIndex);
        }
    });
}

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
    const exchangeRate = getExchangeRate(mappedRecords); // Предполагается, что mappedRecords доступен

    let rowIndex = 28;
    const totalSums = {
        'Процессинг//Переводы': 0,
        'Процессинг//Наличка': 0,
        'Итого': 0
    };
    const cashInTotalSums = {
        'Cash-In//': 0,
        'Итого': 0
    };

    const projectsToProcess = ['Процессинг//Переводы', 'Процессинг//Наличка'];
    const cashInProject = 'Cash-In//';

    if (!methodColIndexes['Cash-In//']) {
        methodColIndexes['Cash-In//'] = { start: colIndex + 1, end: colIndex + 1 };
    }

    const cashInColIndex = methodColIndexes['Cash-In//'].end;
    setCellText(0, cashInColIndex, "Cash-In//", 0);

    let cashInRecountSum = 0;

    filteredRecordsDB2.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sum = parseFloat(record['Бухгалтерия_Сумма']);

        if (operation === "Пересчёт кассы" && project === cashInProject) {
            if (!isNaN(sum)) {
                cashInRecountSum += sum;
            }
        }
    });

    setCellText(10, cashInColIndex, cashInRecountSum.toFixed(2), 0);

    // === Основной цикл обработки операций для проектов процессинга и Cash-In// ===
    filteredRecordsDB2.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sum = parseFloat(record['Бухгалтерия_Сумма']);

        if (
            operation &&
            !excludedOperations.has(operation) &&
            projectsToProcess.includes(project)
        ) {
            uniqueOperations.add(operation);
            if (!isNaN(sum)) {
                if (!operationSums[operation]) {
                    operationSums[operation] = {
                        'Процессинг//Переводы': 0,
                        'Процессинг//Наличка': 0
                    };
                }
                operationSums[operation][project] += sum;
            }
        }
        else if (
            operation &&
            !excludedOperations.has(operation) &&
            operation !== "Пересчёт кассы" &&
            project === cashInProject
        ) {
            uniqueOperations.add(operation);
            if (!isNaN(sum)) {
                if (!operationSums[operation]) {
                    operationSums[operation] = {
                        'Cash-In//': 0
                    };
                }
                operationSums[operation]['Cash-In//'] = (operationSums[operation]['Cash-In//'] || 0) + sum;
            }
        }
    });

    uniqueOperations.forEach(operation => {
        if (operation === "Пересчёт кассы") {
            return;
        }

        setCellText(rowIndex, 0, operation, 0);

        const переводSum = operationSums[operation]['Процессинг//Переводы'] || 0;
        const наличкаSum = operationSums[operation]['Процессинг//Наличка'] || 0;
        const cashInSum = operationSums[operation]['Cash-In//'] || 0;

        const totalProcessingSum = (переводSum + наличкаSum).toFixed(2);

        setCellText(rowIndex, colIndex, totalProcessingSum, 0);

        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            setCellText(rowIndex, methodTotalColIndex, переводSum.toFixed(2), 0);
            totalSums['Процессинг//Переводы'] += переводSum;
        }
        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            setCellText(rowIndex, methodTotalColIndex, наличкаSum.toFixed(2), 0);
            totalSums['Процессинг//Наличка'] += наличкаSum;
        }

        totalSums['Итого'] += parseFloat(totalProcessingSum);

        if (cashInSum !== 0) {
            setCellText(rowIndex, cashInColIndex, cashInSum.toFixed(2), 0);
            cashInTotalSums['Cash-In//'] += cashInSum;
            cashInTotalSums['Итого'] += cashInSum;
        }

        rowIndex++;
    });

    rowIndex++;

    setCellText(rowIndex, 0, "Текущий курс USD/RUB:", 0);
    if (exchangeRate !== null) {
        setCellText(rowIndex, 1, exchangeRate.toFixed(2), 0);
    }

    const exchangeRateRowIndex = rowIndex + 1;

    rowIndex++;

    const cashInColLetter = String.fromCharCode(65 + cashInColIndex);
    setCellText(25, cashInColIndex, `=${cashInColLetter}11*B${exchangeRateRowIndex}`, 0);

    const totalTotalsColIndex = cashInColIndex + 1;
    setCellText(0, totalTotalsColIndex, "Итоги Итогов:", 0);

    const totalProcessingColLetter = String.fromCharCode(65 + colIndex);
    setCellText(25, totalTotalsColIndex, `=${totalProcessingColLetter}26+${cashInColLetter}26`, 0);

    // === Итоговые строки для DB 2 ===

    setCellText(rowIndex, 0, "$ Итого: фикс косты на поднаправление", 0);
    if (methodColIndexes['Переводы']) {
        const methodTotalColIndex = methodColIndexes['Переводы'].end;
        setCellText(rowIndex, methodTotalColIndex, totalSums['Процессинг//Переводы'].toFixed(2), 0);
    }
    if (methodColIndexes['Наличка']) {
        const methodTotalColIndex = methodColIndexes['Наличка'].end;
        setCellText(rowIndex, methodTotalColIndex, totalSums['Процессинг//Наличка'].toFixed(2), 0);
    }
    setCellText(rowIndex, colIndex, totalSums['Итого'].toFixed(2), 0);

    rowIndex++;
    setCellText(rowIndex, 0, "₽ Итого: фикс косты на поднаправление", 0);
    if (exchangeRate !== null) {
        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            const totalInRublesПереводы = (totalSums['Процессинг//Переводы'] * exchangeRate).toFixed(2);
            setCellText(rowIndex, methodTotalColIndex, totalInRublesПереводы, 0);
        }
        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            const totalInRublesНаличka = (totalSums['Процессинг//Наличка'] * exchangeRate).toFixed(2);
            setCellText(rowIndex, methodTotalColIndex, totalInRublesНаличka, 0);
        }
        const totalInRubles = (totalSums['Итого'] * exchangeRate).toFixed(2);
        setCellText(rowIndex, colIndex, totalInRubles, 0);
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }

    const rubleSubTotalRowIndex = rowIndex + 1;

    rowIndex++;
    setCellText(rowIndex, 0, "₽ DB 2", 0);

    const totalColLetter = String.fromCharCode(65 + colIndex);
    const db2Formula = `=${totalColLetter}26 - ${totalColLetter}${rubleSubTotalRowIndex}`;
    setCellText(rowIndex, colIndex, db2Formula, 0);

    // === Отступ двух строк и переходим к DB 3 ===
    rowIndex += 2;
    const processingOperations = {};
    filteredRecordsDB2.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sum = parseFloat(record['Бухгалтерия_Сумма']);

        if (
            operation &&
            !excludedOperations.has(operation) &&
            project.startsWith('Процессинг//') &&
            !projectsToProcess.includes(project)
        ) {
            if (!processingOperations[operation]) {
                processingOperations[operation] = 0;
            }
            if (!isNaN(sum)) {
                processingOperations[operation] += sum;
            }
        }
    });

    let processingTotalSum = 0;
    Object.keys(processingOperations).forEach(operation => {
        setCellText(rowIndex, 0, operation, 0);
        const operationSum = processingOperations[operation].toFixed(2);
        setCellText(rowIndex, colIndex, operationSum, 0);
        processingTotalSum += parseFloat(operationSum);
        rowIndex++;
    });

    rowIndex += 2;

    // === Итоговые строки для DB 3 ===

    setCellText(rowIndex, 0, "$ Итого: фикс косты на направление", 0);
    const totalProcessingAndCashInSum = processingTotalSum + cashInTotalSums['Cash-In//'];
    setCellText(rowIndex, colIndex, totalProcessingAndCashInSum.toFixed(2), 0);

    if (methodColIndexes['Cash-In//']) {
        setCellText(rowIndex, cashInColIndex, cashInTotalSums['Cash-In//'].toFixed(2), 0);
    }

    setCellText(rowIndex, totalTotalsColIndex, totalProcessingAndCashInSum.toFixed(2), 0);

    rowIndex++;
    setCellText(rowIndex, 0, "₽ Итого: фикс косты на направление", 0);
    if (exchangeRate !== null) {
        const totalInRublesProcessing = (totalProcessingAndCashInSum * exchangeRate).toFixed(2);
        setCellText(rowIndex, colIndex, totalInRublesProcessing, 0);

        if (methodColIndexes['Cash-In//']) {
            const totalInRublesCashIn = (cashInTotalSums['Cash-In//'] * exchangeRate).toFixed(2);
            setCellText(rowIndex, cashInColIndex, totalInRublesCashIn, 0);
            const totalInRublesTotal = parseFloat(totalInRublesProcessing) + parseFloat(totalInRublesCashIn);
            setCellText(rowIndex, totalTotalsColIndex, totalInRublesTotal.toFixed(2), 0);
        }
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }

    const rubleTotalRowIndex = rowIndex + 1;

    rowIndex++;
    setCellText(rowIndex, 0, "₽ DB 3", 0);

    if (exchangeRate !== null) {
        const formulaCashInDB3 = `=${cashInColLetter}26  - ${cashInColLetter}${rubleTotalRowIndex}`;
        setCellText(rowIndex, cashInColIndex, formulaCashInDB3, 0);
    }

    const db3Formula = `=${totalColLetter}${rubleSubTotalRowIndex + 1} - ${totalColLetter}${rubleTotalRowIndex}`;
    setCellText(rowIndex, colIndex, db3Formula, 0);

    const cashInDB3ColLetter = cashInColLetter;
    setCellText(rowIndex, totalTotalsColIndex, `=${totalColLetter}${rowIndex + 1}+${cashInDB3ColLetter}${rowIndex + 1}`, 0);

    // === Добавление Правых Границ для Методов ===
    const borderStartRow = 0; // Начальная строка, с которой начинаются методы
    const borderEndRow = rowIndex; // Последняя строка после рендеринга

    Object.values(methodColIndexes).forEach(({ end: lastColIndex }) => {
        for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
            setCellBorder(ri, lastColIndex, {
                right: [defaultBorderStyle.style, defaultBorderStyle.color]
            }, 0);
        }
    });

    // === Добавление Правых Границ для Дополнительных Колонок ===

    // Определяем список колонок, для которых нужно установить правые границы
    const additionalBordersColIndexes = [
        0, // Первый столбец ("Статья")
        colIndex, // Колонка "итого Процессинг"
        cashInColIndex, // Колонка "Cash-In//"
        totalTotalsColIndex // Колонка "Итоги Итогов:"
    ];

    // Устанавливаем толстые границы для указанных колонок
    additionalBordersColIndexes.forEach(targetColIndex => {
        for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
            // Определяем, какую границу использовать: 'thick' или 'thin'
            let borderStyle;
            if (
                targetColIndex === 0 || // "Статья"
                targetColIndex === colIndex || // "итого Процессинг"
                targetColIndex === cashInColIndex || // "Cash-In//"
                targetColIndex === totalTotalsColIndex // "Итоги Итогов:"
            ) {
                borderStyle = getBorderStyle('thick'); // Толстая граница
            } else {
                borderStyle = getBorderStyle('thin'); // Тонкая граница (на всякий случай)
            }

            setCellBorder(ri, targetColIndex, {
                right: [borderStyle.right[0], borderStyle.right[1]]
            }, 0);
        }
    });

    // === Добавление Тонких Горизонтальных Линий Под Каждым DB ===
    // Точное отслеживание строк DB1, DB2 и DB3
    const db1Row = 25;
    const db2Row = rubleSubTotalRowIndex; // Строка под "₽ DB 2"
    const db3Row = rowIndex; // Строка под "₽ DB 3"

    const dbRows = [db1Row, db2Row, db3Row];

    dbRows.forEach(dbRow => {
        for (let col = 0; col <= totalTotalsColIndex; col++) {
            setCellBorder(dbRow, col, {
                bottom: ['thin', '#000000']
            }, 0);
        }
    });

    // === Добавление Тонкой Горизонтальной Линии Под Первой Строкой (Заголовки) ===
    for (let col = 0; col <= totalTotalsColIndex; col++) {
        setCellBorder(0, col, {
            bottom: ['thin', '#000000']
        }, 0);
    }

    // === Завершение Функции ===
}
