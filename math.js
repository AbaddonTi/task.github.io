/********** STYLE FUNCTIONS **********/
function setCellStyle(ri, ci, property, value, sheetIndex = 0) {
    const { styles, rows } = s.datas[sheetIndex];
    const cell = rows.getCellOrNew(ri, ci);
    let cstyle = cell.style !== undefined ? { ...styles[cell.style] } : {};

    if (property === 'format') {
        cstyle.format = value;
    } else if (property.startsWith('font')) {
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
const thickBorderStyle = { style: 'thick', color: '#000000' }; 

const getBorderStyle = style => ({
    top: [style, defaultBorderStyle.color],
    bottom: [style, defaultBorderStyle.color],
    right: [style, defaultBorderStyle.color],
    left: [style, defaultBorderStyle.color]
});

function initializeSpreadsheetDB() {
    [[0, 0], [0, 1]].forEach(([row, col]) => setCellBorder(row, col, getBorderStyle('thick')));
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
    const { colIndex, methodColIndexes } = renderBankData(bankData); 
    renderOperationTypes(filteredRecordsDB2, { colIndex, methodColIndexes }, exchangeRate); 
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

    const formatCell = (row, col, value, formatType) => {
        setCellText(row, col, value, 0);
        if (formatType) {
            setCellStyle(row, col, 'format', formatType);
        }
    };

    const applyBorder = (row, col, style) => setCellBorder(row, col, getBorderStyle(style));

    const applyBackground = (row, col, color) => setCellBackgroundColor(row, col, color, 0);

    const calculateAverage = (sum, count) => (count > 0 ? (sum / count).toFixed(2) : '0.00');

    const averagePrice = calculateAverage(totalPrices, totalPricesCount);
    const averageProfit = calculateAverage(totalProfits, totalProfitsCount);
    const sumProfit = totalProfits.toFixed(2);

    let colIndex = 1;
    const methodColIndexes = {};

    Object.keys(bankMethodMap).forEach(method => {
        const banks = Array.from(bankMethodMap[method]);
        const methodStartCol = colIndex;
        const methodEndCol = colIndex + banks.length - 1;

        methodColIndexes[method] = { start: methodStartCol, end: methodEndCol };

        formatCell(2, methodStartCol, method);
        applyBorder(2, methodStartCol, 'thick');
        for (let i = methodStartCol; i <= methodEndCol; i++) {
            applyBorder(3, i, 'thin');
            applyBackground(0, i, '#d9ead3');
        }

        banks.forEach((bank, index) => {
            const bankCol = colIndex + index;
            const colLetter = String.fromCharCode(66 + bankCol - 1); 

            formatCell(3, bankCol, bank);

            const prices = bankPricesMap[method][bank] || [];
            if (prices.length) {
                const avgPrice = calculateAverage(prices.reduce((a, b) => a + b, 0), prices.length);
                formatCell(6, bankCol, avgPrice, 'rub');
            } else {
                formatCell(6, bankCol, '0.00', 'rub');
            }

            formatCell(7, bankCol, bankCountsMap[method][bank]);
            const profits = bankProfitsMap[method][bank] || [];
            if (profits.length) {
                const avgProfit = calculateAverage(profits.reduce((a, b) => a + b, 0), profits.length);
                formatCell(8, bankCol, avgProfit, 'rub');
            } else {
                formatCell(8, bankCol, '0.00', 'rub');
            }

            const formulaMap = {
                10: { formula: `=${colLetter}8*${colLetter}9`, format: 'rub' },
                12: { formula: `=${colLetter}15/${colLetter}8*100`, format: 'percent' },
                13: { formula: `=${colLetter}7*${colLetter}15`, format: 'rub' },
                14: { value: bankProblemCountsMap[method]?.[bank] || 0 },
                16: { formula: `=${colLetter}19/${colLetter}8*100`, format: 'percent' },
                17: { formula: `=${colLetter}7*${colLetter}19`, format: 'rub' },
                18: { value: bankBlockCountsMap[method]?.[bank] || 0 },
                20: { formula: `=${colLetter}7*${colLetter}8`, format: 'rub' },
                21: { value: '0', format: 'percent' },
                22: { formula: `=(${colLetter}11-${colLetter}18)*${colLetter}22`, format: 'rub' },
                24: { formula: `=${colLetter}14+${colLetter}18+${colLetter}21+${colLetter}23`, format: 'rub' },
                25: { formula: `=${colLetter}11-${colLetter}25`, format: 'rub' }
            };

            Object.entries(formulaMap).forEach(([row, data]) => {
                if (data.formula) {
                    formatCell(Number(row), bankCol, data.formula, data.format);
                } else if (data.value !== undefined) {
                    formatCell(Number(row), bankCol, data.value, data.format);
                }
            });
        });

        const totalCol = colIndex + banks.length;
        const totalColLetter = String.fromCharCode(66 + totalCol - 1);

        formatCell(2, totalCol, `Итого ${method}`);
        applyBorder(2, totalCol, 'thick');

        const bankColLetters = banks.map((_, idx) => String.fromCharCode(66 + colIndex + idx - 1));

        const totalFormulaMap = {
            6: { formula: `=AVERAGE(${bankColLetters.map(l => `${l}7`).join(',')})`, format: 'rub' },
            7: { formula: `=SUM(${bankColLetters.map(l => `${l}8`).join(',')})` },
            8: { formula: `=AVERAGE(${bankColLetters.map(l => `${l}9`).join(',')})`, format: 'rub' },
            10: { formula: `=SUM(${bankColLetters.map(l => `${l}11`).join(',')})`, format: 'rub' },
            12: { formula: `=${totalColLetter}15/${totalColLetter}8*100`, format: 'percent' },
            13: { formula: `=SUM(${bankColLetters.map(l => `${l}14`).join(',')})`, format: 'rub' },
            14: { formula: `=SUM(${bankColLetters.map(l => `${l}15`).join(',')})` },
            16: { formula: `=${totalColLetter}19/${totalColLetter}8*100`, format: 'percent' },
            17: { formula: `=SUM(${bankColLetters.map(l => `${l}18`).join(',')})`, format: 'rub' },
            18: { formula: `=SUM(${bankColLetters.map(l => `${l}19`).join(',')})` },
            20: { formula: `=SUM(${bankColLetters.map(l => `${l}21`).join(',')})`, format: 'rub' },
            21: { formula: `=AVERAGE(${bankColLetters.map(l => `${l}22`).join(',')})`, format: 'percent' },
            22: { formula: `=SUM(${bankColLetters.map(l => `${l}23`).join(',')})`, format: 'rub' },
            24: { formula: `=SUM(${bankColLetters.map(l => `${l}25`).join(',')})`, format: 'rub' },
            25: { formula: `=${totalColLetter}11-${totalColLetter}25`, format: 'rub' }
        };

        Object.entries(totalFormulaMap).forEach(([row, data]) => {
            if (data.formula) {
                formatCell(Number(row), totalCol, data.formula, data.format);
            } else if (data.value !== undefined) {
                formatCell(Number(row), totalCol, data.value, data.format);
            }
        });

        colIndex = totalCol + 1;
    });

    formatCell(0, colIndex, 'итого Процессинг');
    applyBorder(0, colIndex, 'thick');
    applyBackground(0, colIndex, '#d9ead3');

    const totalProcessColLetter = String.fromCharCode(66 + colIndex - 1);

    const processFormulaMap = {
        6: { value: averagePrice, format: 'rub' },
        7: { value: totalBanksCount },
        8: { value: averageProfit, format: 'rub' },
        10: { value: sumProfit, format: 'rub' },
        12: { formula: `=${totalProcessColLetter}15/${totalProcessColLetter}8*100`, format: 'percent' },
        13: { formula: `=${totalProcessColLetter}7*${totalProcessColLetter}15`, format: 'rub' },
        14: { value: totalProblemCount },
        16: { formula: `=${totalProcessColLetter}19/${totalProcessColLetter}8*100`, format: 'percent' },
        17: { formula: `=${totalProcessColLetter}7*${totalProcessColLetter}19`, format: 'rub' },
        18: { value: totalBlockCount },
        20: { formula: `=${totalProcessColLetter}7*${totalProcessColLetter}8`, format: 'rub' },
        21: { value: '0', format: 'percent' },
        22: { formula: `=(${totalProcessColLetter}11-${totalProcessColLetter}18)*${totalProcessColLetter}22`, format: 'rub' },
        24: { formula: `=${totalProcessColLetter}14+${totalProcessColLetter}18+${totalProcessColLetter}21+${totalProcessColLetter}23`, format: 'rub' },
        25: { formula: `=${totalProcessColLetter}11-${totalProcessColLetter}25`, format: 'rub' }
    };

    Object.entries(processFormulaMap).forEach(([row, data]) => {
        if (data.formula) {
            formatCell(Number(row), colIndex, data.formula, data.format);
        } else if (data.value !== undefined) {
            formatCell(Number(row), colIndex, data.value, data.format);
        }
    });

    for (let i = 1; i <= colIndex; i++) {
        applyBackground(0, i, '#d9ead3');
    }

    return { colIndex, methodColIndexes };
}



function renderOperationTypes(filteredRecordsDB2, { colIndex, methodColIndexes }, exchangeRate) {
    const excludedOperations = new Set([
        "Пересчёт кассы",
        "Внесение Слава",
        "Внесение Patek",
        "Внесение Роман",
        "Вывод Роман",
        "Вывод Слава",
        "Вывод Patek",
        "Выплаты Consulting",
        "Доход от рефералов",
        "Расходы на дропов",
        "Апелляция",
        "Перевод (получаем на счёт)",
        "Перевод (снимаем с счёта)"
    ]);
    const uniqueOperations = new Set();
    const operationSums = {};

    let currentRowIndex = 28; // Начальная строка для DB2
    const totalSumsDB2 = {
        'Процессинг//Переводы': 0,
        'Процессинг//Наличка': 0,
        'Итого': 0
    };
    const cashInTotalSumsDB2 = {
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
    setCellStyle(10, cashInColIndex, 'format', 'usd'); // Формат 'usd'

    // === Обработка операций для DB2 ===
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

        setCellText(currentRowIndex, 0, operation, 0);

        const переводSum = operationSums[operation]['Процессинг//Переводы'] || 0;
        const наличкаSum = operationSums[operation]['Процессинг//Наличка'] || 0;
        const cashInSum = operationSums[operation]['Cash-In//'] || 0;

        const totalProcessingSum = (переводSum + наличкаSum).toFixed(2);

        setCellText(currentRowIndex, colIndex, totalProcessingSum, 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'usd');

        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            setCellText(currentRowIndex, methodTotalColIndex, переводSum.toFixed(2), 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'usd'); 
            totalSumsDB2['Процессинг//Переводы'] += переводSum;
        }
        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            setCellText(currentRowIndex, methodTotalColIndex, наличкаSum.toFixed(2), 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'usd'); 
            totalSumsDB2['Процессинг//Наличка'] += наличкаSum;
        }

        totalSumsDB2['Итого'] += parseFloat(totalProcessingSum);

        if (cashInSum !== 0) {
            setCellText(currentRowIndex, cashInColIndex, cashInSum.toFixed(2), 0);
            setCellStyle(currentRowIndex, cashInColIndex, 'format', 'usd'); 
            cashInTotalSumsDB2['Cash-In//'] += cashInSum;
            cashInTotalSumsDB2['Итого'] += cashInSum;
        }
        currentRowIndex++;
    });

    currentRowIndex++;

    // Установка текущего курса USD/RUB
    setCellText(currentRowIndex, 0, "Текущий курс USD/RUB:", 0);

    if (exchangeRate !== null) {
        setCellText(currentRowIndex, 1, exchangeRate.toFixed(2), 0);
        setCellStyle(currentRowIndex, 1, 'format', 'rub');
    }

    const exchangeRateRowIndexDB2 = currentRowIndex + 1;

    currentRowIndex++;

    // Формула для Cash-In// в DB2
    const cashInColLetter = String.fromCharCode(65 + cashInColIndex);
    setCellText(25, cashInColIndex, `=${cashInColLetter}11*B${exchangeRateRowIndexDB2}`, 0);
    setCellStyle(25, cashInColIndex, 'format', 'rub');

    const totalTotalsColIndex = cashInColIndex + 1;
    setCellText(0, totalTotalsColIndex, "Итоги Итогов:", 0);

    const totalProcessingColLetter = String.fromCharCode(65 + colIndex);
    setCellText(25, totalTotalsColIndex, `=${totalProcessingColLetter}26+${cashInColLetter}26`, 0);
    setCellStyle(25, totalTotalsColIndex, 'format', 'rub'); 

    // === Итоговые строки для DB2 ===

    setCellText(currentRowIndex, 0, "$ Итого: фикс косты на поднаправление", 0);
    setCellStyle(currentRowIndex, 0, 'format', ''); 

    if (methodColIndexes['Переводы']) {
        const methodTotalColIndex = methodColIndexes['Переводы'].end;
        setCellText(currentRowIndex, methodTotalColIndex, totalSumsDB2['Процессинг//Переводы'].toFixed(2), 0);
        setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'usd');
    }
    if (methodColIndexes['Наличка']) {
        const methodTotalColIndex = methodColIndexes['Наличка'].end;
        setCellText(currentRowIndex, methodTotalColIndex, totalSumsDB2['Процессинг//Наличка'].toFixed(2), 0);
        setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'usd');
    }
    setCellText(currentRowIndex, colIndex, (totalSumsDB2['Итого']).toFixed(2), 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'usd'); 

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на поднаправление", 0);
    setCellStyle(currentRowIndex, 0, 'format', ''); 

    if (exchangeRate !== null) {
        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            const totalInRublesПереводы = (totalSumsDB2['Процессинг//Переводы'] * exchangeRate).toFixed(2);
            setCellText(currentRowIndex, methodTotalColIndex, totalInRublesПереводы, 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
        }
        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            const totalInRublesНаличka = (totalSumsDB2['Процессинг//Наличка'] * exchangeRate).toFixed(2);
            setCellText(currentRowIndex, methodTotalColIndex, totalInRublesНаличka, 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
        }
        const totalInRubles = (totalSumsDB2['Итого'] + cashInTotalSumsDB2['Итого']) * exchangeRate;
        setCellText(currentRowIndex, colIndex, totalInRubles.toFixed(2), 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'rub');
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }

    const rubleSubTotalRowIndexDB2 = currentRowIndex + 1;

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "₽ DB 2", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    const totalColLetterDB2 = String.fromCharCode(65 + colIndex);
    const db2Formula = `=${totalColLetterDB2}26 - ${totalColLetterDB2}${rubleSubTotalRowIndexDB2}`;
    setCellText(currentRowIndex, colIndex, db2Formula, 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

    // === Переход к DB3 ===
    currentRowIndex += 2; 

    const db3StartRowIndex = currentRowIndex;
    const processingOperationsDB3 = {};
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
            if (!processingOperationsDB3[operation]) {
                processingOperationsDB3[operation] = 0;
            }
            if (!isNaN(sum)) {
                processingOperationsDB3[operation] += sum;
            }
        }
    });

    let processingTotalSumDB3 = 0;
    Object.keys(processingOperationsDB3).forEach(operation => {
        setCellText(currentRowIndex, 0, operation, 0);

        const operationSum = processingOperationsDB3[operation].toFixed(2);
        setCellText(currentRowIndex, colIndex, operationSum, 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'usd');
        processingTotalSumDB3 += parseFloat(operationSum);
        currentRowIndex++;
    });

    currentRowIndex += 2; 

    // === Итоговые строки для DB3 ===

    setCellText(currentRowIndex, 0, "$ Итого: фикс косты на направление", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    const totalProcessingSumOnlyDB3 = processingTotalSumDB3;
    setCellText(currentRowIndex, colIndex, totalProcessingSumOnlyDB3.toFixed(2), 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'usd');

    let cashInSumDB3 = 0;
    if (methodColIndexes['Cash-In//']) {
        cashInSumDB3 = cashInTotalSumsDB2['Cash-In//'];
        setCellText(currentRowIndex, cashInColIndex, cashInSumDB3.toFixed(2), 0);
        setCellStyle(currentRowIndex, cashInColIndex, 'format', 'usd');
    }

    const totalTotalsSumDB3 = totalProcessingSumOnlyDB3 + cashInSumDB3;
    setCellText(currentRowIndex, totalTotalsColIndex, totalTotalsSumDB3.toFixed(2), 0);
    setCellStyle(currentRowIndex, totalTotalsColIndex, 'format', 'usd');

    currentRowIndex++;

    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на направление", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    if (exchangeRate !== null) {
        const totalInRublesProcessingDB3 = (totalProcessingSumOnlyDB3 * exchangeRate).toFixed(2);
        const totalInRublesCashInDB3 = (cashInTotalSumsDB2['Cash-In//'] * exchangeRate).toFixed(2);
        const totalInRublesTotalDB3 = (totalTotalsSumDB3) * exchangeRate;

        setCellText(currentRowIndex, colIndex, totalInRublesProcessingDB3, 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

        if (methodColIndexes['Cash-In//']) {
            setCellText(currentRowIndex, cashInColIndex, totalInRublesCashInDB3, 0);
            setCellStyle(currentRowIndex, cashInColIndex, 'format', 'rub');
        }

        setCellText(currentRowIndex, totalTotalsColIndex, totalInRublesTotalDB3.toFixed(2), 0);
        setCellStyle(currentRowIndex, totalTotalsColIndex, 'format', 'rub');
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }

    const rubleTotalRowIndexDB3 = currentRowIndex + 1;

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "₽ DB 3", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    if (exchangeRate !== null) {
        // Формула для Cash-In// в DB3
        const formulaCashInDB3 = `=${cashInColLetter}26 - ${cashInColLetter}${rubleTotalRowIndexDB3}`;
        setCellText(currentRowIndex, cashInColIndex, formulaCashInDB3, 0);
        setCellStyle(currentRowIndex, cashInColIndex, 'format', 'rub');
    }

    const db3Formula = `=${totalColLetterDB2}${rubleSubTotalRowIndexDB2 + 1} - ${totalColLetterDB2}${rubleTotalRowIndexDB3}`;
    setCellText(currentRowIndex, colIndex, db3Formula, 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

    const cashInDB3ColLetter = cashInColLetter;
    setCellText(currentRowIndex, totalTotalsColIndex, `=${totalColLetterDB2}${currentRowIndex + 1}+${cashInDB3ColLetter}${currentRowIndex + 1}`, 0);
    setCellStyle(currentRowIndex, totalTotalsColIndex, 'format', 'rub');

    // === Добавление Правых Границ для Методов ===
    const borderStartRow = 0;
    const borderEndRow = currentRowIndex;

    Object.values(methodColIndexes).forEach(({ end: lastColIndex }) => {
        for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
            setCellBorder(ri, lastColIndex, {
                right: [defaultBorderStyle.style, defaultBorderStyle.color]
            }, 0);
        }
    });

    // === Добавление Правых Границ для Дополнительных Колонок ===

    const additionalBordersColIndexes = [
        0, 
        colIndex, 
        cashInColIndex, 
        totalTotalsColIndex 
    ];

    additionalBordersColIndexes.forEach(targetColIndex => {
        for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
            let borderStyle;
            if (
                targetColIndex === 0 || 
                targetColIndex === colIndex || 
                targetColIndex === cashInColIndex || 
                targetColIndex === totalTotalsColIndex 
            ) {
                borderStyle = getBorderStyle('thick'); 
            } else {
                borderStyle = getBorderStyle('thin'); 
            }

            setCellBorder(ri, targetColIndex, {
                right: [borderStyle.right[0], borderStyle.right[1]]
            }, 0);
        }
    });

    // === Добавление Тонких Горизонтальных Линий Под Каждым DB ===

    const db1Row = 25;
    const db2Row = rubleSubTotalRowIndexDB2;
    const db3Row = rubleTotalRowIndexDB3; 

    const dbRows = [db1Row, db2Row, db3Row];

    dbRows.forEach(dbRow => {
        for (let col = 0; col <= totalTotalsColIndex; col++) {
            setCellBorder(dbRow, col, {
                bottom: ['thin', '#000000']
            }, 0);
        }
    });
    for (let col = 0; col <= totalTotalsColIndex; col++) {
        setCellBorder(0, col, {
            bottom: ['thin', '#000000']
        }, 0);
    }
}
