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

        const isProblem = status.includes('проблема');

        if (bank && method) {
            bankMethodMap[method] = bankMethodMap[method] || new Set();
            bankMethodMap[method].add(bank);

            bankPricesMap[method] = bankPricesMap[method] || {};
            bankPricesMap[method][bank] = bankPricesMap[method][bank] || [];

            bankCountsMap[method] = bankCountsMap[method] || {};
            bankCountsMap[method][bank] = (bankCountsMap[method][bank] || 0);

            bankProfitsMap[method] = bankProfitsMap[method] || {};
            bankProfitsMap[method][bank] = bankProfitsMap[method][bank] || [];

            if (!isProblem) {
                if (!isNaN(price)) {
                    bankPricesMap[method][bank].push(price);
                    totalPrices += price;
                    totalPricesCount += 1;
                }

                bankCountsMap[method][bank] += 1;
                totalBanksCount += 1;

                if (!isNaN(profit)) {
                    bankProfitsMap[method][bank].push(profit);
                    totalProfits += profit;
                    totalProfitsCount += 1;
                }
            }

            if (isProblem) {
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
                setCellStyle(6, bankColIndex, 'format', 'rub'); 
            }

            setCellText(7, bankColIndex, bankCountsMap[method][bank], 0);

            const profits = bankProfitsMap[method][bank];
            if (profits && profits.length) {
                const averageProfit = (profits.reduce((a, b) => a + b, 0) / profits.length).toFixed(2);
                setCellText(8, bankColIndex, averageProfit, 0);
                setCellStyle(8, bankColIndex, 'format', 'rub'); 
            }

            setCellText(10, bankColIndex, `=${colLetter}8*${colLetter}9`, 0);
            setCellStyle(10, bankColIndex, 'format', 'rub'); 

            setCellText(12, bankColIndex, `=${colLetter}15/${colLetter}8*100`, 0);
            setCellStyle(12, bankColIndex, 'format', 'percent'); 

            setCellText(13, bankColIndex, `=${colLetter}7*${colLetter}15`, 0);
            setCellStyle(13, bankColIndex, 'format', 'rub'); 

            setCellText(14, bankColIndex, bankProblemCountsMap[method]?.[bank] || 0, 0);

            setCellText(16, bankColIndex, `=${colLetter}19/${colLetter}8*100`, 0);
            setCellStyle(16, bankColIndex, 'format', 'percent'); 

            setCellText(17, bankColIndex, `=${colLetter}7*${colLetter}19`, 0);
            setCellStyle(17, bankColIndex, 'format', 'rub'); 

            setCellText(18, bankColIndex, bankBlockCountsMap[method]?.[bank] || 0, 0);

            setCellText(20, bankColIndex, `=${colLetter}7*${colLetter}8`, 0);
            setCellStyle(20, bankColIndex, 'format', 'rub'); 

            setCellText(21, bankColIndex, '0', 0);
            setCellStyle(21, bankColIndex, 'format', 'percent'); 

            setCellText(22, bankColIndex, `=(${colLetter}11-${colLetter}18)*${colLetter}22`, 0);
            setCellStyle(22, bankColIndex, 'format', 'rub'); 

            setCellText(24, bankColIndex, `=${colLetter}14+${colLetter}18+${colLetter}21+${colLetter}23`, 0);
            setCellStyle(24, bankColIndex, 'format', 'rub'); 

            setCellText(25, bankColIndex, `=${colLetter}11-${colLetter}25`, 0);
            setCellStyle(25, bankColIndex, 'format', 'rub');
        });

        const totalColIndex = colIndex + banks.length;
        const totalColLetter = String.fromCharCode(66 + totalColIndex - 1);
        setCellText(2, totalColIndex, `Итого ${method}`, 0);
        setCellBorder(2, totalColIndex, getBorderStyle('thick'), 0);

        const bankColLetters = banks.map((_, index) => String.fromCharCode(66 + (colIndex + index) -1));

        setCellText(6, totalColIndex, `=AVERAGE(${bankColLetters.map(l => l + '7').join(',')})`, 0);
        setCellStyle(6, totalColIndex, 'format', 'rub'); 

        setCellText(7, totalColIndex, `=SUM(${bankColLetters.map(l => l + '8').join(',')})`, 0);

        setCellText(8, totalColIndex, `=AVERAGE(${bankColLetters.map(l => l + '9').join(',')})`, 0);
        setCellStyle(8, totalColIndex, 'format', 'rub'); 

        setCellText(10, totalColIndex, `=SUM(${bankColLetters.map(l => l + '11').join(',')})`, 0);
        setCellStyle(10, totalColIndex, 'format', 'rub'); 

        setCellText(12, totalColIndex, `=${totalColLetter}15/${totalColLetter}8*100`, 0);
        setCellStyle(12, totalColIndex, 'format', 'percent'); 

        setCellText(13, totalColIndex, `=SUM(${bankColLetters.map(l => l + '14').join(',')})`, 0);
        setCellStyle(13, totalColIndex, 'format', 'rub'); 

        setCellText(14, totalColIndex, `=SUM(${bankColLetters.map(l => l + '15').join(',')})`, 0);

        setCellText(16, totalColIndex, `=${totalColLetter}19/${totalColLetter}8*100`, 0);
        setCellStyle(16, totalColIndex, 'format', 'percent'); 

        setCellText(17, totalColIndex, `=SUM(${bankColLetters.map(l => l + '18').join(',')})`, 0);
        setCellStyle(17, totalColIndex, 'format', 'rub'); 

        setCellText(18, totalColIndex, `=SUM(${bankColLetters.map(l => l + '19').join(',')})`, 0);

        setCellText(20, totalColIndex, `=SUM(${bankColLetters.map(l => l + '21').join(',')})`, 0);
        setCellStyle(20, totalColIndex, 'format', 'rub');

        setCellText(21, totalColIndex, `=AVERAGE(${bankColLetters.map(l => l + '22').join(',')})`, 0);
        setCellStyle(21, totalColIndex, 'format', 'percent'); 

        setCellText(22, totalColIndex, `=SUM(${bankColLetters.map(l => l + '23').join(',')})`, 0);
        setCellStyle(22, totalColIndex, 'format', 'rub'); 

        setCellText(24, totalColIndex, `=SUM(${bankColLetters.map(l => l + '25').join(',')})`, 0);
        setCellStyle(24, totalColIndex, 'format', 'rub'); 

        setCellText(25, totalColIndex, `=${totalColLetter}11-${totalColLetter}25`, 0);
        setCellStyle(25, totalColIndex, 'format', 'rub'); 

        colIndex = totalColIndex + 1;
    });

    setCellText(0, colIndex, 'итого Процессинг', 0);
    setCellBorder(0, colIndex, getBorderStyle('thick'), 0);
    setCellBackgroundColor(0, colIndex, '#d9ead3', 0);

    const totalColLetter = String.fromCharCode(66 + colIndex - 1);

    setCellText(6, colIndex, averagePrice, 0);
    setCellStyle(6, colIndex, 'format', 'rub'); 

    setCellText(7, colIndex, totalBanksCount, 0);

    setCellText(8, colIndex, averageProfit, 0);
    setCellStyle(8, colIndex, 'format', 'rub'); 

    setCellText(10, colIndex, sumProfit, 0);
    setCellStyle(10, colIndex, 'format', 'rub'); 

    setCellText(12, colIndex, `=${totalColLetter}15/${totalColLetter}8*100`, 0);
    setCellStyle(12, colIndex, 'format', 'percent'); 

    setCellText(13, colIndex, `=${totalColLetter}7*${totalColLetter}15`, 0);
    setCellStyle(13, colIndex, 'format', 'rub'); 

    setCellText(14, colIndex, totalProblemCount, 0);

    setCellText(16, colIndex, `=${totalColLetter}19/${totalColLetter}8*100`, 0);
    setCellStyle(16, colIndex, 'format', 'percent');

    setCellText(17, colIndex, `=${totalColLetter}7*${totalColLetter}19`, 0);
    setCellStyle(17, colIndex, 'format', 'rub'); 

    setCellText(18, colIndex, totalBlockCount, 0);

    setCellText(20, colIndex, `=${totalColLetter}7*${totalColLetter}8`, 0);
    setCellStyle(20, colIndex, 'format', 'rub'); 

    setCellText(21, colIndex, '0', 0);
    setCellStyle(21, colIndex, 'format', 'percent');

    setCellText(22, colIndex, `=(${totalColLetter}11-${totalColLetter}18)*${totalColLetter}22`, 0);
    setCellStyle(22, colIndex, 'format', 'rub'); 

    setCellText(24, colIndex, `=${totalColLetter}14+${totalColLetter}18+${totalColLetter}21+${totalColLetter}23`, 0);
    setCellStyle(24, colIndex, 'format', 'rub');

    setCellText(25, colIndex, `=${totalColLetter}11-${totalColLetter}25`, 0);
    setCellStyle(25, colIndex, 'format', 'rub'); 

    for (let i = 1; i <= colIndex; i++) {
        setCellBackgroundColor(0, i, '#d9ead3', 0);
    }

    return { colIndex, methodColIndexes };
}

const EXCLUDED_OPERATIONS = new Set([
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

const PROJECTS_TO_PROCESS = ['Процессинг//Переводы', 'Процессинг//Наличка'];

const DEFAULT_ROW_INDEX = 28;
const TOTAL_SUMS_KEYS = ['Процессинг//Переводы', 'Процессинг//Наличка', 'Итого'];

function formatRub(value) {
    return value.toFixed(2);
}

function setFormattedCellText(row, col, value) {
    setCellText(row, col, formatRub(value), 0);
    setCellStyle(row, col, 'format', 'rub');
}

function processRecords(records, exchangeRate, projects) {
    const operationSums = new Map();

    records.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sumUSD = parseFloat(record['Бухгалтерия_Сумма']);

        if (operation && !EXCLUDED_OPERATIONS.has(operation) && projects.includes(project) && !isNaN(sumUSD)) {
            const sumRUB = exchangeRate !== null ? sumUSD * exchangeRate : 0;
            if (!operationSums.has(operation)) {
                operationSums.set(operation, {
                    'Процессинг//Переводы': 0,
                    'Процессинг//Наличка': 0
                });
            }
            const projectSum = operationSums.get(operation);
            projectSum[project] += sumRUB;
        }
    });

    return operationSums;
}

function renderTotals(currentRowIndex, colIndex, operation, sums, methodColIndexes, totalSumsDB2) {
    setCellText(currentRowIndex, 0, operation, 0);

    const переводSumRUB = sums['Процессинг//Переводы'] || 0;
    const наличкаSumRUB = sums['Процессинг//Наличка'] || 0;

    const totalProcessingSumRUB = переводSumRUB + наличкаSumRUB;

    setFormattedCellText(currentRowIndex, colIndex, totalProcessingSumRUB);
    totalSumsDB2['Процессинг//Переводы'] += переводSumRUB;
    totalSumsDB2['Процессинг//Наличка'] += наличкаSumRUB;
    totalSumsDB2['Итого'] += totalProcessingSumRUB;

    ['Переводы', 'Наличка'].forEach(method => {
        if (methodColIndexes[method]) {
            const methodTotalColIndex = methodColIndexes[method].end;
            const sum = sums[`Процессинг//${method}`] || 0;
            setFormattedCellText(currentRowIndex, methodTotalColIndex, sum);
        }
    });

    return currentRowIndex + 1;
}

function addExchangeRateRow(currentRowIndex, exchangeRate) {
    setCellText(currentRowIndex, 0, "Текущий курс USD/RUB:", 0);
    if (exchangeRate !== null) {
        setFormattedCellText(currentRowIndex, 1, exchangeRate);
    }
    return currentRowIndex + 1;
}

function addFinalTotals(currentRowIndex, colIndex, methodColIndexes, totalSumsDB2) {
    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на поднаправление", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    ['Процессинг//Переводы', 'Процессинг//Наличка'].forEach(project => {
        if (methodColIndexes[project.split('//')[1]]) {
            const methodTotalColIndex = methodColIndexes[project.split('//')[1]].end;
            setFormattedCellText(currentRowIndex, methodTotalColIndex, totalSumsDB2[project]);
        }
    });

    setFormattedCellText(currentRowIndex, colIndex, totalSumsDB2['Итого']);
    return currentRowIndex + 1;
}

function addDBSummary(currentRowIndex, colIndex, totalSumsDB2, exchangeRateRowIndexDB2) {
    setCellText(currentRowIndex, 0, "₽ DB 2", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    const totalColLetterDB2 = String.fromCharCode(65 + colIndex);
    const db2Formula = `=${totalColLetterDB2}26 - ${totalColLetterDB2}${exchangeRateRowIndexDB2}`;
    setCellText(currentRowIndex, colIndex, db2Formula, 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'rub');
    return currentRowIndex + 1;
}

function renderDB3(currentRowIndex, colIndex, records, exchangeRate) {
    const processingOperationsDB3 = new Map();

    records.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sumUSD = parseFloat(record['Бухгалтерия_Сумма']);

        if (
            operation &&
            !EXCLUDED_OPERATIONS.has(operation) &&
            project.startsWith('Процессинг//') &&
            !PROJECTS_TO_PROCESS.includes(project) &&
            !isNaN(sumUSD)
        ) {
            const sumRUB = exchangeRate !== null ? sumUSD * exchangeRate : 0;
            processingOperationsDB3.set(operation, (processingOperationsDB3.get(operation) || 0) + sumRUB);
        }
    });

    let processingTotalSumDB3 = 0;

    processingOperationsDB3.forEach((sum, operation) => {
        setCellText(currentRowIndex, 0, operation, 0);
        setFormattedCellText(currentRowIndex, colIndex, sum);
        processingTotalSumDB3 += sum;
        currentRowIndex++;
    });

    return { currentRowIndex, processingTotalSumDB3 };
}

function addDB3FinalRows(currentRowIndex, colIndex, processingTotalSumDB3, totalColLetterDB2, exchangeRateRowIndexDB2) {
    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на направление", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    setFormattedCellText(currentRowIndex, colIndex, processingTotalSumDB3);
    setFormattedCellText(currentRowIndex, colIndex + 1, processingTotalSumDB3);
    currentRowIndex++;

    setCellText(currentRowIndex, 0, "₽ DB 3", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    const db3Formula = `=${totalColLetterDB2}${exchangeRateRowIndexDB2 + 1} - ${totalColLetterDB2}${currentRowIndex}`;
    setCellText(currentRowIndex, colIndex, db3Formula, 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'rub');
    return currentRowIndex + 1;
}

function addBorders(borderStartRow, borderEndRow, methodColIndexes, colIndex, totalTotalsColIndex) {
    Object.values(methodColIndexes).forEach(({ end: lastColIndex }) => {
        for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
            setCellBorder(ri, lastColIndex, {
                right: [defaultBorderStyle.style, defaultBorderStyle.color]
            }, 0);
        }
    });

    const additionalBordersColIndexes = [0, colIndex, totalTotalsColIndex];
    additionalBordersColIndexes.forEach(targetColIndex => {
        for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
            const borderStyle = ['thick', 'thin'].includes(targetColIndex) ? getBorderStyle('thick') : getBorderStyle('thin');
            setCellBorder(ri, targetColIndex, {
                right: [borderStyle.right[0], borderStyle.right[1]]
            }, 0);
        }
    });

    const dbRows = [25, exchangeRateRowIndexDB2, currentRowIndex];
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

function renderOperationTypes(filteredRecordsDB2, { colIndex, methodColIndexes }, exchangeRate) {
    let currentRowIndex = DEFAULT_ROW_INDEX;

    // Инициализация сумм
    const totalSumsDB2 = {
        'Процессинг//Переводы': 0,
        'Процессинг//Наличка': 0,
        'Итого': 0
    };

    // Обработка операций для DB2
    const operationSums = processRecords(filteredRecordsDB2, exchangeRate, PROJECTS_TO_PROCESS);

    operationSums.forEach((sums, operation) => {
        if (operation === "Пересчёт кассы") return;
        currentRowIndex = renderTotals(currentRowIndex, colIndex, operation, sums, methodColIndexes, totalSumsDB2);
    });

    // Добавление курса обмена
    currentRowIndex = addExchangeRateRow(currentRowIndex, exchangeRate);

    // Итоговые строки для DB2
    currentRowIndex = addFinalTotals(currentRowIndex, colIndex, methodColIndexes, totalSumsDB2);

    // Добавление DB2
    const exchangeRateRowIndexDB2 = currentRowIndex + 2;
    currentRowIndex = addDBSummary(currentRowIndex, colIndex, totalSumsDB2, exchangeRateRowIndexDB2);

    // Обработка DB3
    const { currentRowIndex: newRowIndex, processingTotalSumDB3 } = renderDB3(currentRowIndex, colIndex, filteredRecordsDB2, exchangeRate);
    currentRowIndex = newRowIndex;

    // Итоговые строки для DB3
    currentRowIndex = addDB3FinalRows(currentRowIndex, colIndex, processingTotalSumDB3, String.fromCharCode(65 + colIndex), exchangeRateRowIndexDB2);

    // Добавление границ
    const totalTotalsColIndex = colIndex + 1;
    addBorders(0, currentRowIndex, methodColIndexes, colIndex, totalTotalsColIndex);
}


