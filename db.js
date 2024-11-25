function processBankRecords(filteredRecordsDB1) {
    const bankMethodMap = {},
        bankPricesMap = {},
        bankCountsMap = {},
        bankProfitsMap = {},
        bankProblemCountsMap = {},
        bankBlockCountsMap = {},
        bankBlockSumMap = {};

    let totalPrices = 0,
        totalPricesCount = 0,
        totalBanksCount = 0,
        totalProfits = 0,
        totalProfitsCount = 0,
        totalProblemCount = 0,
        totalBlockCount = 0,
        totalBlockAmount = 0;

    filteredRecordsDB1.forEach(record => {
        const bank = Array.isArray(record['Банк']) ? record['Банк'][0] : record['Банк'];
        const method = Array.isArray(record['Метод']) ? record['Метод'][0] : record['Метод'];
        const price = parseFloat(record['Цена']);
        const profit = parseFloat(record['Профит']);
        const status = Array.isArray(record['Статус']) ? record['Статус'] : [record['Статус']];
        const blockAmount = parseFloat(record['BLOCK']) || 0;

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

            bankBlockSumMap[method] = bankBlockSumMap[method] || {};
            bankBlockSumMap[method][bank] = (bankBlockSumMap[method][bank] || 0);

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

            if (status.includes('блокировка') && !isNaN(blockAmount)) {
                bankBlockSumMap[method][bank] += blockAmount;
                totalBlockAmount += blockAmount;
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
        bankBlockSumMap,
        totalPrices,
        totalPricesCount,
        totalBanksCount,
        totalProfits,
        totalProfitsCount,
        totalProblemCount,
        totalBlockCount,
        totalBlockAmount
    };
}


function setCell(row, col, text, format) {
    setCellText(row, col, text, 0);
    if (format) {
        setCellStyle(row, col, 'format', format);
    }
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
        bankBlockSumMap,
        totalBlockAmount,
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

        setCell(2, methodStartColIndex, method);
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

            const cellData = [
                { row: 3, text: bank },
                { row: 6, text: calculateAverage(bankPricesMap[method][bank]), format: 'rub' },
                { row: 7, text: bankCountsMap[method][bank] },
                { row: 8, text: calculateAverage(bankProfitsMap[method][bank]), format: 'rub' },
                { row: 10, text: `=${colLetter}8*${colLetter}9`, format: 'rub' },
                { row: 12, text: `=${colLetter}15/${colLetter}8*100`, format: 'percent' },
                { row: 13, text: `=${colLetter}7*${colLetter}15`, format: 'rub' },
                { row: 14, text: bankProblemCountsMap[method]?.[bank] || 0 },
                { row: 16, text: `=${colLetter}19/${colLetter}8*100`, format: 'percent' },
                { row: 17, text: bankBlockSumMap[method]?.[bank]?.toFixed(2) || '0.00', format: 'rub' },
                { row: 18, text: bankBlockCountsMap[method]?.[bank] || 0 },
                { row: 20, text: `=${colLetter}7*${colLetter}8`, format: 'rub' },
                { row: 21, text: '0', format: 'percent' },
                { row: 22, text: `=(${colLetter}11-${colLetter}18)*${colLetter}22`, format: 'rub' },
                { row: 24, text: `=${colLetter}14+${colLetter}18+${colLetter}21+${colLetter}23`, format: 'rub' },
                { row: 25, text: `=${colLetter}11-${colLetter}25`, format: 'rub' },
                { row: 26, text: `=(${colLetter}11-${colLetter}25)/${colLetter}11`, format: 'percent' },
            ];

            cellData.forEach(({ row, text, format }) => {
                setCell(row, bankColIndex, text, format);
            });
        });

        const totalColIndex = colIndex + banks.length;
        const totalColLetter = String.fromCharCode(66 + totalColIndex - 1);

        setCell(2, totalColIndex, `Итого ${method}`);
        setCellBorder(2, totalColIndex, getBorderStyle('thick'));

        const bankColLetters = banks.map((_, index) => String.fromCharCode(66 + (colIndex + index) - 1));

        const totalCellData = [
            { row: 6, text: `=AVERAGE(${bankColLetters.map(l => l + '7').join(',')})`, format: 'rub' },
            { row: 7, text: `=SUM(${bankColLetters.map(l => l + '8').join(',')})` },
            { row: 8, text: `=AVERAGE(${bankColLetters.map(l => l + '9').join(',')})`, format: 'rub' },
            { row: 10, text: `=SUM(${bankColLetters.map(l => l + '11').join(',')})`, format: 'rub' },
            { row: 12, text: `=${totalColLetter}15/${totalColLetter}8*100`, format: 'percent' },
            { row: 13, text: `=SUM(${bankColLetters.map(l => l + '14').join(',')})`, format: 'rub' },
            { row: 14, text: `=SUM(${bankColLetters.map(l => l + '15').join(',')})` },
            { row: 16, text: `=${totalColLetter}19/${totalColLetter}8*100`, format: 'percent' },
            { row: 17, text: Object.values(bankBlockSumMap[method] || {}).reduce((a, b) => a + b, 0).toFixed(2), format: 'rub' },
            { row: 18, text: `=SUM(${bankColLetters.map(l => l + '19').join(',')})` },
            { row: 20, text: `=SUM(${bankColLetters.map(l => l + '21').join(',')})`, format: 'rub' },
            { row: 21, text: `=AVERAGE(${bankColLetters.map(l => l + '22').join(',')})`, format: 'percent' },
            { row: 22, text: `=SUM(${bankColLetters.map(l => l + '23').join(',')})`, format: 'rub' },
            { row: 24, text: `=SUM(${bankColLetters.map(l => l + '25').join(',')})`, format: 'rub' },
            { row: 25, text: `=${totalColLetter}11-${totalColLetter}25`, format: 'rub' },
            { row: 26, text: `=(${totalColLetter}11 - ${totalColLetter}25)/${totalColLetter}11`, format: 'percent' },
        ];

        totalCellData.forEach(({ row, text, format }) => {
            setCell(row, totalColIndex, text, format);
        });

        colIndex = totalColIndex + 1;
    });

    setCell(0, colIndex, 'итого Процессинг');
    setCellBorder(0, colIndex, getBorderStyle('thick'));
    setCellBackgroundColor(0, colIndex, '#d9ead3', 0);

    const totalColLetter = String.fromCharCode(66 + colIndex - 1);

    const processingCellData = [
        { row: 6, text: averagePrice, format: 'rub' },
        { row: 7, text: totalBanksCount },
        { row: 8, text: averageProfit, format: 'rub' },
        { row: 10, text: sumProfit, format: 'rub' },
        { row: 12, text: `=${totalColLetter}15/${totalColLetter}8*100`, format: 'percent' },
        { row: 13, text: `=${totalColLetter}7*${totalColLetter}15`, format: 'rub' },
        { row: 14, text: totalProblemCount },
        { row: 16, text: `=${totalColLetter}19/${totalColLetter}8*100`, format: 'percent' },
        { row: 17, text: totalBlockAmount.toFixed(2), format: 'rub' },
        { row: 18, text: totalBlockCount },
        { row: 20, text: `=${totalColLetter}7*${totalColLetter}8`, format: 'rub' },
        { row: 21, text: '0', format: 'percent' },
        { row: 22, text: `=(${totalColLetter}11-${totalColLetter}18)*${totalColLetter}22`, format: 'rub' },
        { row: 24, text: `=${totalColLetter}14+${totalColLetter}18+${totalColLetter}21+${totalColLetter}23`, format: 'rub' },
        { row: 25, text: `=${totalColLetter}11-${totalColLetter}25`, format: 'rub' },
        { row: 26, text: `=(${totalColLetter}11 - ${totalColLetter}25)/${totalColLetter}11`, format: 'percent' },
    ];

    processingCellData.forEach(({ row, text, format }) => {
        setCell(row, colIndex, text, format);
    });

    for (let i = 1; i <= colIndex; i++) {
        setCellBackgroundColor(0, i, '#d9ead3', 0);
    }

    return { colIndex, methodColIndexes };

    function calculateAverage(values) {
        if (values && values.length) {
            return (values.reduce((a, b) => a + b, 0) / values.length).toFixed(2);
        }
        return '0.00';
    }
}


function renderOperationTypes(filteredRecordsDB2, { colIndex, methodColIndexes }, exchangeRate) {
    const excludedOperations = new Set([
        "Пересчёт кассы",
        "Внесение Слава", "Внесение Patek", "Внесение Роман",
        "Вывод Роман", "Вывод Слава", "Вывод Patek",
        "Выплаты Consulting", "Доход от рефералов", "Расходы на дропов",
        "Апелляция", "Перевод (получаем на счёт)", "Перевод (снимаем с счёта)"
    ]);

    const projectsToProcess = ['Процессинг//Переводы', 'Процессинг//Наличка'];
    const cashInProject = 'Cash-In//';

    let currentRowIndex = 29; // Начальная строка для вывода данных

    // Инициализация сумм для DB2
    const totalSumsDB2 = { 'Процессинг//Переводы': 0, 'Процессинг//Наличка': 0, 'Итого': 0 };

    // Проверка и установка индекса столбца для Cash-In
    if (!methodColIndexes[cashInProject]) {
        methodColIndexes[cashInProject] = { start: colIndex + 1, end: colIndex + 1 };
    }
    const cashInColIndex = methodColIndexes[cashInProject].end;

    setCellText(0, cashInColIndex, cashInProject);

    // Вспомогательные функции
    const setCell = (row, col, text, format) => {
        setCellText(row, col, text, 0);
        if (format) setCellStyle(row, col, 'format', format);
    };

    const addSum = (operationSums, operation, project, sum) => {
        if (!operationSums[operation]) operationSums[operation] = {};
        if (!operationSums[operation][project]) operationSums[operation][project] = 0;
        operationSums[operation][project] += sum;
    };

    const convertToRubles = amount => exchangeRate ? (amount * exchangeRate).toFixed(2) : amount.toFixed(2);

    // Обработка пересчета кассы для Cash-In
    let cashInRecountSum = 0;
    filteredRecordsDB2.forEach(({ 'Бухгалтерия_Операция': operation, 'Бухгалтерия_Проект': project, 'Бухгалтерия_Сумма': amount }) => {
        const sum = parseFloat(amount);
        if (operation === "Пересчёт кассы" && project === cashInProject && !isNaN(sum)) {
            cashInRecountSum += sum;
        }
    });

    setCell(10, cashInColIndex, cashInRecountSum.toFixed(2), 'usd');

    // Инициализация операций и сумм для DB2
    const uniqueOperationsDB2 = new Set();
    const operationSumsDB2 = {};

    // Обработка операций для DB2
    filteredRecordsDB2.forEach(({ 'Бухгалтерия_Операция': operation, 'Бухгалтерия_Проект': project, 'Бухгалтерия_Сумма': amount }) => {
        const sum = parseFloat(amount);
        if (!operation || excludedOperations.has(operation) || isNaN(sum)) return;

        if (projectsToProcess.includes(project)) {
            addSum(operationSumsDB2, operation, project, sum);
            uniqueOperationsDB2.add(operation);
        }
    });

    // Обработка уникальных операций для DB2
    uniqueOperationsDB2.forEach(operation => {
        setCell(currentRowIndex, 0, operation);

        const sums = operationSumsDB2[operation] || {};
        const переводSumRub = convertToRubles(sums['Процессинг//Переводы'] || 0);
        const наличкаSumRub = convertToRubles(sums['Процессинг//Наличка'] || 0);

        const totalProcessingSumRub = (parseFloat(переводSumRub) + parseFloat(наличкаSumRub)).toFixed(2);

        // Установка сумм в ячейки
        setCell(currentRowIndex, colIndex, totalProcessingSumRub, 'rub');

        ['Переводы', 'Наличка'].forEach(method => {
            if (methodColIndexes[method]) {
                const methodTotalColIndex = methodColIndexes[method].end;
                const sumRub = method === 'Переводы' ? переводSumRub : наличкаSumRub;
                setCell(currentRowIndex, methodTotalColIndex, sumRub, 'rub');
                totalSumsDB2[`Процессинг//${method}`] += parseFloat(sumRub);
            }
        });

        totalSumsDB2['Итого'] += parseFloat(totalProcessingSumRub);
        currentRowIndex++;
    });

    // Установка текущего курса USD/RUB перед '₽ Итого: фикс косты на поднаправление'
    currentRowIndex++;
    setCell(currentRowIndex, 0, "Текущий курс USD/RUB:");
    if (exchangeRate !== null) {
        setCell(currentRowIndex, 1, exchangeRate.toFixed(2), 'rub');
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }
    const exchangeRateRowIndexDB2 = currentRowIndex + 1;

    // Умножаем сумму на курс и записываем в строку 25
    const cashInColLetter = String.fromCharCode(65 + cashInColIndex);
    if (exchangeRate !== null) {
        setCell(25, cashInColIndex, `=${cashInColLetter}11*B${exchangeRateRowIndexDB2}`, 'rub');
    }

    // Итоги итогов
    const totalTotalsColIndex = cashInColIndex + 1;
    setCell(0, totalTotalsColIndex, "Итоги Итогов:");
    const totalProcessingColLetter = String.fromCharCode(65 + colIndex);
    setCell(25, totalTotalsColIndex, `=${totalProcessingColLetter}26+${cashInColLetter}26`, 'rub');

    // Итоговые строки для DB2
    currentRowIndex++;
    setCell(currentRowIndex, 0, "₽ Итого: фикс косты на поднаправление");

    // Устанавливаем значение 0 для 'Cash-In'
    setCell(currentRowIndex, cashInColIndex, 0, 'rub');

    if (exchangeRate !== null) {
        ['Переводы', 'Наличка'].forEach(method => {
            if (methodColIndexes[method]) {
                const methodTotalColIndex = methodColIndexes[method].end;
                const totalInRubles = totalSumsDB2[`Процессинг//${method}`].toFixed(2);
                setCell(currentRowIndex, methodTotalColIndex, totalInRubles, 'rub');
            }
        });
        const totalInRubles = totalSumsDB2['Итого'].toFixed(2);
        setCell(currentRowIndex, colIndex, totalInRubles, 'rub');
    }

    const rubleSubTotalRowIndexDB2 = currentRowIndex + 1;

    currentRowIndex++;
    setCell(currentRowIndex, 0, "₽ DB 2");

    // Определяем переменную totalColLetterDB2
    const totalColLetterDB2 = String.fromCharCode(65 + colIndex);

    // Формула для 'Cash-In' столбца в '₽ DB 2'
    setCell(currentRowIndex, cashInColIndex, `=${cashInColLetter}26 - ${cashInColLetter}${rubleSubTotalRowIndexDB2}`, 'rub');

    // Формула для 'Процессинг' столбца
    const db2Formula = `=${totalColLetterDB2}26 - ${totalColLetterDB2}${rubleSubTotalRowIndexDB2}`;
    setCell(currentRowIndex, colIndex, db2Formula, 'rub');

    // === Переход к DB3 ===
    currentRowIndex += 2;
    const db3StartRowIndex = currentRowIndex;

    // Обработка операций для DB3
    const processingOperationsDB3 = {};
    const cashInOperationsDB3 = {};

    filteredRecordsDB2.forEach(({ 'Бухгалтерия_Операция': operation, 'Бухгалтерия_Проект': project, 'Бухгалтерия_Сумма': amount }) => {
        const sum = parseFloat(amount);
        if (!operation || excludedOperations.has(operation) || isNaN(sum)) return;

        if (project.startsWith('Процессинг//') && !projectsToProcess.includes(project)) {
            processingOperationsDB3[operation] = (processingOperationsDB3[operation] || 0) + sum;
        } else if (project === cashInProject && operation !== "Пересчёт кассы") {
            cashInOperationsDB3[operation] = (cashInOperationsDB3[operation] || 0) + sum;
        }
    });

    let processingTotalSumDB3 = 0;
    let cashInTotalSumDB3 = 0;

    const allOperationsDB3 = new Set([
        ...Object.keys(processingOperationsDB3),
        ...Object.keys(cashInOperationsDB3)
    ]);

    allOperationsDB3.forEach(operation => {
        setCell(currentRowIndex, 0, operation);

        const processingSumRub = convertToRubles(processingOperationsDB3[operation] || 0);
        const cashInSumRub = convertToRubles(cashInOperationsDB3[operation] || 0);

        if (processingSumRub !== "0.00") {
            setCell(currentRowIndex, colIndex, processingSumRub, 'rub');
            processingTotalSumDB3 += parseFloat(processingSumRub);
        }

        if (cashInSumRub !== "0.00") {
            setCell(currentRowIndex, cashInColIndex, cashInSumRub, 'rub');
            cashInTotalSumDB3 += parseFloat(cashInSumRub);
        }

        currentRowIndex++;
    });

    currentRowIndex += 2;

    // Итоговые строки для DB3
    setCell(currentRowIndex, 0, "₽ Итого: фикс косты на направление");

    if (exchangeRate !== null) {
        const totalInRublesProcessingDB3 = processingTotalSumDB3.toFixed(2);
        const totalInRublesCashInDB3 = cashInTotalSumDB3.toFixed(2);
        const totalInRublesTotalDB3 = (processingTotalSumDB3 + cashInTotalSumDB3).toFixed(2);

        if (totalInRublesProcessingDB3 !== "0.00") {
            setCell(currentRowIndex, colIndex, totalInRublesProcessingDB3, 'rub');
        }

        if (totalInRublesCashInDB3 !== "0.00") {
            setCell(currentRowIndex, cashInColIndex, totalInRublesCashInDB3, 'rub');
        }

        if (totalInRublesTotalDB3 !== "0.00") {
            setCell(currentRowIndex, totalTotalsColIndex, totalInRublesTotalDB3, 'rub');
        }
    }

    const rubleTotalRowIndexDB3 = currentRowIndex + 1;

    currentRowIndex++;
    setCell(currentRowIndex, 0, "₽ DB 3");

    // Формулы для DB3
    const totalColLetterDB3 = String.fromCharCode(65 + colIndex);

    const db3FormulaProcessing = `=${totalColLetterDB3}${db3StartRowIndex - 1} - ${totalColLetterDB3}${rubleTotalRowIndexDB3}`;
    setCell(currentRowIndex, colIndex, db3FormulaProcessing, 'rub');

    const db3FormulaCashIn = `=${cashInColLetter}${db3StartRowIndex - 1} - ${cashInColLetter}${rubleTotalRowIndexDB3}`;
    setCell(currentRowIndex, cashInColIndex, db3FormulaCashIn, 'rub');

    const db3FormulaTotal = `=${totalColLetterDB3}${currentRowIndex + 1}+${cashInColLetter}${currentRowIndex + 1}`;
    setCell(currentRowIndex, totalTotalsColIndex, db3FormulaTotal, 'rub');

    // Добавление границ и стилей
    addBordersAndStyles();

    // Вспомогательная функция для добавления границ и стилей
    function addBordersAndStyles() {
        const borderStartRow = 0;
        const borderEndRow = currentRowIndex;

        // Добавление правых границ для методов
        Object.values(methodColIndexes).forEach(({ end: lastColIndex }) => {
            for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
                setCellBorder(ri, lastColIndex, { right: [defaultBorderStyle.style, defaultBorderStyle.color] }, 0);
            }
        });

        // Добавление правых границ для дополнительных колонок
        const additionalBordersColIndexes = [0, colIndex, cashInColIndex, totalTotalsColIndex];

        additionalBordersColIndexes.forEach(targetColIndex => {
            for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
                const borderStyle = getBorderStyle(
                    [0, colIndex, cashInColIndex, totalTotalsColIndex].includes(targetColIndex) ? 'thick' : 'thin'
                );
                setCellBorder(ri, targetColIndex, { right: [borderStyle.right[0], borderStyle.right[1]] }, 0);
            }
        });

        // Добавление горизонтальных линий под каждым DB
        const dbRows = [26, rubleSubTotalRowIndexDB2, rubleTotalRowIndexDB3];

        dbRows.forEach(dbRow => {
            for (let col = 0; col <= totalTotalsColIndex; col++) {
                setCellBorder(dbRow, col, { bottom: ['thin', '#000000'] }, 0);
            }
        });

        // Добавление верхней границы
        for (let col = 0; col <= totalTotalsColIndex; col++) {
            setCellBorder(0, col, { bottom: ['thin', '#000000'] }, 0);
        }
    }
}

