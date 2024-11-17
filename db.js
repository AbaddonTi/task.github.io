function processBankRecords(filteredRecordsDB1) {
    const bankMethodMap = {},
        bankPricesMap = {},
        bankCountsMap = {},
        bankProfitsMap = {},
        bankProblemCountsMap = {},
        bankBlockSumMap = {},
        bankBlockCountsMap = {};

    let totalPrices = 0,
        totalPricesCount = 0,
        totalBanksCount = 0,
        totalProfits = 0,
        totalProfitsCount = 0,
        totalProblemCount = 0,
        totalBlockAmount = 0,
        totalBlockCount = 0;

    filteredRecordsDB1.forEach(record => {
        const bank = Array.isArray(record['Банк']) ? record['Банк'][0] : record['Банк'];
        const method = Array.isArray(record['Метод']) ? record['Метод'][0] : record['Метод'];
        const price = parseFloat(record['Цена']);
        const profit = parseFloat(record['Профит']);
        const status = Array.isArray(record['Статус']) ? record['Статус'] : [record['Статус']];
        const blockAmount = parseFloat(record['BLOCK']);

        if (!isNaN(blockAmount)) {
            bankBlockSumMap[method] = bankBlockSumMap[method] || {};
            bankBlockSumMap[method][bank] = (bankBlockSumMap[method][bank] || 0) + blockAmount;
            totalBlockAmount += blockAmount;
        }

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
        totalBlockCount,
        bankBlockSumMap,
        totalBlockAmount
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

            const blockSum = bankBlockSumMap[method]?.[bank]?.toFixed(2) || '0.00';
            setCellText(17, bankColIndex, blockSum, 0);
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

        const methodBlockSum = Object.values(bankBlockSumMap[method] || {}).reduce((a, b) => a + b, 0).toFixed(2);
        setCellText(17, totalColIndex, methodBlockSum, 0);
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

    setCellText(17, colIndex, totalBlockAmount.toFixed(2), 0);
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

    let currentRowIndex = 28; 
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
    setCellStyle(10, cashInColIndex, 'format', 'usd'); 

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

        // Конвертация сумм из USD в RUB
        const переводSumRub = exchangeRate ? (переводSum * exchangeRate).toFixed(2) : переводSum.toFixed(2);
        const наличкаSumRub = exchangeRate ? (наличкаSum * exchangeRate).toFixed(2) : наличкаSum.toFixed(2);
        const cashInSumRub = exchangeRate ? (cashInSum * exchangeRate).toFixed(2) : cashInSum.toFixed(2);

        const totalProcessingSumRub = (parseFloat(переводSumRub) + parseFloat(наличкаSumRub)).toFixed(2);

        setCellText(currentRowIndex, colIndex, totalProcessingSumRub, 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            setCellText(currentRowIndex, methodTotalColIndex, переводSumRub, 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
            totalSumsDB2['Процессинг//Переводы'] += parseFloat(переводSumRub);
        }
        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            setCellText(currentRowIndex, methodTotalColIndex, наличкаSumRub, 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
            totalSumsDB2['Процессинг//Наличка'] += parseFloat(наличкаSumRub);
        }

        totalSumsDB2['Итого'] += parseFloat(totalProcessingSumRub);

        if (cashInSumRub !== "0.00") {
            setCellText(currentRowIndex, cashInColIndex, cashInSumRub, 0);
            setCellStyle(currentRowIndex, cashInColIndex, 'format', 'rub'); 
            cashInTotalSumsDB2['Cash-In//'] += parseFloat(cashInSumRub);
            cashInTotalSumsDB2['Итого'] += parseFloat(cashInSumRub);
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

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на поднаправление", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    if (exchangeRate !== null) {
        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            const totalInRublesПереводы = (totalSumsDB2['Процессинг//Переводы']).toFixed(2); 
            setCellText(currentRowIndex, methodTotalColIndex, totalInRublesПереводы, 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
        }
        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            const totalInRublesНаличka = (totalSumsDB2['Процессинг//Наличка']).toFixed(2); 
            setCellText(currentRowIndex, methodTotalColIndex, totalInRublesНаличka, 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
        }
        const totalInRubles = totalSumsDB2['Итого'];
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
        const operationSumRub = exchangeRate ? (processingOperationsDB3[operation] * exchangeRate).toFixed(2) : processingOperationsDB3[operation].toFixed(2);

        setCellText(currentRowIndex, colIndex, operationSumRub, 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'rub');
        processingTotalSumDB3 += parseFloat(operationSumRub);
        currentRowIndex++;
    });

    currentRowIndex += 2; 

    // === Итоговые строки для DB3 ===
    const totalProcessingSumOnlyDB3 = processingTotalSumDB3;
    let cashInSumDB3 = 0;
    const totalTotalsSumDB3 = totalProcessingSumOnlyDB3 + cashInSumDB3;

    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на направление", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    if (exchangeRate !== null) {
        const totalInRublesProcessingDB3 = (totalProcessingSumOnlyDB3).toFixed(2); 
        const totalInRublesCashInDB3 = (cashInTotalSumsDB2['Cash-In//']).toFixed(2); 
        const totalInRublesTotalDB3 = totalTotalsSumDB3.toFixed(2);

        setCellText(currentRowIndex, colIndex, totalInRublesProcessingDB3, 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

        if (methodColIndexes['Cash-In//']) {
            setCellText(currentRowIndex, cashInColIndex, totalInRublesCashInDB3, 0);
            setCellStyle(currentRowIndex, cashInColIndex, 'format', 'rub');
        }

        setCellText(currentRowIndex, totalTotalsColIndex, totalInRublesTotalDB3, 0);
        setCellStyle(currentRowIndex, totalTotalsColIndex, 'format', 'rub');
    } else {
        console.warn('Exchange rate is not available for the рубль calculation.');
    }

    const rubleTotalRowIndexDB3 = currentRowIndex + 1;

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "₽ DB 3", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    if (exchangeRate !== null) {
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

