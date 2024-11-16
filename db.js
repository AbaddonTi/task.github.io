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
        'Cash-In//': 0, // Добавляем новый проект
        'Итого': 0
    };

    const projectsToProcess = ['Процессинг//Переводы', 'Процессинг//Наличка', 'Cash-In//']; // Добавляем новый проект

    // === Обработка операций для DB2 ===
    filteredRecordsDB2.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sumUSD = parseFloat(record['Бухгалтерия_Сумма']);

        if (
            operation &&
            !excludedOperations.has(operation) &&
            projectsToProcess.includes(project)
        ) {
            uniqueOperations.add(operation);
            if (!isNaN(sumUSD)) {
                if (!operationSums[operation]) {
                    operationSums[operation] = {
                        'Процессинг//Переводы': 0,
                        'Процессинг//Наличка': 0,
                        'Cash-In//': 0 // Добавляем новый проект
                    };
                }
                const sumRUB = exchangeRate !== null ? sumUSD * exchangeRate : 0;
                operationSums[operation][project] += sumRUB;
            }
        }
    });

    uniqueOperations.forEach(operation => {
        if (operation === "Пересчёт кассы") {
            return;
        }

        setCellText(currentRowIndex, 0, operation, 0);

        const переводSumRUB = operationSums[operation]['Процессинг//Переводы'] || 0;
        const наличкаSumRUB = operationSums[operation]['Процессинг//Наличка'] || 0;
        const cashInSumRUB = operationSums[operation]['Cash-In//'] || 0; // Сумма для нового проекта

        // Общая сумма по Процессинг без Cash-In
        const totalProcessingSumRUB = (переводSumRUB + наличкаSumRUB).toFixed(2);

        // Сумма для Cash-In
        const totalCashInSumRUB = cashInSumRUB.toFixed(2);

        // Устанавливаем сумму Процессинг
        setCellText(currentRowIndex, colIndex, totalProcessingSumRUB, 0);
        setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

        // Устанавливаем сумму Cash-In перед итоговыми итогами
        setCellText(currentRowIndex, colIndex + 1, totalCashInSumRUB, 0); // Новый столбец
        setCellStyle(currentRowIndex, colIndex + 1, 'format', 'rub');

        // Обновляем итоги
        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            setCellText(currentRowIndex, methodTotalColIndex, переводSumRUB.toFixed(2), 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
            totalSumsDB2['Процессинг//Переводы'] += переводSumRUB;
        }
        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            setCellText(currentRowIndex, methodTotalColIndex, наличкаSumRUB.toFixed(2), 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub'); 
            totalSumsDB2['Процессинг//Наличка'] += наличкаSumRUB;
        }
        if (methodColIndexes['Cash-In//']) { // Убедимся, что есть индекс для нового проекта
            const methodTotalColIndex = methodColIndexes['Cash-In//'].end;
            setCellText(currentRowIndex, methodTotalColIndex, cashInSumRUB.toFixed(2), 0);
            setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub');
            totalSumsDB2['Cash-In//'] += cashInSumRUB;
        }

        totalSumsDB2['Итого'] += parseFloat(totalProcessingSumRUB) + parseFloat(totalCashInSumRUB);

        currentRowIndex++;
    });

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "Текущий курс USD/RUB:", 0);

    if (exchangeRate !== null) {
        setCellText(currentRowIndex, 1, exchangeRate.toFixed(2), 0);
        setCellStyle(currentRowIndex, 1, 'format', 'rub');
    }

    const exchangeRateRowIndexDB2 = currentRowIndex + 2;

    currentRowIndex++;

    const totalTotalsColIndex = colIndex + 2; // Увеличиваем на 1 из-за нового столбца
    setCellText(0, totalTotalsColIndex, "Итоги Итогов:", 0);

    const totalProcessingColLetter = String.fromCharCode(65 + colIndex);
    const cashInColLetter = String.fromCharCode(65 + colIndex + 1); // Буква нового столбца
    setCellText(25, totalTotalsColIndex, `=${totalProcessingColLetter}26 + ${cashInColLetter}26`, 0);
    setCellStyle(25, totalTotalsColIndex, 'format', 'rub'); 

    // === Итоговые строки для DB2 ===

    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на поднаправление", 0);
    setCellStyle(currentRowIndex, 0, 'format', ''); 

    if (methodColIndexes['Переводы']) {
        const methodTotalColIndex = methodColIndexes['Переводы'].end;
        setCellText(currentRowIndex, methodTotalColIndex, totalSumsDB2['Процессинг//Переводы'].toFixed(2), 0);
        setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub');
    }
    if (methodColIndexes['Наличка']) {
        const methodTotalColIndex = methodColIndexes['Наличка'].end;
        setCellText(currentRowIndex, methodTotalColIndex, totalSumsDB2['Процессинг//Наличка'].toFixed(2), 0);
        setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub');
    }
    if (methodColIndexes['Cash-In//']) { // Итог для нового проекта
        const methodTotalColIndex = methodColIndexes['Cash-In//'].end;
        setCellText(currentRowIndex, methodTotalColIndex, totalSumsDB2['Cash-In//'].toFixed(2), 0);
        setCellStyle(currentRowIndex, methodTotalColIndex, 'format', 'rub');
    }

    setCellText(currentRowIndex, colIndex, (totalSumsDB2['Итого']).toFixed(2), 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'rub'); 

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "₽ DB 2", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    const totalColLetterDB2 = String.fromCharCode(65 + colIndex);
    const db2Formula = `=${totalColLetterDB2}26 + ${cashInColLetter}26 - ${totalColLetterDB2}${exchangeRateRowIndexDB2}`;
    setCellText(currentRowIndex, colIndex, db2Formula, 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

    // === Переход к DB3 ===
    currentRowIndex += 2; 

    const db3StartRowIndex = currentRowIndex;
    const processingOperationsDB3 = {};
    filteredRecordsDB2.forEach(record => {
        const operation = record['Бухгалтерия_Операция'];
        const project = record['Бухгалтерия_Проект'];
        const sumUSD = parseFloat(record['Бухгалтерия_Сумма']);

        if (
            operation &&
            !excludedOperations.has(operation) &&
            (project.startsWith('Процессинг//') || project === 'Cash-In//') && // Учитываем новый проект
            !projectsToProcess.includes(project)
        ) {
            if (!processingOperationsDB3[operation]) {
                processingOperationsDB3[operation] = 0;
            }
            if (!isNaN(sumUSD)) {
                const sumRUB = exchangeRate !== null ? sumUSD * exchangeRate : 0;
                processingOperationsDB3[operation] += sumRUB;
            }
        }
    });

    let processingTotalSumDB3 = 0;
    Object.keys(processingOperationsDB3).forEach(operation => {
        setCellText(currentRowIndex, 0, operation, 0);

        const operationSumRUB = processingOperationsDB3[operation].toFixed(2);
        setCellText(currentRowIndex, colIndex + 1, operationSumRUB, 0); // Новый столбец
        setCellStyle(currentRowIndex, colIndex + 1, 'format', 'rub');
        processingTotalSumDB3 += parseFloat(operationSumRUB);
        currentRowIndex++;
    });

    currentRowIndex += 2; 

    // === Итоговые строки для DB3 ===

    setCellText(currentRowIndex, 0, "₽ Итого: фикс косты на направление", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    const totalProcessingSumOnlyDB3 = processingTotalSumDB3;
    setCellText(currentRowIndex, colIndex + 1, totalProcessingSumOnlyDB3.toFixed(2), 0); // Новый столбец
    setCellStyle(currentRowIndex, colIndex + 1, 'format', 'rub');

    const totalTotalsSumDB3 = totalProcessingSumOnlyDB3;
    setCellText(currentRowIndex, colIndex + 2, totalTotalsSumDB3.toFixed(2), 0); // Обновляем индекс для итогов
    setCellStyle(currentRowIndex, colIndex + 2, 'format', 'rub');

    currentRowIndex++;
    setCellText(currentRowIndex, 0, "₽ DB 3", 0);
    setCellStyle(currentRowIndex, 0, 'format', '');

    const db3Formula = `=${totalColLetterDB2}${exchangeRateRowIndexDB2 + 1} + ${cashInColLetter}${exchangeRateRowIndexDB2 + 1} - ${totalColLetterDB2}${currentRowIndex}`;
    setCellText(currentRowIndex, colIndex, db3Formula, 0);
    setCellStyle(currentRowIndex, colIndex, 'format', 'rub');

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
        colIndex + 1, // Добавляем новый столбец
        totalTotalsColIndex 
    ];

    additionalBordersColIndexes.forEach(targetColIndex => {
        for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
            let borderStyle;
            if (
                targetColIndex === 0 || 
                targetColIndex === colIndex || 
                targetColIndex === colIndex + 1 || // Новый столбец
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
    const db2Row = exchangeRateRowIndexDB2;
    const db3Row = currentRowIndex; 

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
