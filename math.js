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

    const projectsToProcess = ['Процессинг//Переводы', 'Процессинг//Наличка'];
    const cashInProject = 'Cash-In//';

    let currentRowIndex = 28; // Начальная строка для вывода данных

    // Инициализация сумм
    const totalSumsDB2 = {
        'Процессинг//Переводы': 0,
        'Процессинг//Наличка': 0,
        'Итого': 0
    };

    // Вспомогательные функции

    function addSum(operationSums, operation, project, sum) {
        if (!operationSums[operation]) {
            operationSums[operation] = {};
        }
        if (!operationSums[operation][project]) {
            operationSums[operation][project] = 0;
        }
        operationSums[operation][project] += sum;
    }

    function convertToRubles(amount) {
        return exchangeRate ? (amount * exchangeRate).toFixed(2) : amount.toFixed(2);
    }

    // Инициализация операций и сумм для DB2
    const uniqueOperationsDB2 = new Set();
    const operationSumsDB2 = {};

    // Обработка операций для DB2
    filteredRecordsDB2.forEach(record => {
        const { 'Бухгалтерия_Операция': operation, 'Бухгалтерия_Проект': project, 'Бухгалтерия_Сумма': amount } = record;
        const sum = parseFloat(amount);

        if (!operation || excludedOperations.has(operation) || isNaN(sum)) {
            return;
        }

        if (projectsToProcess.includes(project)) {
            addSum(operationSumsDB2, operation, project, sum);
            uniqueOperationsDB2.add(operation);
        }
    });

    // Обработка уникальных операций для DB2
    uniqueOperationsDB2.forEach(operation => {
        setCell(currentRowIndex, 0, operation);

        const переводSum = operationSumsDB2[operation]['Процессинг//Переводы'] || 0;
        const наличкаSum = operationSumsDB2[operation]['Процессинг//Наличка'] || 0;

        const переводSumRub = convertToRubles(переводSum);
        const наличкаSumRub = convertToRubles(наличкаSum);

        const totalProcessingSumRub = (parseFloat(переводSumRub) + parseFloat(наличкаSumRub)).toFixed(2);

        // Установка сумм в ячейки
        setCell(currentRowIndex, colIndex, totalProcessingSumRub, 'rub');

        if (methodColIndexes['Переводы']) {
            const methodTotalColIndex = methodColIndexes['Переводы'].end;
            setCell(currentRowIndex, methodTotalColIndex, переводSumRub, 'rub');
            totalSumsDB2['Процессинг//Переводы'] += parseFloat(переводSumRub);
        }

        if (methodColIndexes['Наличка']) {
            const methodTotalColIndex = methodColIndexes['Наличка'].end;
            setCell(currentRowIndex, methodTotalColIndex, наличкаSumRub, 'rub');
            totalSumsDB2['Процессинг//Наличка'] += parseFloat(наличкаSumRub);
        }

        totalSumsDB2['Итого'] += parseFloat(totalProcessingSumRub);

        currentRowIndex++;
    });

    // Остальная часть обработки DB2 (итоги, формулы) остается без изменений

    const rubleSubTotalRowIndexDB2 = currentRowIndex + 1;

    currentRowIndex++;
    setCell(currentRowIndex, 0, "₽ DB 2");

    const totalColLetterDB2 = String.fromCharCode(65 + colIndex);
    const db2Formula = `=${totalColLetterDB2}26 - ${totalColLetterDB2}${rubleSubTotalRowIndexDB2}`;
    setCell(currentRowIndex, colIndex, db2Formula, 'rub');

    // Переход к DB3
    currentRowIndex += 2;
    const db3StartRowIndex = currentRowIndex;

    // Обработка операций для DB3
    const processingOperationsDB3 = {};
    filteredRecordsDB2.forEach(record => {
        const { 'Бухгалтерия_Операция': operation, 'Бухгалтерия_Проект': project, 'Бухгалтерия_Сумма': amount } = record;
        const sum = parseFloat(amount);

        if (
            operation &&
            !excludedOperations.has(operation) &&
            (
                (project.startsWith('Процессинг//') && !projectsToProcess.includes(project)) ||
                project === cashInProject
            ) &&
            !isNaN(sum)
        ) {
            if (!processingOperationsDB3[operation]) {
                processingOperationsDB3[operation] = {};
            }
            if (!processingOperationsDB3[operation][project]) {
                processingOperationsDB3[operation][project] = 0;
            }
            processingOperationsDB3[operation][project] += sum;
        }
    });

    // Собираем список проектов в DB3
    const projectsDB3 = new Set();
    Object.values(processingOperationsDB3).forEach(projectSums => {
        Object.keys(projectSums).forEach(project => {
            projectsDB3.add(project);
        });
    });

    // Обновляем methodColIndexes для новых проектов
    projectsDB3.forEach(project => {
        if (!methodColIndexes[project]) {
            colIndex += 1;
            methodColIndexes[project] = { start: colIndex, end: colIndex };
            setCellText(0, colIndex, project);
        }
    });

    // Обновляем totalTotalsColIndex
    const totalTotalsColIndex = colIndex + 1;
    setCellText(0, totalTotalsColIndex, "Итоги Итогов:");

    // Обработка уникальных операций в DB3
    let processingTotalSumDB3 = 0;
    Object.keys(processingOperationsDB3).forEach(operation => {
        setCell(currentRowIndex, 0, operation);

        let operationTotalSumRub = 0;

        projectsDB3.forEach(project => {
            const sum = processingOperationsDB3[operation][project] || 0;

            if (sum !== 0) {
                const sumRub = convertToRubles(sum);

                const projectColIndex = methodColIndexes[project].end;

                setCell(currentRowIndex, projectColIndex, sumRub, 'rub');

                operationTotalSumRub += parseFloat(sumRub);
            }
        });

        // Устанавливаем общую сумму по операции
        setCell(currentRowIndex, totalTotalsColIndex, operationTotalSumRub.toFixed(2), 'rub');

        processingTotalSumDB3 += operationTotalSumRub;

        currentRowIndex++;
    });

    currentRowIndex += 2;

    // Итоговые строки для DB3
    const totalSumsDB3 = {};
    projectsDB3.forEach(project => {
        totalSumsDB3[project] = 0;
    });

    Object.values(processingOperationsDB3).forEach(projectSums => {
        projectsDB3.forEach(project => {
            const sumRub = parseFloat(convertToRubles(projectSums[project] || 0));
            totalSumsDB3[project] += sumRub;
        });
    });

    setCell(currentRowIndex, 0, "₽ Итого: фикс косты на направление");

    projectsDB3.forEach(project => {
        const projectColIndex = methodColIndexes[project].end;
        const totalSum = totalSumsDB3[project].toFixed(2);
        setCell(currentRowIndex, projectColIndex, totalSum, 'rub');
    });

    const totalSumDB3 = Object.values(totalSumsDB3).reduce((a, b) => a + b, 0);
    setCell(currentRowIndex, totalTotalsColIndex, totalSumDB3.toFixed(2), 'rub');

    const rubleTotalRowIndexDB3 = currentRowIndex + 1;

    currentRowIndex++;
    setCell(currentRowIndex, 0, "₽ DB 3");

    projectsDB3.forEach(project => {
        const projectColIndex = methodColIndexes[project].end;
        const projectColLetter = String.fromCharCode(65 + projectColIndex);
        const formula = `=${projectColLetter}${db3StartRowIndex} - ${projectColLetter}${rubleTotalRowIndexDB3}`;
        setCell(currentRowIndex, projectColIndex, formula, 'rub');
    });

    const totalColLetterDB2 = String.fromCharCode(65 + totalTotalsColIndex);
    const db3Formula = `=${totalColLetterDB2}${rubleSubTotalRowIndexDB2 + 1} - ${totalColLetterDB2}${rubleTotalRowIndexDB3}`;
    setCell(currentRowIndex, totalTotalsColIndex, db3Formula, 'rub');

    // Добавление границ и стилей
    addBordersAndStyles();

    // Вспомогательная функция для добавления границ и стилей
    function addBordersAndStyles() {
        const borderStartRow = 0;
        const borderEndRow = currentRowIndex;

        // Добавление правых границ для методов
        Object.values(methodColIndexes).forEach(({ end: lastColIndex }) => {
            for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
                setCellBorder(ri, lastColIndex, {
                    right: [defaultBorderStyle.style, defaultBorderStyle.color]
                }, 0);
            }
        });

        // Добавление правых границ для дополнительных колонок
        const additionalBordersColIndexes = [
            0,
            colIndex,
            totalTotalsColIndex
        ];

        additionalBordersColIndexes.forEach(targetColIndex => {
            for (let ri = borderStartRow; ri <= borderEndRow; ri++) {
                const borderStyle = getBorderStyle(
                    targetColIndex === 0 || targetColIndex === colIndex || targetColIndex === totalTotalsColIndex
                        ? 'thick'
                        : 'thin'
                );
                setCellBorder(ri, targetColIndex, {
                    right: [borderStyle.right[0], borderStyle.right[1]]
                }, 0);
            }
        });

        // Добавление горизонтальных линий под каждым DB
        const dbRows = [25, rubleSubTotalRowIndexDB2, rubleTotalRowIndexDB3];

        dbRows.forEach(dbRow => {
            for (let col = 0; col <= totalTotalsColIndex; col++) {
                setCellBorder(dbRow, col, {
                    bottom: ['thin', '#000000']
                }, 0);
            }
        });

        // Добавление верхней границы
        for (let col = 0; col <= totalTotalsColIndex; col++) {
            setCellBorder(0, col, {
                bottom: ['thin', '#000000']
            }, 0);
        }
    }
}
