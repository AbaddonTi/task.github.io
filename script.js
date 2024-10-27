const s = x_spreadsheet('#x-spreadsheet-demo');
function setCellText(row, col, text) {
    s.cellText(row, col, text);
}

function setCellStyle(ri, ci, property, value, sheetIndex = 0) {
    const dataProxy = s.datas[sheetIndex];
    const {styles, rows} = dataProxy;
    const cell = rows.getCellOrNew(ri, ci);
    let cstyle = {};

    if (cell.style !== undefined) {
        cstyle = {...styles[cell.style]};
    }

    if (property.startsWith('font')) {
        const nfont = {};
        nfont[property.split('-')[1]] = value;
        cstyle.font = {...(cstyle.font || {}), ...nfont};
    } else {
        cstyle[property] = value;
    }

    cell.style = dataProxy.addStyle(cstyle);

    s.reRender();
}

function setCellBackgroundColor(row, col, color, sheetIndex = 0) {
    setCellStyle(row, col, 'bgcolor', color, sheetIndex);
}

function setColumnWidth(colIndex, width, sheetIndex = 0) {
    const dataProxy = s.datas[sheetIndex];
    dataProxy.cols.setWidth(colIndex, width);
    s.reRender();
}

function setRowHeight(rowIndex, height, sheetIndex = 0) {
    const dataProxy = s.datas[sheetIndex];
    dataProxy.rows.setHeight(rowIndex, height);
    s.reRender();
}

function enableTextWrap(row, col, sheetIndex = 0) {
    setCellStyle(row, col, 'textwrap', true, sheetIndex);
}

// Функция для инициализации статических данных и стилей
function initializeSpreadsheet() {
    // Добавление рамок к ячейке A1
    setCellStyle(0, 0, 'border', {
        top: ['thick', '#000000'],
        bottom: ['thick', '#000000'],
        right: ['thick', '#000000'],
        left: ['thick', '#000000'],
    });
    // Установка заголовков и статических данных
    setCellText(0, 0, 'Статья:');
    setCellText(6, 0, 'Цена закуп. карты');
    setCellText(7, 0, 'Кол-во карт');
    setCellText(8, 0, 'Профит с 1 карты');
    setCellText(10, 0, 'Выручка наша:');
    setCellText(12, 0, '% проблемных');
    setCellText(13, 0, 'Сумма замороженных денег в проблемных картах');
    setCellText(14, 0, 'Кол-во карт проблемных');
    setCellText(16, 0, '% блока');
    setCellText(17, 0, 'Сумма в блоки (заморозка)');
    setCellText(18, 0, 'Кол-во карт в блок');
    setCellText(20, 0, 'закуп карточки');
    setCellText(21, 0, '% парней с выручки на руки');
    setCellText(22, 0, 'ФОТ операторы');
    setCellText(24, 0, 'Итого: переменные косты на подподнаправления');
    setCellText(25, 0, 'DB 1');

    // Статьи
    setCellText(0, 1, 'Процессинг:');

    // Стили
    setCellStyle(0, 0, 'font-bold', true);
    setCellStyle(25, 0, 'font-bold', true);

    setColumnWidth(0, 210);

    setRowHeight(0, 50);
    setRowHeight(13, 50);
    setRowHeight(24, 50);

    enableTextWrap(13, 0);
    enableTextWrap(24, 0);
}

async function fetchUniqueBanks(records) {
    try {
        const bankColumn = 'Банк';

        if (!records || records.length === 0) {
            
            return [];
        }

        const banks = records.map(record => record[bankColumn]).flat();

        const uniqueBanks = [...new Set(banks)].filter(bank => bank);

        return uniqueBanks;
    } catch (error) {
        console.error('Ошибка при получении банков:', error);
        return [];
    }
}

async function calculateMetrics(records, banks) {
    // Фильтрация записей, исключая строки с 'Дроповод' равным 'Consult Kayos' или 'Consult Titan'
    records = records.filter(record => record['Дроповод'] !== 'Consult Kayos' && record['Дроповод'] !== 'Consult Titan');
    // Фильтрация записей, включающая только строки, у которых значение 'Дата_BLOCK' не пустое, либо 'Приёмка' не пустое и 'Статус' равен 'проблема'
    records = records.filter(record => {
    const dateBlock = record['Дата_BLOCK'];

    const acceptanceDate = record['Приёмка'];


    if (!dateBlock && !acceptanceDate) {
        return false;
    }

const DATE_RANGE = {
    start: new Date(2024, 9, 19, 0, 0, 0), // Октябрь (месяц 9, так как месяцы в JavaScript начинаются с 0)
    end: new Date(2024, 9, 30, 23, 59, 59) // Октябрь
};

const isValidDateRange = (date) => {
    if (!(date instanceof Date)) {
        
        return false;
    }
    
    return date >= DATE_RANGE.start && date <= DATE_RANGE.end;
};

    const validDateBlock = dateBlock ? isValidDateRange(dateBlock) : false;
    const validAcceptanceDate = acceptanceDate ? isValidDateRange(acceptanceDate) : false;

    return validDateBlock || validAcceptanceDate;
});

    const paidColumn = 'Выплачено';
    try {
        const priceColumn = 'Цена';
        const bankColumn = 'Банк';
        const profitColumn = 'Профит';
        const statusColumn = 'Статус';

        const bankPrices = {};
        const bankCardCounts = {};
        const bankProfits = {};
        const bankProblemCounts = {};
        const bankBlockCounts = {};

        banks.forEach(bank => {
            bankPrices[bank] = [];
            bankCardCounts[bank] = 0;
            bankProfits[bank] = 0;
            bankProblemCounts[bank] = 0;
            bankBlockCounts[bank] = 0;
        });

        let totalBankCount = 0;
        let totalPriceSum = 0;
        let totalProfitSum = 0;
        let totalProblemCount = 0;
        let totalBlockCount = 0;
        let totalPaidForProblem = 0;
        let totalPaidForBlock = 0;

        records.forEach(record => {
            const bank = record[bankColumn];
            const price = record[priceColumn];
            const profit = record[profitColumn];
            const status = record[statusColumn];
            if (bank && bank.trim() !== '') {
                if (bankPrices[bank] !== undefined) {
                    bankPrices[bank].push(price || 0);
                    bankCardCounts[bank] += 1;
                    bankProfits[bank] += profit || 0;
                    totalBankCount += 1;
                    totalPriceSum += price || 0;
                    totalProfitSum += profit || 0;
                    if (status && status.toLowerCase() === 'проблема') {
                        bankProblemCounts[bank] += 1;
                        totalProblemCount += 1;
                        totalPaidForProblem += record[paidColumn] || 0;
                    }
                    if (status && status.toLowerCase() === 'блокировка') {
                        bankBlockCounts[bank] += 1;
                        totalBlockCount += 1;
                        totalPaidForBlock += record[paidColumn] || 0;
                    }
                }
            }
        });

        let averagePrice = 0;
        if (totalBankCount > 0) {
            averagePrice = totalPriceSum / totalBankCount;
        }

        let averageProfit = 0;
        if (totalBankCount > 0) {
            averageProfit = totalProfitSum / totalBankCount;
        }

        let problemPercentage = 0;
        if (totalBankCount > 0) {
            problemPercentage = (totalProblemCount / totalBankCount) * 100;
        }

        const totalProcessingCol = 1 + banks.length;
        setCellText(6, totalProcessingCol, averagePrice.toFixed(2));
        setCellText(7, totalProcessingCol, totalBankCount);
        setCellText(8, totalProcessingCol, averageProfit.toFixed(2));
        setCellText(10, totalProcessingCol, totalProfitSum.toFixed(2));
        setCellText(12, totalProcessingCol, (problemPercentage / 100).toFixed(4));
        let blockPercentage = 0;
        if (totalBankCount > 0) {
            blockPercentage = (totalBlockCount / totalBankCount) * 100;
        }
        setCellText(16, totalProcessingCol, (blockPercentage / 100).toFixed(4));
        setCellText(14, totalProcessingCol, totalProblemCount);
        setCellText(18, totalProcessingCol, totalBlockCount);
        const totalProcessingColLetter = String.fromCharCode(65 + totalProcessingCol);
        setCellText(20, totalProcessingCol, `=${totalProcessingColLetter}8*${totalProcessingColLetter}7`);
        
        let sumFormula = "=";
        for (let col = 1; col <= banks.length; col++) {
            const colLetter = String.fromCharCode(65 + col);
            sumFormula += `${colLetter}23`;
            if (col < banks.length) {
                sumFormula += "+";
            }
        }
        setCellText(22, totalProcessingCol, sumFormula);

        setCellText(13, totalProcessingCol, totalPaidForProblem.toFixed(2));
        setCellText(17, totalProcessingCol, totalPaidForBlock.toFixed(2));

        let totalCostsSumFormula = "=";
        for (let col = 1; col <= banks.length; col++) {
            const colLetter = String.fromCharCode(65 + col);
            totalCostsSumFormula += `${colLetter}25`;
            if (col < banks.length) {
                totalCostsSumFormula += "+";
            }
        }
        setCellText(24, totalProcessingCol, totalCostsSumFormula);

        let db1SumFormula = "=";
        for (let col = 1; col <= banks.length; col++) {
            const colLetter = String.fromCharCode(65 + col);
            db1SumFormula += `${colLetter}26`;
            if (col < banks.length) {
                db1SumFormula += "+";
            }
        }
        setCellText(25, totalProcessingCol, db1SumFormula);

        const averagePrices = banks.map(bank => {
            const prices = bankPrices[bank];
            if (prices.length > 0) {
                const total = prices.reduce((sum, price) => sum + price, 0);
                return total / prices.length;
            }
            return 0;
        });

        const averageProfits = banks.map(bank => {
            if (bankCardCounts[bank] > 0) {
                return bankProfits[bank] / bankCardCounts[bank];
            }
            return 0;
        });

        const problemPercentages = banks.map(bank => {
            if (bankCardCounts[bank] > 0) {
                return (bankProblemCounts[bank] / bankCardCounts[bank]) * 100;
            }
            return 0;
        });

        const blockPercentages = banks.map(bank => {
            if (bankCardCounts[bank] > 0) {
                return (bankBlockCounts[bank] / bankCardCounts[bank]) * 100;
            }
            return 0;
        });

        return { averagePrices, bankCardCounts, averageProfits, problemPercentages, blockPercentages, bankProblemCounts, bankBlockCounts };
    } catch (error) {
        console.error('Ошибка при вычислении метрик:', error);
        return { averagePrices: [], bankCardCounts: {}, averageProfits: [], problemPercentages: [], blockPercentages: [], bankProblemCounts: {}, bankBlockCounts: {} };
    }
}

function populateBanks(banks, averagePrices, bankCardCounts, averageProfits, problemPercentages, blockPercentages, bankProblemCounts, bankBlockCounts) {

    banks.forEach((bank, index) => {
        const row = 3;
        const col = 1 + index;

        setCellText(row, col, bank);
        setCellText(6, col, averagePrices[index].toFixed(2));
        setCellText(7, col, bankCardCounts[bank]);
        setCellText(8, col, averageProfits[index].toFixed(2));
        setCellText(12, col, (problemPercentages[index] / 100).toFixed(4));
        setCellText(14, col, bankProblemCounts[bank]);
        setCellText(16, col, (blockPercentages[index] / 100).toFixed(4));
        setCellText(18, col, bankBlockCounts[bank]);
    });

    banks.forEach((bank, index) => {
        const col = 1 + index;
        const colLetter = String.fromCharCode(65 + col);

        setCellText(10, col, `=${colLetter}9*${colLetter}8`);
        setCellText(13, col, `=${colLetter}7*${colLetter}15`);
        setCellText(17, col, `=ЕСЛИОШИБКА(${colLetter}11*${colLetter}17)`);
        setCellText(20, col, `=${colLetter}8*${colLetter}7`);
        setCellText(22, col, `=(${colLetter}11-${colLetter}18)*${colLetter}22`);
        setCellText(24, col, `=СУММ(${colLetter}14+${colLetter}18+${colLetter}21+${colLetter}23)`);
        setCellText(25, col, `=СУММ(${colLetter}11-${colLetter}25)`);
    });

    banks.forEach((bank, index) => {
        const col = 1 + index;
        setCellBackgroundColor(0, col, '#d9ead3');
        setCellBackgroundColor(25, col, '#d9ead3');
    });

    const totalProcessingCol = 1 + banks.length;
    setCellText(0, totalProcessingCol, 'итого Процессинг');
    setCellStyle(0, totalProcessingCol, 'font-bold', true);
    setCellBackgroundColor(0, totalProcessingCol, '#d9ead3');
    setCellBackgroundColor(25, totalProcessingCol, '#d9ead3');
}

async function initializeWidget(records) {
    initializeSpreadsheet();

    const uniqueBanks = await fetchUniqueBanks(records);
    const { averagePrices, bankCardCounts, averageProfits, problemPercentages, blockPercentages, bankProblemCounts, bankBlockCounts } = await calculateMetrics(records, uniqueBanks);
    populateBanks(uniqueBanks, averagePrices, bankCardCounts, averageProfits, problemPercentages, blockPercentages, bankProblemCounts, bankBlockCounts);
    uniqueBanks.forEach((bank, index) => {
        setColumnWidth(2 + index, 150);
    });
}

grist.ready({
    columns: ['Банк', 'Цена', 'Профит', 'Статус', 'Выплачено', 'Дроповод', 'Дата_BLOCK', 'Приёмка', 'Метод', 'Площадка'],
    requiredAccess: 'read table'
});

grist.onRecords(function(records, mappings) {
    let mappedRecords = grist.mapColumnNames(records);

// Период дат для фильтрации
const DATE_RANGE = {
    start: new Date(2024, 9, 19, 0, 0, 0), // Октябрь (месяц 9, так как месяцы в JavaScript начинаются с 0)
    end: new Date(2024, 9, 30, 23, 59, 59) // Октябрь
};

// Проверка, входит ли дата в заданный диапазон
const isValidDateRange = (date) => {
    if (!(date instanceof Date)) {
        
        return false;
    }
    
    return date >= DATE_RANGE.start && date <= DATE_RANGE.end;
};

// Кастомные фильтры
const FILTERS = {
    bank: ['Россельхозбанк', 'Совкомбанк'], // Замените на нужные значения банков
    dropovod: ['Мария Металлист', 'Гриша'], // Замените на нужные значения дроповодов
    status: ['проблема', 'блокировка'], // Замените на нужные значения статусов
    method: ['Переводы', 'Наличка'], // Замените на нужные значения методов
    platform: ['Gate'] // Замените на нужные значения площадок
};

const applyCustomFilters = (record) => {
    const bankFilter = !FILTERS.bank.length || FILTERS.bank.includes(record['Банк']);
    const dropovodFilter = !FILTERS.dropovod.length || FILTERS.dropovod.includes(record['Дроповод']);
    const statusFilter = !FILTERS.status.length || FILTERS.status.includes(record['Статус']);
    const methodFilter = !FILTERS.method.length || FILTERS.method.includes(record['Метод']);
    const platformFilter = !FILTERS.platform.length || FILTERS.platform.includes(record['Площадка']);

    return bankFilter && dropovodFilter && statusFilter && methodFilter && platformFilter;
};
    if (mappedRecords) {
        // Фильтрация записей, исключая строки с 'Дроповод' равным 'Consult Kayos' или 'Consult Titan'
        mappedRecords = mappedRecords.filter(record => record['Дроповод'] !== 'Consult Kayos' && record['Дроповод'] !== 'Consult Titan');
        // Фильтрация записей, включающая только строки, у которых значение 'Дата_BLOCK' не пустое, либо 'Приёмка' не пустое и 'Статус' равен 'проблема'
        mappedRecords = mappedRecords.filter(record => {
    const dateBlockValid = record['Дата_BLOCK'] ? isValidDateRange(new Date(record['Дата_BLOCK'])) : false;
    const acceptanceDateValid = record['Приёмка'] ? isValidDateRange(new Date(record['Приёмка'])) : false;
    const dateFilter = dateBlockValid || acceptanceDateValid;
    return dateFilter && applyCustomFilters(record);
});
        initializeWidget(mappedRecords);
    } else {
        console.error("Please map all columns correctly");
    }
});
