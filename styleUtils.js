// styleUtils.js

// Function to set the background color of a specific cell
function setCellBackgroundColor(row, col, color, sheetIndex = 0) {
    setCellStyle(row, col, 'bgcolor', color, sheetIndex);
}

// Function to set the width of a specific column
function setColumnWidth(colIndex, width, sheetIndex = 0) {
    const dataProxy = s.datas[sheetIndex];
    dataProxy.cols.setWidth(colIndex, width);
    s.reRender();
}

// Function to set the height of a specific row
function setRowHeight(rowIndex, height, sheetIndex = 0) {
    const dataProxy = s.datas[sheetIndex];
    dataProxy.rows.setHeight(rowIndex, height);
    s.reRender();
}

// Function to enable text wrapping in a specific cell
function enableTextWrap(row, col, sheetIndex = 0) {
    setCellStyle(row, col, 'textwrap', true, sheetIndex);
}

function setCellText(row, col, text) {
    s.cellText(row, col, text);
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
