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
        { row: 25, text: 'DB 1' },
        { row: 26, text: 'Рентабельность' }
    ].forEach(({ row, text }) => setCellText(row, 0, text, 0));
}

/********** DYNAMIC RENDERING FUNCTIONS **********/
function renderUniqueBankValues(filteredRecordsDB1, filteredRecordsDB2) {
    const bankData = processBankRecords(filteredRecordsDB1);
    const { colIndex, methodColIndexes } = renderBankData(bankData); 
    renderOperationTypes(filteredRecordsDB2, { colIndex, methodColIndexes }, exchangeRate); 
    s.reRender();
}
