// main.js

// Инициализация Grist и x-spreadsheet
grist.ready({
    columns: [
        'Банк', 'Цена', 'Профит', 'Статус', 'Выплачено', 'Дроповод', 'Дата_BLOCK',
        'Приёмка', 'Метод', 'Площадка', 'Бухгалтерия_Дата', 'Бухгалтерия_Проект',
        'Бухгалтерия_Операция', 'Бухгалтерия_Сумма', 'Бухгалтерия_Курс_Usd_Rub'
    ],
    requiredAccess: 'read table'
});

const s = x_spreadsheet("#x-spreadsheet-demo");

let mappedRecords = [];
const fieldsForDB1 = [
    'Банк', 'Цена', 'Профит', 'Статус', 'Выплачено', 'Дроповод', 'Дата_BLOCK',
    'Приёмка', 'Метод', 'Площадка'
];
const fieldsForDB2 = [
    'Бухгалтерия_Дата', 'Бухгалтерия_Проект', 'Бухгалтерия_Операция',
    'Бухгалтерия_Сумма', 'Бухгалтерия_Курс_Usd_Rub'
];

let exchangeRate = null;


grist.onRecords(records => {
    mappedRecords = grist.mapColumnNames(records);
    exchangeRate = getExchangeRate(mappedRecords); // Извлекаем курс из всех записей
    populateFilterOptions(); // Вызов из UI.js
    setupFilterEventListeners(); // Регистрация обработчиков событий
    applyFiltersAndRender();
});

function getExchangeRate(records) {
    for (let record of records) {
        if (record['Бухгалтерия_Курс_Usd_Rub']) {
            return parseFloat(record['Бухгалтерия_Курс_Usd_Rub']);
        }
    }
    return null;
}

function applyFiltersAndRender() {
    const filteredRecordsDB1 = mappedRecords.filter(applyCustomFiltersDB1);
    const filteredRecordsDB2 = mappedRecords.filter(applyCustomFiltersDB2);

    s.loadData([
        { name: 'DB', rows: {} },
        {
            name: 'Таблица для DB 1',
            rows: createRowsData(filteredRecordsDB1, fieldsForDB1)
        },
        {
            name: 'Таблица для DB 2',
            rows: createRowsData(filteredRecordsDB2, fieldsForDB2)
        }
    ]);
    s.reRender();

    initializeSpreadsheetDB();
    renderUniqueBankValues(filteredRecordsDB1, filteredRecordsDB2);
}

// Функция создания данных для таблицы
const createRowsData = (records, fields) => ({
    len: records.length,
    ...Object.fromEntries(records.map((record, rowIndex) => [
        rowIndex,
        {
            cells: Object.fromEntries(fields.map((field, colIndex) => [
                colIndex,
                { text: record[field] != null ? (Array.isArray(record[field]) ? record[field].join(', ') : String(record[field])) : '' }
            ]))
        }
    ]))
});

// Функция регистрации обработчиков событий для фильтров
function setupFilterEventListeners() {
    $('.selectpicker').on('changed.bs.select', function(e, clickedIndex, isSelected, previousValue) {
        const selectId = $(this).attr('id');
        let selectedValues = $(this).val() || [];

        if (clickedIndex === 0) {
            if (isSelected) {
                $(this).selectpicker('selectAll');
                $(this).selectpicker('deselect', 'all');
                selectedValues = $(this).val() || [];
            } else {
                $(this).selectpicker('deselectAll');
                selectedValues = [];
            }
        }

        const filterValue = selectedValues.filter(v => v !== 'all');

        switch (selectId) {
            case 'bank-filter':
                FILTERS_DB1.bank = filterValue;
                break;
            case 'dropovod-filter':
                FILTERS_DB1.dropovod = filterValue;
                break;
            case 'status-filter':
                FILTERS_DB1.status = filterValue;
                break;
            case 'method-filter':
                FILTERS_DB1.method = filterValue;
                break;
            case 'platform-filter':
                FILTERS_DB1.platform = filterValue;
                break;
            case 'project-filter':
                FILTERS_DB2.project = filterValue;
                break;
            case 'operation-filter':
                FILTERS_DB2.operation = filterValue;
                break;
        }

        applyFiltersAndRender();
    });

    $('#in-work-filter').on('change', function() {
        FILTERS_DB1.inWork = $(this).is(':checked');
        applyFiltersAndRender();
    });

    if (!window.flatpickrInstance) {
        window.flatpickrInstance = flatpickr("#date-range", {
            mode: "range",
            dateFormat: "d.m.Y",
            onClose: function(selectedDates) {
                DATE_RANGE.start = selectedDates[0] || null;
                DATE_RANGE.end = selectedDates[1] || null;
                applyFiltersAndRender();
            }
        });
    }
}
