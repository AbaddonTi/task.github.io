// script.js

const s = x_spreadsheet('#x-spreadsheet-demo');

grist.ready({
    columns: ['Банк', 'Цена', 'Профит', 'Статус', 'Выплачено', 'Дроповод', 'Дата_BLOCK', 'Приёмка', 'Метод', 'Площадка', 'Бухгалетрия_Дата', 'Бухгалетрия_Проект', 'Бухгалетрия_Операция', 'Бухгалетрия_Сумма', 'Бухгалетрия_Курс_Usd_Rub'],
    requiredAccess: 'read table'
});

grist.onRecords(function(records, mappings) {
    let mappedRecords = grist.mapColumnNames(records);

    let DATE_RANGE = {
        start: null,
        end: null
    };

    const isValidDateRange = (date) => {
        if (!(date instanceof Date) || DATE_RANGE.start === null || DATE_RANGE.end === null) {
            return true; // No date range is set, allow all records
        }
        return date >= DATE_RANGE.start && date <= DATE_RANGE.end;
    };

    const FILTERS = {
        bank: [],
        dropovod: [], 
        status: [], 
        method: [], 
        platform: [] 
    };

    const applyCustomFilters = (record) => {
        const bankFilter = !FILTERS.bank.length || FILTERS.bank.includes(record['Банк']);
        const dropovodFilter = !FILTERS.dropovod.length || FILTERS.dropovod.includes(record['Дроповод']);
        const statusFilter = !FILTERS.status.length || FILTERS.status.includes(record['Статус']);
        const methodFilter = !FILTERS.method.length || FILTERS.method.includes(record['Метод']);
        const platformFilter = !FILTERS.platform.length || FILTERS.platform.includes(record['Площадка']);

        return bankFilter && dropovodFilter && statusFilter && methodFilter && platformFilter;
    };
    
    const clearSpreadsheet = () => {
        const sheetData = s.datas[0];
        const maxRows = sheetData.rows.len;
        const maxCols = sheetData.cols.len;
        
        for (let row = 0; row < maxRows; row++) {
            for (let col = 0; col < maxCols; col++) {
                const cell = sheetData.getCell(row, col);
                if (cell && (cell.text || cell.style)) {
                    s.cellText(row, col, '');
                    setCellStyle(row, col, 'bgcolor', null); // Reset the background color to default
                }
            }
        }
        s.reRender();
    };

    const filterAndRender = () => {
        clearSpreadsheet(); // Clear the spreadsheet before re-rendering
        let filteredRecords = mappedRecords;
        if (filteredRecords) {
            filteredRecords = filteredRecords.filter(record => record['Дроповод'] !== 'Consult Kayos' && record['Дроповод'] !== 'Consult Titan');
            filteredRecords = filteredRecords.filter(record => {
                const dateBlockValid = record['Дата_BLOCK'] ? isValidDateRange(new Date(record['Дата_BLOCK'])) : false;
                const acceptanceDateValid = record['Приёмка'] ? isValidDateRange(new Date(record['Приёмка'])) : false;
                const dateFilter = dateBlockValid || acceptanceDateValid;
                return dateFilter && applyCustomFilters(record);
            });
            initializeWidget(filteredRecords);
        } else {
            console.error("Please map all columns correctly");
        }
    };

    // Initial rendering without date restrictions
    if (mappedRecords) {
        initializeWidget(mappedRecords);
    } else {
        console.error("Please map all columns correctly");
    }

    // Add date range picker with dropdown and preset options
    const dateRangeButton = document.createElement('button');
    dateRangeButton.innerText = 'Выбрать диапазон дат';
    dateRangeButton.classList.add('date-range-button');

    const dropdownMenu = document.createElement('div');
    dropdownMenu.classList.add('dropdown-menu');
    dropdownMenu.style.display = 'none';

    const presets = [
        { label: 'Сегодня', action: () => setDateRangePreset('today') },
        { label: 'Вчера', action: () => setDateRangePreset('yesterday') },
        { label: 'Последние 7 дней', action: () => setDateRangePreset('last7days') },
        { label: 'Последние 30 дней', action: () => setDateRangePreset('last30days') },
        { label: 'Этот месяц', action: () => setDateRangePreset('thisMonth') },
        { label: 'Прошлый месяц', action: () => setDateRangePreset('lastMonth') },
    ];

    presets.forEach(preset => {
        const button = document.createElement('button');
        button.innerText = preset.label;
        button.onclick = () => {
            preset.action();
            dropdownMenu.style.display = 'none';
        };
        dropdownMenu.appendChild(button);
    });

    // Custom date range input
    const customStartDateInput = document.createElement('input');
    customStartDateInput.type = 'date';
    customStartDateInput.placeholder = 'Начальная дата';

    const customEndDateInput = document.createElement('input');
    customEndDateInput.type = 'date';
    customEndDateInput.placeholder = 'Конечная дата';

    const applyCustomDateButton = document.createElement('button');
    applyCustomDateButton.innerText = 'Применить';
    applyCustomDateButton.onclick = () => {
        const startDate = new Date(customStartDateInput.value);
        const endDate = new Date(customEndDateInput.value);

        if (isNaN(startDate) || isNaN(endDate)) {
            alert('Неверный формат даты. Пожалуйста, введите корректные даты.');
        } else {
            DATE_RANGE.start = startDate;
            DATE_RANGE.end = endDate;
            filterAndRender();
            dropdownMenu.style.display = 'none';
        }
    };

    dropdownMenu.appendChild(customStartDateInput);
    dropdownMenu.appendChild(customEndDateInput);
    dropdownMenu.appendChild(applyCustomDateButton);

    document.body.appendChild(dateRangeButton);
    dateRangeButton.parentElement.appendChild(dropdownMenu);

    dateRangeButton.onclick = () => {
        dropdownMenu.style.display = dropdownMenu.style.display === 'none' ? 'block' : 'none';
    };

    const setDateRangePreset = (preset) => {
        const today = new Date();
        let startDate, endDate;

        switch (preset) {
            case 'today':
                startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
                endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
                break;
            case 'yesterday':
                startDate = new Date(today);
                startDate.setDate(today.getDate() - 1);
                startDate.setHours(0, 0, 0, 0);
                endDate = new Date(today);
                endDate.setDate(today.getDate() - 1);
                endDate.setHours(23, 59, 59, 999);
                break;
            case 'last7days':
                startDate = new Date(today);
                startDate.setDate(today.getDate() - 7);
                startDate.setHours(0, 0, 0, 0);
                endDate = new Date(today);
                endDate.setHours(23, 59, 59, 999);
                break;
            case 'last30days':
                startDate = new Date(today);
                startDate.setDate(today.getDate() - 30);
                startDate.setHours(0, 0, 0, 0);
                endDate = new Date(today);
                endDate.setHours(23, 59, 59, 999);
                break;
            case 'thisMonth':
                startDate = new Date(today.getFullYear(), today.getMonth(), 1);
                endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59);
                break;
            case 'lastMonth':
                startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
                endDate = new Date(today.getFullYear(), today.getMonth(), 0, 23, 59, 59);
                break;
        }

        DATE_RANGE.start = startDate;
        DATE_RANGE.end = endDate;
        filterAndRender();
    };
});
 
