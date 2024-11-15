// filters.js

const DATE_RANGE = {
    start: null,
    end: null
};

const FILTERS_DB1 = {
    bank: [],
    dropovod: [],
    status: [],
    method: [],
    platform: [],
    inWork: false 
};

const FILTERS_DB2 = {
    project: [],
    operation: []
};

const excludedDropovods = ['Consult Kayos', 'Consult Titan'];

const parseDate = (value) => {
    if (!value) return null;
    if (value instanceof Date) return value;
    if (typeof value === 'string') {
        const [datePart, timePart = '00:00'] = value.split(' ');
        const [day, month] = datePart.split('.');
        const year = new Date().getFullYear();
        return new Date(`${year}-${month}-${day}T${timePart}:00`);
    }
    return null;
};

const isDateInRange = (dateValue) => {
    if (!DATE_RANGE.start || !DATE_RANGE.end) return true;
    const date = parseDate(dateValue);
    return date && date >= DATE_RANGE.start && date <= DATE_RANGE.end;
};

const applyCustomFiltersDB1 = (record) => {
    const isInWork = FILTERS_DB1.inWork && record['Статус'] === 'в работе';

    if (isInWork) {
        const dropovodValue = record['Дроповод'];
        return !(Array.isArray(dropovodValue) ? dropovodValue.some(dropovod => excludedDropovods.includes(dropovod)) : excludedDropovods.includes(dropovodValue));
    }

    if (!record['Дата_BLOCK'] && !record['Приёмка']) return false;

    let status = record['Статус'];
    if (!Array.isArray(status)) {
        status = [status];
    }

    const dateField = status.includes('проблема') ? 'Приёмка' : 'Дата_BLOCK';

    const dateFilter = isDateInRange(record[dateField]);

    const bankFilter = !FILTERS_DB1.bank.length || (Array.isArray(record['Банк']) ? record['Банк'].some(bank => FILTERS_DB1.bank.includes(bank)) : FILTERS_DB1.bank.includes(record['Банк']));
    const dropovodFilter = !FILTERS_DB1.dropovod.length || (Array.isArray(record['Дроповод']) ? record['Дроповод'].some(dropovod => FILTERS_DB1.dropovod.includes(dropovod)) : FILTERS_DB1.dropovod.includes(record['Дроповод']));
    const statusFilter = !FILTERS_DB1.status.length || status.some(s => FILTERS_DB1.status.includes(s));
    const methodFilter = !FILTERS_DB1.method.length || (Array.isArray(record['Метод']) ? record['Метод'].some(method => FILTERS_DB1.method.includes(method)) : FILTERS_DB1.method.includes(record['Метод']));
    const platformFilter = !FILTERS_DB1.platform.length || (Array.isArray(record['Площадка']) ? record['Площадка'].some(platform => FILTERS_DB1.platform.includes(platform)) : FILTERS_DB1.platform.includes(record['Площадка']));

    const dropovodValue = record['Дроповод'];
    const excludedDropovodFilter = !(Array.isArray(dropovodValue) ? dropovodValue.some(dropovod => excludedDropovods.includes(dropovod)) : excludedDropovods.includes(dropovodValue));

    return dateFilter && bankFilter && dropovodFilter && statusFilter && methodFilter && platformFilter && excludedDropovodFilter;
};


const applyCustomFiltersDB2 = (record) => {
    const dateFilter = isDateInRange(record['Бухгалтерия_Дата']);
    if (!dateFilter) return false;

    const projectFilter = !FILTERS_DB2.project.length || (Array.isArray(record['Бухгалтерия_Проект']) ? record['Бухгалтерия_Проект'].some(project => FILTERS_DB2.project.includes(project)) : FILTERS_DB2.project.includes(record['Бухгалтерия_Проект']));
    const operationFilter = !FILTERS_DB2.operation.length || (Array.isArray(record['Бухгалтерия_Операция']) ? record['Бухгалтерия_Операция'].some(operation => FILTERS_DB2.operation.includes(operation)) : FILTERS_DB2.operation.includes(record['Бухгалтерия_Операция']));

    return projectFilter && operationFilter;
};
