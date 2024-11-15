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

const applyFilters = (record, filters, applyDate = true, dateField = 'Дата_BLOCK') => {
    const { bank, dropovod, status, method, platform } = filters;

    const fieldChecks = [
        { field: 'Банк', filter: bank },
        { field: 'Дроповод', filter: dropovod },
        { field: 'Статус', filter: status },
        { field: 'Метод', filter: method },
        { field: 'Площадка', filter: platform }
    ];

    for (const { field, filter } of fieldChecks) {
        if (filter.length === 0) continue; 
        const recordValue = record[field];
        const values = Array.isArray(recordValue) ? recordValue : [recordValue];
        const hasMatch = values.some(value => filter.includes(value));
        if (!hasMatch) return false;
    }

    if (applyDate) {
        if (!record[dateField]) return false;
        if (!isDateInRange(record[dateField])) return false;
    }

    return true;
};

const isExcludedDropovod = (dropovodValue) => {
    if (Array.isArray(dropovodValue)) {
        return dropovodValue.some(dropovod => excludedDropovods.includes(dropovod));
    }
    return excludedDropovods.includes(dropovodValue);
};

const applyCustomFiltersDB1 = (record) => {
    const status = Array.isArray(record['Статус']) ? record['Статус'] : [record['Статус']];
    const isInWork = FILTERS_DB1.inWork && status.includes('в работе');

    if (isExcludedDropovod(record['Дроповод'])) return false;

    if (isInWork) {
        return applyFilters(record, FILTERS_DB1, false);
    } else {
        const applicableDateField = status.includes('проблема') ? 'Приёмка' : 'Дата_BLOCK';
        return applyFilters(record, FILTERS_DB1, true, applicableDateField);
    }
};

const applyCustomFiltersDB2 = (record) => {
    const dateFilter = isDateInRange(record['Бухгалтерия_Дата']);
    if (!dateFilter) return false;

    const { project, operation } = FILTERS_DB2;

    const projectFilter = project.length === 0 || (Array.isArray(record['Бухгалтерия_Проект']) 
        ? record['Бухгалтерия_Проект'].some(p => project.includes(p)) 
        : project.includes(record['Бухгалтерия_Проект']));

    const operationFilter = operation.length === 0 || (Array.isArray(record['Бухгалтерия_Операция']) 
        ? record['Бухгалтерия_Операция'].some(o => operation.includes(o)) 
        : operation.includes(record['Бухгалтерия_Операция']));

    return projectFilter && operationFilter;
};
