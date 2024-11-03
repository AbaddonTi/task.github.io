// UI.js

function populateFilterOptions() {
    const uniqueValues = (field) => {
        const values = mappedRecords.flatMap(r => {
            let value = r[field];
            if (Array.isArray(value)) {
                return value;
            }
            return [value];
        }).filter(Boolean);
        return [...new Set(values)];
    };

    const filterFields = {
        'bank-filter': 'Банк',
        'dropovod-filter': 'Дроповод',
        'status-filter': 'Статус',
        'method-filter': 'Метод',
        'platform-filter': 'Площадка',
        'project-filter': 'Бухгалтерия_Проект',
        'operation-filter': 'Бухгалтерия_Операция'
    };

    Object.entries(filterFields).forEach(([selectId, fieldName]) => {
        const select = document.getElementById(selectId);
        const options = uniqueValues(fieldName).map(value => `<option value="${value}">${value}</option>`).join('');
        select.innerHTML = `
            <option value="all">Все</option>
            ${options}
        `;
    });

    $('.selectpicker').selectpicker('refresh');
}
