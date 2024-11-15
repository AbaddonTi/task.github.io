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

// UI.js

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
                if (selectedDates[1]) {
                    DATE_RANGE.end = new Date(selectedDates[1]);
                    DATE_RANGE.end.setHours(23, 59, 59, 999);
                } else {
                    DATE_RANGE.end = null;
                }
                applyFiltersAndRender();
            }
        });
    }
}

