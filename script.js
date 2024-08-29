// script.js
grist.ready({ requiredAccess: 'read table' });

const board = document.getElementById('board');

// Function to create a new column (list)
function createColumn(title) {
    const column = document.createElement('div');
    column.classList.add('column');
    const columnTitle = document.createElement('div');
    columnTitle.classList.add('column-title');
    columnTitle.innerText = title;
    column.appendChild(columnTitle);

    // Add drop zone for cards
    column.addEventListener('dragover', function(e) {
        e.preventDefault();
        const draggingCard = document.querySelector('.dragging');
        column.appendChild(draggingCard);
    });

    board.appendChild(column);
    return column;
}

// Function to create a new card (task)
function createCard(text) {
    const card = document.createElement('div');
    card.classList.add('card');
    card.innerText = text;

    // Make the card draggable
    card.setAttribute('draggable', true);
    card.addEventListener('dragstart', function() {
        card.classList.add('dragging');
    });
    card.addEventListener('dragend', function() {
        card.classList.remove('dragging');
    });

    return card;
}

// Event listener for Grist data
grist.onRecords(function(records) {
    board.innerHTML = '';  // Clear existing board
    const columnsMap = {};

    // Create columns first
    records.forEach(record => {
        if (!columnsMap[record.Column]) {
            columnsMap[record.Column] = createColumn(record.Column);
        }
    });

    // Add cards to their respective columns
    records.forEach(record => {
        const column = columnsMap[record.Column];
        const card = createCard(record.Task);
        column.appendChild(card);
    });
});
