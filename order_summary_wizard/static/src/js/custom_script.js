document.addEventListener('DOMContentLoaded', function () {
    // Dynamically adjust the table container height
    const tableContainer = document.querySelector('.table-container');
    if (tableContainer) {
        tableContainer.style.height = 'calc(100vh - 150px)'; // Adjust dynamically based on viewport
        tableContainer.style.overflowY = 'auto'; // Ensure scrolling for overflow
    }

    // Ensure the tree table occupies full width
    const table = document.querySelector('.tree-table');
    if (table) {
        table.style.width = '100%'; // Full width
        table.style.tableLayout = 'fixed'; // Equal column widths
    }
});