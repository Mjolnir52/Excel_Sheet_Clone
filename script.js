// script.js - Excel-like Spreadsheet Functionality

// Global variables
let currentSheet = 0;
let sheets = [{
    name: "Sheet1",
    data: {},
    styles: {}
}];

// Initialize the spreadsheet
document.addEventListener('DOMContentLoaded', function() {
    initializeGrid();
    renderSheet(currentSheet);
    setupEventListeners();
});

// Initialize the grid structure
function initializeGrid() {
    const grid = document.getElementById('spreadsheetGrid');
    
    // Clear existing content
    grid.innerHTML = '';
    
    // Add column headers (A-Z)
    for (let col = 1; col <= 26; col++) {
        const header = document.createElement('div');
        header.className = 'header-cell';
        header.style.gridColumn = col + 1;
        header.style.gridRow = 1;
        header.textContent = String.fromCharCode(64 + col);
        grid.appendChild(header);
    }
    
    // Add row headers (1-100) and cells
    for (let row = 1; row <= 100; row++) {
        // Row number
        const rowHeader = document.createElement('div');
        rowHeader.className = 'row-header';
        rowHeader.style.gridColumn = 1;
        rowHeader.style.gridRow = row + 1;
        rowHeader.textContent = row;
        grid.appendChild(rowHeader);
        
        // Cells for this row
        for (let col = 1; col <= 26; col++) {
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.style.gridColumn = col + 1;
            cell.style.gridRow = row + 1;
            cell.setAttribute('contenteditable', 'true');
            const cellId = `${String.fromCharCode(64 + col)}${row}`;
            cell.setAttribute('id', cellId);
            cell.setAttribute('data-cell', cellId);
            grid.appendChild(cell);
        }
    }
}

// Render a sheet
function renderSheet(sheetIndex) {
    const sheet = sheets[sheetIndex];
    const cells = document.querySelectorAll('.cell');
    
    cells.forEach(cell => {
        const cellId = cell.getAttribute('id');
        // Set cell value
        cell.textContent = sheet.data[cellId] || '';
        
        // Apply cell styles
        if (sheet.styles[cellId]) {
            Object.entries(sheet.styles[cellId]).forEach(([property, value]) => {
                cell.style[property] = value;
            });
        } else {
            // Reset styles if none exist
            cell.style.fontWeight = '';
            cell.style.fontStyle = '';
            cell.style.textDecoration = '';
            cell.style.textAlign = '';
            cell.style.color = '';
            cell.style.backgroundColor = '';
            cell.style.fontFamily = '';
            cell.style.fontSize = '';
        }
    });
    
    // Update active tab
    document.querySelectorAll('.sheet-tab').forEach((tab, index) => {
        tab.classList.toggle('active', index === sheetIndex);
    });
}

// Set up event listeners
function setupEventListeners() {
    // Cell input handling
    document.querySelectorAll('.cell').forEach(cell => {
        cell.addEventListener('input', handleCellInput);
        cell.addEventListener('focus', handleCellFocus);
        cell.addEventListener('blur', handleCellBlur);
        cell.addEventListener('keydown', handleCellKeyDown);
    });
    
    // Toolbar event listeners
    document.getElementById('fontSelect').addEventListener('change', function() {
        applyFormatting('fontFamily', this.value);
    });
    
    document.getElementById('fontSize').addEventListener('change', function() {
        applyFormatting('fontSize', this.value + 'px');
    });
    
    // Import/export
    document.getElementById('importFile').addEventListener('change', handleFileImport);
}

// Handle cell input
function handleCellInput(e) {
    const cell = e.target;
    const cellId = cell.getAttribute('id');
    let value = cell.textContent.trim();
    
    // Check if the value is a formula (starts with =)
    if (value.startsWith('=')) {
        value = evaluateFormula(value.substring(1), cellId);
    }
    
    // Update the sheet data
    sheets[currentSheet].data[cellId] = value;
    
    // Update dependent cells
    updateDependentCells(cellId);
}

// Handle cell focus
function handleCellFocus(e) {
    const cell = e.target;
    cell.classList.add('active-cell');
    updateFormulaBar(cell.getAttribute('id'));
}

// Handle cell blur
function handleCellBlur(e) {
    const cell = e.target;
    cell.classList.remove('active-cell');
}

// Handle cell key events
function handleCellKeyDown(e) {
    const cell = e.target;
    const cellId = cell.getAttribute('id');
    const col = cellId.match(/[A-Z]+/)[0];
    const row = parseInt(cellId.match(/\d+/)[0]);
    
    let newCell;
    
    switch(e.key) {
        case 'Enter':
            e.preventDefault();
            newCell = document.getElementById(`${col}${row + 1}`);
            if (newCell) newCell.focus();
            break;
        case 'Tab':
            e.preventDefault();
            if (e.shiftKey) {
                // Move left
                const prevCol = String.fromCharCode(col.charCodeAt(0) - 1);
                if (prevCol >= 'A') {
                    newCell = document.getElementById(`${prevCol}${row}`);
                }
            } else {
                // Move right
                const nextCol = String.fromCharCode(col.charCodeAt(0) + 1);
                if (nextCol <= 'Z') {
                    newCell = document.getElementById(`${nextCol}${row}`);
                }
            }
            if (newCell) newCell.focus();
            break;
        case 'ArrowUp':
            newCell = document.getElementById(`${col}${Math.max(1, row - 1)}`);
            if (newCell) newCell.focus();
            break;
        case 'ArrowDown':
            newCell = document.getElementById(`${col}${row + 1}`);
            if (newCell) newCell.focus();
            break;
        case 'ArrowLeft':
            const leftCol = String.fromCharCode(col.charCodeAt(0) - 1);
            if (leftCol >= 'A') {
                newCell = document.getElementById(`${leftCol}${row}`);
                if (newCell) newCell.focus();
            }
            break;
        case 'ArrowRight':
            const rightCol = String.fromCharCode(col.charCodeAt(0) + 1);
            if (rightCol <= 'Z') {
                newCell = document.getElementById(`${rightCol}${row}`);
                if (newCell) newCell.focus();
            }
            break;
    }
}

// Evaluate formulas
function evaluateFormula(formula, cellId) {
    try {
        // Handle basic arithmetic
        if (/^[\d+\-*/.() ]+$/.test(formula)) {
            return eval(formula);
        }
        
        // Handle cell references (e.g., =A1+B2)
        const cellRefPattern = /[A-Z]+\d+/g;
        let evaluatedFormula = formula;
        let match;
        
        while ((match = cellRefPattern.exec(formula)) !== null) {
            const refCellId = match[0];
            const refValue = sheets[currentSheet].data[refCellId] || 0;
            evaluatedFormula = evaluatedFormula.replace(new RegExp(refCellId, 'g'), isNaN(refValue) ? 0 : parseFloat(refValue));
        }
        
        // Handle basic functions
        if (evaluatedFormula.startsWith('SUM(')) {
            const range = evaluatedFormula.match(/SUM\(([^)]+)\)/)[1];
            const [start, end] = range.split(':');
            return sumRange(start, end);
        }
        
        // Evaluate the final expression
        return eval(evaluatedFormula);
    } catch (error) {
        console.error('Formula error:', error);
        return '#ERROR!';
    }
}

// Sum a range of cells
function sumRange(startCell, endCell) {
    if (!startCell || !endCell) return 0;
    
    const startCol = startCell.match(/[A-Z]+/)[0];
    const startRow = parseInt(startCell.match(/\d+/)[0]);
    const endCol = endCell.match(/[A-Z]+/)[0];
    const endRow = parseInt(endCell.match(/\d+/)[0]);
    
    let sum = 0;
    
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        for (let row = startRow; row <= endRow; row++) {
            const cellId = `${String.fromCharCode(col)}${row}`;
            const value = parseFloat(sheets[currentSheet].data[cellId]) || 0;
            sum += value;
        }
    }
    
    return sum;
}

// Update cells that depend on the changed cell
function updateDependentCells(cellId) {
    // In a real implementation, we would track dependencies
    // For now, we'll just re-evaluate all cells that start with =
    document.querySelectorAll('.cell').forEach(cell => {
        const cellContent = sheets[currentSheet].data[cell.getAttribute('id')] || '';
        if (typeof cellContent === 'string' && cellContent.startsWith('=')) {
            const newValue = evaluateFormula(cellContent.substring(1), cell.getAttribute('id'));
            cell.textContent = newValue;
            sheets[currentSheet].data[cell.getAttribute('id')] = newValue;
        }
    });
}

// Formatting functions
function formate(command) {
    document.execCommand(command, false, null);
    saveSelectionStyles();
}

function applyColor(color, type) {
    const property = type === 'color' ? 'color' : 'backgroundColor';
    applyFormatting(property, color);
}

function applyFormatting(property, value) {
    const selection = window.getSelection();
    if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        const cell = range.startContainer.parentElement.closest('.cell');
        
        if (cell) {
            cell.style[property] = value;
            const cellId = cell.getAttribute('id');
            
            // Save the style to the sheet
            if (!sheets[currentSheet].styles[cellId]) {
                sheets[currentSheet].styles[cellId] = {};
            }
            sheets[currentSheet].styles[cellId][property] = value;
        }
    }
}

function saveSelectionStyles() {
    const selection = window.getSelection();
    if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        const cell = range.startContainer.parentElement.closest('.cell');
        
        if (cell) {
            const cellId = cell.getAttribute('id');
            const styles = {
                fontWeight: cell.style.fontWeight,
                fontStyle: cell.style.fontStyle,
                textDecoration: cell.style.textDecoration,
                textAlign: cell.style.textAlign,
                color: cell.style.color,
                backgroundColor: cell.style.backgroundColor,
                fontFamily: cell.style.fontFamily,
                fontSize: cell.style.fontSize
            };
            
            sheets[currentSheet].styles[cellId] = styles;
        }
    }
}

// Sheet management
function addNewSheet() {
    const newSheetNumber = sheets.length + 1;
    const newSheet = {
        name: `Sheet${newSheetNumber}`,
        data: {},
        styles: {}
    };
    
    sheets.push(newSheet);
    
    // Add tab
    const tab = document.createElement('div');
    tab.className = 'sheet-tab';
    tab.textContent = newSheet.name;
    tab.addEventListener('click', () => switchSheet(sheets.length - 1));
    
    document.getElementById('sheetTabs').appendChild(tab);
    
    // Switch to new sheet
    switchSheet(sheets.length - 1);
}

function switchSheet(sheetIndex) {
    currentSheet = sheetIndex;
    renderSheet(sheetIndex);
}

// Import/Export functions
function exportToCSV() {
    const sheet = sheets[currentSheet];
    let csv = '';
    
    // Get max row and column
    let maxRow = 1;
    let maxCol = 1;
    
    Object.keys(sheet.data).forEach(cellId => {
        const col = cellId.match(/[A-Z]+/)[0];
        const row = parseInt(cellId.match(/\d+/)[0]);
        
        maxCol = Math.max(maxCol, col.charCodeAt(0) - 64);
        maxRow = Math.max(maxRow, row);
    });
    
    // Generate CSV
    for (let row = 1; row <= maxRow; row++) {
        const rowData = [];
        for (let col = 1; col <= maxCol; col++) {
            const cellId = `${String.fromCharCode(64 + col)}${row}`;
            rowData.push(`"${sheet.data[cellId] || ''}"`);
        }
        csv += rowData.join(',') + '\n';
    }
    
    // Download
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${sheet.name}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function handleFileImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const contents = e.target.result;
        importCSV(contents);
    };
    reader.readAsText(file);
}

function importCSV(csvData) {
    const rows = csvData.split('\n');
    const sheet = sheets[currentSheet];
    
    // Clear existing data
    sheet.data = {};
    sheet.styles = {};
    
    rows.forEach((row, rowIndex) => {
        const columns = row.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
        
        columns.forEach((value, colIndex) => {
            // Remove surrounding quotes if present
            value = value.replace(/^"|"$/g, '').trim();
            
            if (value) {
                const cellId = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
                sheet.data[cellId] = value;
                
                // Update the cell in the UI
                const cell = document.getElementById(cellId);
                if (cell) {
                    cell.textContent = value;
                }
            }
        });
    });
    
    // Reset file input
    document.getElementById('importFile').value = '';
}

// Update formula bar (placeholder - would need a formula bar element in HTML)
function updateFormulaBar(cellId) {
    // In a real implementation, this would update a formula bar display
    console.log(`Selected cell: ${cellId}, Value: ${sheets[currentSheet].data[cellId] || ''}`);
}