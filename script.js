// dimensions
const COLS = 26;
const ROWS = 100;

// constants
const transparent = "transparent";
const transparentBlue = "#ddddff";
const arrMatrix = 'arrMatrix';

// table components
const tHeadRow = document.getElementById("table-heading-row");
const tBody = document.getElementById("table-body");
const buttonContainer = document.getElementById('button-container');

// excel buttons
const boldBtn = document.getElementById("bold-btn");
const italicsBtn = document.getElementById("italics-btn");
const underlineBtn = document.getElementById("underline-btn");
const leftBtn = document.getElementById("left-btn");
const centerBtn = document.getElementById("center-btn");
const rightBtn = document.getElementById("right-btn");
const cutBtn = document.getElementById('cut-btn');
const copyBtn = document.getElementById('copy-btn');
const pasteBtn = document.getElementById('paste-btn');
const uploadInput = document.getElementById('upload-input');
const addSheetBtn = document.getElementById('add-sheet-btn');
const deleteSheetBtn = document.getElementById('delete-sheet-btn');
const saveSheetBtn = document.getElementById('save-sheet-btn');

// dropdown
const fontStyleDropdown = document.getElementById("font-style-dropdown");
const fontSizeDropdown = document.getElementById("font-size-dropdown");

// input tags
const bgColorInput = document.getElementById("bgColor");
const fontColorInput = document.getElementById("fontColor");

// cache
let currentCell;
let previousCell;
let cutCell; // this cutCell will store my cell data;
let lastPressBtn;
let matrix = new Array(ROWS);
let numSheets = 1; // size
let currentSheet = 1; // index
let prevSheet;
let sheets = { 1: matrix };
let isSelecting = false;

function createNewMatrix() {
    let newMatrix = new Array(ROWS);
    for (let i = 0; i < ROWS; i++) {
        newMatrix[i] = new Array(COLS).fill("");
    }
    return newMatrix;
}

// this is creating matrix for the first time
matrix = createNewMatrix();

function colGen(typeOfCell, tableRow, isInnerText, rowNumber) {
    for (let i = 0; i < COLS; i++) {
        const cell = document.createElement(typeOfCell);
        if (isInnerText) {
            cell.innerText = String.fromCharCode(65 + i);
        } else {
            cell.contentEditable = true;
            cell.addEventListener('mousedown', () => startSelection(cell));
            cell.addEventListener('mouseover', () => continueSelection(cell));
            cell.addEventListener('mouseup', () => endSelection(cell));
        }
        tableRow.appendChild(cell);
    }
}

// this is for heading
colGen("th", tHeadRow, true);

function updateObjectInMatrix() {
    // Implement the logic to update the matrix with the current cell's data
}

function setHeaderColor(colId, rowId, color) {
    // Implement the logic to set the header color
}

function downloadMatrix() {
    // Implement the logic to download the matrix
}

function uploadMatrix(event) {
    // Implement the logic to upload the matrix
}

uploadInput.addEventListener('input', uploadMatrix);

function buttonHighlighter(button, styleProperty, style) {
    // Implement the logic to highlight the button
}

function focusHandler(cell) {
    // Implement the logic to handle cell focus
}

function tableBodyGen() {
    for (let i = 0; i < ROWS; i++) {
        const row = document.createElement("tr");
        const th = document.createElement("th");
        th.innerText = i + 1;
        row.appendChild(th);
        colGen("td", row, false, i);
        tBody.appendChild(row);
    }
}

tableBodyGen();

function clearSelection() {
    document.querySelectorAll('.selected-cell').forEach(cell => {
        cell.classList.remove('selected-cell');
    });
}

function startSelection(cell) {
    isSelecting = true;
    clearSelection();
    selectCell(cell);
}

function continueSelection(cell) {
    if (isSelecting) {
        selectCell(cell);
    }
}

function endSelection(cell) {
    isSelecting = false;
    selectCell(cell);
}

function selectCell(cell) {
    cell.classList.add('selected-cell');
    const colLetter = String.fromCharCode(65 + cell.cellIndex - 1);
    currentCell = cell;
}

// Add event listeners to ensure they are interactive
boldBtn.addEventListener("click", () => {
    if (currentCell.style.fontWeight === "bold") {
        currentCell.style.fontWeight = "normal";
    } else {
        currentCell.style.fontWeight = "bold";
    }
    updateObjectInMatrix();
});

italicsBtn.addEventListener("click", () => {
    if (currentCell.style.fontStyle === "italic") {
        currentCell.style.fontStyle = "normal";
    } else {
        currentCell.style.fontStyle = "italic";
    }
    updateObjectInMatrix();
});

underlineBtn.addEventListener("click", () => {
    if (currentCell.style.textDecoration === "underline") {
        currentCell.style.textDecoration = "none";
    } else {
        currentCell.style.textDecoration = "underline";
    }
    updateObjectInMatrix();
});

leftBtn.addEventListener("click", () => {
    currentCell.style.textAlign = "left";
    updateObjectInMatrix();
});

centerBtn.addEventListener("click", () => {
    currentCell.style.textAlign = "center";
    updateObjectInMatrix();
});

rightBtn.addEventListener("click", () => {
    currentCell.style.textAlign = "right";
    updateObjectInMatrix();
});

fontStyleDropdown.addEventListener("change", () => {
    currentCell.style.fontFamily = fontStyleDropdown.value;
    updateObjectInMatrix();
});

fontSizeDropdown.addEventListener("change", () => {
    currentCell.style.fontSize = fontSizeDropdown.value;
    updateObjectInMatrix();
});

// Add event listener to table cells
document.querySelectorAll('#table-body td').forEach(cell => {
    cell.addEventListener('mousedown', () => startSelection(cell));
    cell.addEventListener('mouseover', () => continueSelection(cell));
    cell.addEventListener('mouseup', () => endSelection(cell));
});

addSheetBtn.addEventListener('click', () => {
  numSheets++;
  currentSheet = numSheets;
  const newSheetButton = document.createElement('button');
  newSheetButton.id = `sheet-${currentSheet}`;
  newSheetButton.innerText = `Sheet ${currentSheet}`;
  newSheetButton.onclick = viewSheet;
  buttonContainer.appendChild(newSheetButton);
  sheets[currentSheet] = createNewMatrix();
  tBody.innerHTML = '';
  tableBodyGen();
});

deleteSheetBtn.addEventListener('click', () => {
    if (numSheets > 1) {
        delete sheets[currentSheet];
        document.getElementById(`sheet-${currentSheet}`).remove();
        numSheets--;
        currentSheet = numSheets;
        tBody.innerHTML = '';
        tableBodyGen();
    }
});

function viewSheet(event) {
    const sheetId = event.target.id.split('-')[1];
    currentSheet = parseInt(sheetId);
    // Implement logic to switch to the selected sheet
}

// Mathematical Functions
function sum(range) {
    return range.reduce((acc, cell) => acc + (parseFloat(cell.innerText) || 0), 0);
}

function difference(range) {
    if (range.length < 2) return 0;
    return range.reduce((acc, cell) => acc - (parseFloat(cell.innerText) || 0), parseFloat(range[0].innerText) * 2);
}

function average(range) {
    const total = sum(range);
    return total / range.length;
}

function max(range) {
    return Math.max(...range.map(cell => parseFloat(cell.innerText) || 0));
}

function min(range) {
    return Math.min(...range.map(cell => parseFloat(cell.innerText) || 0));
}

function count(range) {
    return range.filter(cell => !isNaN(parseFloat(cell.innerText))).length;
}

// Data Quality Functions
function trim(cell) {
    cell.innerText = cell.innerText.trim();
}

function upper(cell) {
    cell.innerText = cell.innerText.toUpperCase();
}

function lower(cell) {
    cell.innerText = cell.innerText.toLowerCase();
}

function removeDuplicates(range) {
    const uniqueValues = new Set();
    range.forEach(cell => {
        if (uniqueValues.has(cell.innerText)) {
            cell.innerText = '';
        } else {
            uniqueValues.add(cell.innerText);
        }
    });
}

function findAndReplace(range, findText, replaceText) {
    range.forEach(cell => {
        cell.innerText = cell.innerText.replace(new RegExp(findText, 'g'), replaceText);
    });
}

function applyFunction(funcName) {
    const selectedCells = Array.from(document.querySelectorAll('.selected-cell'));
    switch (funcName) {
        case 'sum':
            currentCell.innerText = `${sum(selectedCells)}`;
            break;
        case 'difference':
            currentCell.innerText = `${difference(selectedCells)}`;
            break;
        case 'average':
            currentCell.innerText = `${average(selectedCells)}`;
            break;
        case 'max':
            currentCell.innerText = `${max(selectedCells)}`;
            break;
        case 'min':
            currentCell.innerText = `${min(selectedCells)}`;
            break;
        case 'count':
            currentCell.innerText = `${count(selectedCells)}`;
            break;
        case 'trim':
            selectedCells.forEach(cell => trim(cell));
            break;
        case 'upper':
            selectedCells.forEach(cell => upper(cell));
            break;
        case 'lower':
            selectedCells.forEach(cell => lower(cell));
            break;
        case 'removeDuplicates':
            removeDuplicates(selectedCells);
            break;
        case 'findAndReplace':
            const findText = document.getElementById('find-text').value;
            const replaceText = document.getElementById('replace-text').value;
            findAndReplace(selectedCells, findText, replaceText);
            break;
        default:
            alert("Unknown function");
    }
}