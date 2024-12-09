"use strict";

if (!Array.from) {  
        Array.from = function (arrayLike, mapFn, thisArg) {
        const arr = [];  
         
        if (typeof arrayLike[Symbol.iterator] === 'function') {
            const iterator = arrayLike[Symbol.iterator]();  
            let index = 0;  
            let result = iterator.next();  
             
            while (!result.done) {  
                const value = result.value;  
                 
                arr.push(mapFn ? mapFn.call(thisArg, value, index) : value);  
                result = iterator.next();  
                index++;  
            }
        }
        return arr;  
    };
}
 
if (!Object.fromEntries) {  
         Object.fromEntries = function (entries) {
        const obj = {};  
         
        for (const [key, value] of entries) {
            obj[key] = value;  
        }
        return obj;  
    };
}
 
const ROWS = Math.floor(window.innerHeight / 40);  
const COLUMNS = Math.floor(window.innerWidth / 100);  
const sheetData = {};  
 
let activeSheetName = null;  

console.log("Скрипт завантажений");  
 
  const generateEmptySheetData = () => 
 
Array.from({ length: ROWS }, () => 
Object.fromEntries(
Array.from({ length: COLUMNS }, (_, j) => 
[String.fromCharCode(65 + j), ""])));
 
 const createTable = (sheetName) => {
     
    if (document.querySelector(`.sheet[data-sheet-name="${sheetName}"]`)) {
        console.warn(`Таблиця для ліста "${sheetName}" вже існує!`);
        return;
    }
    console.log(`Створюємо таблицю для ліста: ${sheetName}`);
    const sheetsContainer = document.getElementById("sheetsContainer");
    const sheetDiv = document.createElement("div");
    sheetDiv.className = "sheet";
    sheetDiv.dataset.sheetName = sheetName;
    const table = document.createElement("table");
    table.className = "table table-bordered text-center";
    const thead = document.createElement("thead");
    const tbody = document.createElement("tbody");
     
    thead.innerHTML = `<tr>${Array.from({ length: COLUMNS }, (_, i) => `<th>${String.fromCharCode(65 + i)}</th>`).join('')}</tr>`;
     
    tbody.innerHTML = Array.from({ length: ROWS }, (_, i) => `<tr>${Array.from({ length: COLUMNS }, (_, j) => {
        const col = String.fromCharCode(65 + j);
         
        return `<td><input type="text" data-row="${i}" data-col="${col}" value="${sheetData[sheetName][i][col] || ""}" /></td>`;
    }).join('')}</tr>`).join('');
     
    table.appendChild(thead);
    table.appendChild(tbody);
     
    sheetDiv.appendChild(table);
     
    sheetsContainer.appendChild(sheetDiv);
     
    if (sheetName === activeSheetName)
        sheetDiv.classList.add("active");
     
    sheetDiv.querySelectorAll("input").forEach(input => input.addEventListener("input", handleInput));
};
 
 const addSheet = () => {
    const sheetCount = Object.keys(sheetData).length + 1;
    const newSheetName = `Sheet ${sheetCount}`;
    sheetData[newSheetName] = generateEmptySheetData();
    console.log(`Добавляємо новий лист: ${newSheetName}`);
    createTable(newSheetName);
    addSheetTab(newSheetName);
    switchSheet(newSheetName);
};
 
 const addSheetTab = (sheetName) => {
    var _a;
     
    if (document.querySelector(`.sheet-tab[data-sheet-name="${sheetName}"]`)) {
        console.warn(`Вкладка для ліста "${sheetName}" вже існує!`);
        return;
    }
    const tab = document.createElement("div");
    tab.className = "sheet-tab";
    tab.textContent = sheetName;
    tab.dataset.sheetName = sheetName;
    tab.addEventListener("click", () => switchSheet(sheetName));
    (_a = document.getElementById("tabs")) === null || _a === void 0 ? void 0 : _a.appendChild(tab);
};
 
 const switchSheet = (sheetName) => {
    var _a, _b;
    console.log(`Переключаемось на ліст: ${sheetName}`);
     
    document.querySelectorAll(".sheet, .sheet-tab").forEach(el => el.classList.remove("active"));
     
    (_a = document.querySelector(`.sheet[data-sheet-name="${sheetName}"]`)) === null || _a === void 0 ? void 0 : _a.classList.add("active");
    (_b = document.querySelector(`.sheet-tab[data-sheet-name="${sheetName}"]`)) === null || _b === void 0 ? void 0 : _b.classList.add("active");
    activeSheetName = sheetName;
};
 
  const saveToFile = () => {
    if (validateSheetData()) {
        const dataStr = JSON.stringify({ sheetData, activeSheetName });
        const blob = new Blob([dataStr], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "sheetData.json";
        a.click();
        URL.revokeObjectURL(url);
        console.log("Дані збережені у файл.");
    }
    else {
        console.error("Помилка: Дані не валідні и не можуть бути збережені.");
    }
};
 
 const loadFromFile = (file) => {
    const reader = new FileReader();
    reader.onload = (event) => {
        var _a;
        try {
            const result = (_a = event.target) === null || _a === void 0 ? void 0 : _a.result;
            const { sheetData: loadedSheetData, activeSheetName: loadedActiveSheetName } = JSON.parse(result);
            if (validateSheetData(loadedSheetData)) {
                 
                for (const sheetName in loadedSheetData) {
                    const loadedSheet = loadedSheetData[sheetName];
                     
                    if (!sheetData[sheetName]) {
                        sheetData[sheetName] = generateEmptySheetData();
                        createTable(sheetName);
                        addSheetTab(sheetName);
                    }
                     
                    activeSheetName = sheetName;
                    switchSheet(sheetName);
                    const currentSheet = sheetData[sheetName];
                     
                    loadedSheet.forEach((row, rowIndex) => {
                        Object.entries(row).forEach(([col, value]) => {
                            if (currentSheet[rowIndex][col] !== value) {
                                setCellValueInSheet(sheetName, `${col}${rowIndex + 1}`, value);
                            }
                        });
                    });
                }
                 
                activeSheetName = loadedActiveSheetName || null;
                if (activeSheetName) {
                    switchSheet(activeSheetName);
                }
                console.log("Дані завантаженні з файлу.");
            }
            else {
                console.error("Помилка: Дані в файлі не валидні.");
            }
        }
        catch (e) {
            console.error("Помилка: Не вдалось розпарити дані з файлу.");
        }
    };
    reader.readAsText(file);
};
 
 const validateSheetData = (data = sheetData) => {
    return Object.values(data).every(sheet => sheet.every(row => Object.values(row).every(value => typeof value === 'string')));
};
 
 const handleInput = (event) => {
    const target = event.target;
    const { row, col } = target.dataset;
    if (row !== undefined && col !== undefined) {
        sheetData[activeSheetName][parseInt(row)][col] = target.value;
        console.log(`Оновлено значення: ліст=${activeSheetName}, рядок=${row}, колонка=${col}, значення="${target.value}"`);
         
    }
};
 
 const handleFormulaInput = (event) => {
    const formulaInput = event.target;
    const formula = formulaInput.value.trim();
     
     
    const match = formula.match(/^(.+?)\((\w+\d+)\)\s*=\s*(.+?)\((\w+\d+)\)\s*([\+\-\*/%])\s*(.+?)\((\w+\d+)\)$/);
    if (match) {
         
        const [, targetSheet, targetCell, sheet1, cell1, operator, sheet2, cell2] = match;
         
        const value1 = getCellValueFromSheet(sheet1.trim(), cell1);
        const value2 = getCellValueFromSheet(sheet2.trim(), cell2);
        if (value1 !== null && value2 !== null) {
            let result;
             
            switch (operator) {
                case '+':
                    result = parseFloat(value1) + parseFloat(value2);
                    break;
                case '-':
                    result = parseFloat(value1) - parseFloat(value2);
                    break;
                case '*':
                    result = parseFloat(value1) * parseFloat(value2);
                    break;
                case '/':
                    result = parseFloat(value1) / parseFloat(value2);
                    break;
                case '%':
                    result = (parseFloat(value1) * parseFloat(value2)) / 100;
                    break;
                default:
                    console.error("Помилка: Операція не підтримується.");
                    return;
            }
             
            setCellValueInSheet(targetSheet.trim(), targetCell, result.toString());
            console.log(`Формула виконана: ${targetSheet}(${targetCell}) = ${result}`);
        }
        else {
            console.error("Помилка: Хибні дані в комірках.");
        }
    }
    else {
        console.error("Помилка: Хибний формат формули.");
    }
};
 
 const getCellValueFromSheet = (sheetName, cell) => {
    var _a, _b;
    const [col, row] = [cell.charAt(0), parseInt(cell.slice(1)) - 1];
    return ((_b = (_a = sheetData[sheetName]) === null || _a === void 0 ? void 0 : _a[row]) === null || _b === void 0 ? void 0 : _b[col]) || null;
};
 
 const setCellValueInSheet = (sheetName, cell, value) => {
    const [col, row] = [cell.charAt(0), parseInt(cell.slice(1)) - 1];
    sheetData[sheetName][row][col] = value;
    if (sheetName === activeSheetName) {
        const input = document.querySelector(`.sheet.active input[data-row="${row}"][data-col="${col}"]`);
        if (input)
            input.value = value;
    }
};
 
 document.addEventListener("DOMContentLoaded", () => {
    var _a, _b, _c;
    const formulaInput = document.getElementById("formulaInput");
    if (formulaInput) {
        formulaInput.addEventListener("keydown", (event) => {
            console.log(`event.key=${event.key}`);
            if (event.key === "Enter") {
                handleFormulaInput(event);
            }
        });
    }
    if (Object.keys(sheetData).length === 0)
        addSheet();
    else
        switchSheet(activeSheetName || "Sheet 1");
    (_a = document.getElementById("addSheetBtn")) === null || _a === void 0 ? void 0 : _a.addEventListener("click", addSheet);
    (_b = document.getElementById("saveDataBtn")) === null || _b === void 0 ? void 0 : _b.addEventListener("click", saveToFile);
    const loadDataInput = document.getElementById("loadDataInput");
    (_c = document.getElementById("loadDataBtn")) === null || _c === void 0 ? void 0 : _c.addEventListener("click", () => loadDataInput.click());
    loadDataInput.addEventListener("change", (event) => {
        var _a;
        const file = (_a = event.target.files) === null || _a === void 0 ? void 0 : _a[0];
        if (file) {
            loadFromFile(file);
        }
    });
    console.log("Добавляємо обробник для вводу формули");
});
