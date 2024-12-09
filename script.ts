if (!Array.from) {
  Array.from = function <T, U>(arrayLike: ArrayLike<T> | Iterable<T>, mapFn?: (v: T, k: number) => U, thisArg?: any): U[] {
    const arr: U[] = [];

    if (typeof (arrayLike as Iterable<T>)[Symbol.iterator] === 'function') {
      const iterator = (arrayLike as Iterable<T>)[Symbol.iterator]();
      let index = 0;
      let result = iterator.next();

      while (!result.done) {
        const value = result.value;

        arr.push(mapFn ? mapFn.call(thisArg, value, index) : (value as unknown as U));
        result = iterator.next();
        index++;
      }
    }
    return arr;
  };
}


if (!Object.fromEntries) {
  Object.fromEntries = function <T = any>(entries: Iterable<[string, T]>): Record<string, T> {
    const obj: Record<string, T> = {};

    for (const [key, value] of entries) {
      obj[key] = value;
    }
    return obj;
  };
}


const ROWS: number = Math.floor(window.innerHeight / 40);
const COLUMNS: number = Math.floor(window.innerWidth / 100);
const sheetData: Record<string, Array<Record<string, string>>> = {};

let activeSheetName: string | null = null;

console.log("Скрипт завантажений");

const generateEmptySheetData = (): Array<Record<string, string>> =>

  Array.from({ length: ROWS }, () =>

    Object.fromEntries(

      Array.from({ length: COLUMNS }, (_, j) =>

        [String.fromCharCode(65 + j), ""]
      )
    )
  );


const createTable = (sheetName: string): void => {

  if (document.querySelector(`.sheet[data-sheet-name="${sheetName}"]`)) {
    console.warn(`Таблиця для ліста "${sheetName}" вже існує!`);
    return;
  }

  console.log(`Створюємо таблицю для ліста: ${sheetName}`);
  const sheetsContainer = document.getElementById("sheetsContainer") as HTMLElement;

  const sheetDiv = document.createElement("div");
  sheetDiv.className = "sheet";
  sheetDiv.dataset.sheetName = sheetName;

  const table = document.createElement("table");
  table.className = "table table-bordered text-center";
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");


  thead.innerHTML = `<tr>${Array.from({ length: COLUMNS }, (_, i) => `<th>${String.fromCharCode(65 + i)}</th>`).join('')}</tr>`;


  tbody.innerHTML = Array.from({ length: ROWS }, (_, i) =>
    `<tr>${Array.from({ length: COLUMNS }, (_, j) => {

      const col = String.fromCharCode(65 + j);


      return `<td><input type="text" data-row="${i}" data-col="${col}" value="${sheetData[sheetName][i][col] || ""}" /></td>`;
    }).join('')}</tr>`
  ).join('');


  table.appendChild(thead);
  table.appendChild(tbody);

  sheetDiv.appendChild(table);

  sheetsContainer.appendChild(sheetDiv);


  if (sheetName === activeSheetName) sheetDiv.classList.add("active");


  sheetDiv.querySelectorAll("input").forEach(input => input.addEventListener("input", handleInput));
};


const addSheet = (): void => {
  const sheetCount = Object.keys(sheetData).length + 1;
  const newSheetName = `Sheet ${sheetCount}`;
  sheetData[newSheetName] = generateEmptySheetData();
  console.log(`Добавляємо новий ліст: ${newSheetName}`);
  createTable(newSheetName);
  addSheetTab(newSheetName);
  switchSheet(newSheetName);
};


const addSheetTab = (sheetName: string): void => {

  if (document.querySelector(`.sheet-tab[data-sheet-name="${sheetName}"]`)) {
    console.warn(`Вкладка для ліста "${sheetName}" вже існує!`);
    return;
  }
  const tab = document.createElement("div");
  tab.className = "sheet-tab";
  tab.textContent = sheetName;
  tab.dataset.sheetName = sheetName;
  tab.addEventListener("click", () => switchSheet(sheetName));
  document.getElementById("tabs")?.appendChild(tab);
};


const switchSheet = (sheetName: string): void => {
  console.log(`Переключаемось на лист: ${sheetName}`);


  document.querySelectorAll(".sheet, .sheet-tab").forEach(el => el.classList.remove("active"));

  
  document.querySelector(`.sheet[data-sheet-name="${sheetName}"]`)?.classList.add("active");
  document.querySelector(`.sheet-tab[data-sheet-name="${sheetName}"]`)?.classList.add("active");

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
    console.log("Дані збереженні в файл.");
  } else {
    console.error("Помилка: Дані не валідні и не можуть бути збережені.");
  }
};


const loadFromFile = (file: File) => {
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const result = event.target?.result as string;
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


          loadedSheet.forEach((row: Record<string, string>, rowIndex: number) => {
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
        console.log("Дані завантажені з файлу.");
      } else {
        console.error("Ошибка: Дані в файлі не валідні.");
      }
    } catch (e) {
      console.error("Помилка: Невдалось розпарити дані з файлу.");
    }
  };
  reader.readAsText(file);
};


const validateSheetData = (data: Record<string, Array<Record<string, string>>> = sheetData): boolean => {
  return Object.values(data).every(sheet =>
    sheet.every(row =>
      Object.values(row).every(value => typeof value === 'string')
    )
  );
};


const handleInput = (event: Event): void => {
  const target = event.target as HTMLInputElement;
  const { row, col } = target.dataset;
  if (row !== undefined && col !== undefined) {
    sheetData[activeSheetName!][parseInt(row)][col] = target.value;
    console.log(`Оновлено значення: ліст=${activeSheetName}, рядок=${row}, колонка=${col}, значення="${target.value}"`);

  }
};


const handleFormulaInput = (event: KeyboardEvent): void => {
  const formulaInput = event.target as HTMLInputElement;
  const formula = formulaInput.value.trim();



  const match = formula.match(/^(.+?)\((\w+\d+)\)\s*=\s*(.+?)\((\w+\d+)\)\s*([\+\-\*/%])\s*(.+?)\((\w+\d+)\)$/);

  if (match) {

    const [, targetSheet, targetCell, sheet1, cell1, operator, sheet2, cell2] = match;


    const value1 = getCellValueFromSheet(sheet1.trim(), cell1);
    const value2 = getCellValueFromSheet(sheet2.trim(), cell2);

    if (value1 !== null && value2 !== null) {
      let result: number;


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
          console.error("Помилка: Операція не піддтримується.");
          return;
      }


      setCellValueInSheet(targetSheet.trim(), targetCell, result.toString());
      console.log(`Формула виконана: ${targetSheet}(${targetCell}) = ${result}`);
    } else {
      console.error("Помилка: Хибні данні в комірках.");
    }
  } else {
    console.error("ПОмилка: Хибний формат формули.");
  }
};


const getCellValueFromSheet = (sheetName: string, cell: string): string | null => {
  const [col, row] = [cell.charAt(0), parseInt(cell.slice(1)) - 1];
  return sheetData[sheetName]?.[row]?.[col] || null;
};


const setCellValueInSheet = (sheetName: string, cell: string, value: string): void => {
  const [col, row] = [cell.charAt(0), parseInt(cell.slice(1)) - 1];
  sheetData[sheetName][row][col] = value;
  if (sheetName === activeSheetName) {
    const input = document.querySelector(`.sheet.active input[data-row="${row}"][data-col="${col}"]`) as HTMLInputElement;
    if (input) input.value = value;
  }
};


document.addEventListener("DOMContentLoaded", () => {
  const formulaInput = document.getElementById("formulaInput");
  if (formulaInput) {
    formulaInput.addEventListener("keydown", (event: KeyboardEvent) => {
      console.log(`event.key=${event.key}`);
      if (event.key === "Enter") {
        handleFormulaInput(event);
      }
    });
  }

  if (Object.keys(sheetData).length === 0) addSheet();
  else switchSheet(activeSheetName || "Sheet 1");

  document.getElementById("addSheetBtn")?.addEventListener("click", addSheet);
  document.getElementById("saveDataBtn")?.addEventListener("click", saveToFile);

  const loadDataInput = document.getElementById("loadDataInput") as HTMLInputElement;
  document.getElementById("loadDataBtn")?.addEventListener("click", () => loadDataInput.click());
  loadDataInput.addEventListener("change", (event) => {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      loadFromFile(file);
    }
  });

  console.log("Добавляємо обробник для вводу формули");
});
