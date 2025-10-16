function createSheetsAndLayouts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Получение массивов
  const documentProperties = PropertiesService.getDocumentProperties();

  const basicItemsString = documentProperties.getProperty('listOfBasicItems');
  const recipesString = documentProperties.getProperty('listOfRecipes');
  const costsString = documentProperties.getProperty('listOfCosts');
  const listOfRepetitionsString = documentProperties.getProperty('listOfRepetitions');
  const maxRepetitionsCountString = documentProperties.getProperty('maxRepetitionsCount');

  const listOfBasicItems = JSON.parse(basicItemsString);
  const listOfRecipes = JSON.parse(recipesString);
  const listOfCosts = JSON.parse(costsString);
  const listOfRepetitions = JSON.parse(listOfRepetitionsString);
  const maxRepetitionsCount = JSON.parse(maxRepetitionsCountString);

  // получение массива всех рецептов и максимального уровня крафтов
  const allRecipes = listOfRecipes.flatMap(recipe => recipe.list);
  const maxLevel = Math.max(...allRecipes.map(recipe => recipe.level));

  // создание листа FAQ
  const shFAQ = ss.insertSheet("FAQ");
  shFAQ.getRange("C1:D1").setValues([['Бонус к продаже, %', 0]]);
  shFAQ.getRange("C2").setValue('Учитывать в себестоимости возможность продажи ингредиентов');
  shFAQ.getRange("D2").insertCheckboxes();
  shFAQ.getRange("C3:D3").setValues([['Цена 1 ед. энергии, руб', '2,2']]);
  shFAQ.getRange(4, 3, listOfRecipes.length).setValues(listOfRecipes.map(recipe => [recipe.name]));
  const dropDownMenuRange = shFAQ.getRange(4, 4, listOfRecipes.length);
  dropDownMenuRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({length: maxLevel}, (_, i) => i + 1)).build());
  dropDownMenuRange.setValue(1);

  //создание Тех. листа
  const shTech = ss.insertSheet("Тех. лист");
  shTech.getRange("A1:B1").setValues([['Предмет','Цена продажи']]);
  shTech.getRange(2, 1, listOfBasicItems.length, 2).setValues(listOfBasicItems.map(b => [b, listOfCosts.find(c => c.name === b).cost]));
  shTech.getRange("D1:E1").setValues([['Предмет','Цена продажи']]);
  shTech.getRange(1, 6, 1, maxLevel).setValues([Array.from({length: maxLevel}, (_, i) => `${i + 1} ур`)]);
  var repetitionsCounter = listOfRepetitions.map(item => ({name: item, count: 0}));
  shTech.getRange(2, 4, allRecipes.length, maxLevel + 2).setValues(allRecipes.map(a => {
    var result = a.result;
    if(repetitionsCounter.some(item => item.name === result)) {
      const item = repetitionsCounter.find(item => item.name === result);
      result = `${result}${item.count}`;
      item.count++;
    }
    return [
      result,
      listOfCosts.find(l => l.name === a.result).cost,
      ...Array.from({length: a.level - 1}, () => '0'),
      a.count,
      ...Array.from({length: maxLevel - a.level}, () => null)
    ];
  }));
  shTech.getRange(1, maxLevel + 7).setValue("Предмет");
  shTech.getRange(2, maxLevel + 7, listOfRepetitions.length).setValues(listOfRepetitions.map(item => [item]));
  var header = shTech.getRange(1, maxLevel + 8, 1, maxRepetitionsCount);
  header.merge();
  header.setValue("Ссылки на себестоимость");
  shTech.getRange(1, maxLevel + maxRepetitionsCount + 8).setValue("Мин. себес.");
  shTech.getRange(1, maxLevel + maxRepetitionsCount + 9).setValue("Ваша цена");
  shTech.getRange(1, maxLevel + maxRepetitionsCount + 10).setValue("Результат");
  shTech.getRange(1, maxLevel + maxRepetitionsCount + 11).setValue("Ссылка на мин. себес.");
  shTech.getRange(1, maxLevel + maxRepetitionsCount + 12).setValue("Ссылка на Ваша цена");
  shTech.getRange(1, maxLevel + maxRepetitionsCount + 13).setValue("Итоговая ссылка");
  shTech.getRange(2, maxLevel + maxRepetitionsCount + 8, listOfRepetitions.length).setFormula(`=IF(SUMPRODUCT(--ISNUMBER(R[0]C[-${maxRepetitionsCount}]:R[0]C[-1]));MIN(FILTER(R[0]C[-${maxRepetitionsCount}]:R[0]C[-1];ISNUMBER(R[0]C[-${maxRepetitionsCount}]:R[0]C[-1])));IF(COUNTIF(R[0]C[-${maxRepetitionsCount}]:R[0]C[-1];"Н/Д")>0;"Н/Д";"Низкий уровень"))`);
  shTech.getRange(2, maxLevel + maxRepetitionsCount + 10, listOfRepetitions.length).setFormula(`=IF(ISNUMBER(R[0]C[-1]); R[0]C[-1]; IF(AND(FAQ!D$2;ISNUMBER(R[0]C[-2]));MAX(VLOOKUP(R[0]C[${-maxRepetitionsCount-3}]&"1";'Тех. лист'!D$2:E$` + (allRecipes.length + 1) + `; 2; FALSE)*(1+(FAQ!D$1)/100); R[0]C[-2]); R[0]C[-2]))`);
  shTech.getRange(2, maxLevel + maxRepetitionsCount + 13, listOfRepetitions.length).setFormula(`=IF(ISBlANK(R[0]C[-1]);R[0]C[-2];R[0]C[-1])`);
  shTech.getRange(1, 6, allRecipes.length + 1, maxLevel).setBorder(true, true, true, true, false, false);

  // создание листа Цены
  const shCosts = ss.insertSheet("Цены");
  shCosts.insertColumns(1, (listOfRecipes.length + 1) * 4 - shCosts.getMaxColumns());
  header = shCosts.getRange("A1:C1");
  header.merge();
  header.setValue('Базовые компоненты');
  shCosts.getRange("A2:C2").setValues([['Предмет', 'Средняя цена', 'Ваша цена']]);
  shCosts.getRange("A1:C2").setHorizontalAlignment("center");
  shCosts.getRange(3, 1, listOfBasicItems.length).setValues(listOfBasicItems.map(item => [item]));
  shCosts.getRange(3, 4, listOfBasicItems.length).setFormula(`=IF(ISNUMBER(R[0]C[-1]); R[0]C[-1]; IF(AND(FAQ!D$2;ISNUMBER(R[0]C[-2]));MAX(VLOOKUP(R[0]C[-3];'Тех. лист'!A$2:B$` + (listOfBasicItems.length + 1) + `; 2; FALSE)*(1+(FAQ!D$1)/100); R[0]C[-2]); R[0]C[-2]))`);
  shCosts.getRange(1, 1, listOfBasicItems.length + 2, 3).setBorder(true, true, true, true, false, false);
  listOfRecipes.forEach(branch => {
    var column = (listOfRecipes.indexOf(branch) + 1) * 4 + 1;
    header = shCosts.getRange(1, column, 1, 3);
    header.merge();
    header.setValue(branch.name);
    shCosts.getRange(2, column, 1, 3).setValues([['Предмет', 'Себестоимость', 'Ваша цена']]);
    shCosts.getRange(1, column, 2, 3).setHorizontalAlignment("center");
    shCosts.getRange(1, column, branch.list.length + 2, 3).setBorder(true, true, true, true, false, false);
  });

  repetitionsCounter = listOfRepetitions.map(item => ({name: item, count: 0}));

  //создание листа Прокачка
  const shImp = ss.insertSheet("Прокачка");
  shImp.insertColumns(1, listOfRecipes.length * 4 - shImp.getMaxColumns() - 1);
  listOfRecipes.forEach(branch => {
    var column = listOfRecipes.indexOf(branch) * 4 + 1;
    header = shImp.getRange(1, column, 1, 3);
    header.merge();
    header.setValue(branch.name);
    shImp.getRange(2, column, 1, 3).setValues([['Предмет', 'Продажа', 'Цена']]);
    shImp.getRange(1, column, 2, 3).setHorizontalAlignment("center");
    shImp.getRange(1, column, branch.list.length + 2, 3).setBorder(true, true, true, true, false, false);
  });

  // проход по каждому рецепту
  listOfRecipes.forEach(branch => {

    // создание нового листа
    const sh = ss.insertSheet(branch.name);

    // создание переменных для навигации
    const costC = (listOfRecipes.indexOf(branch) + 1) * 4 + 1;
    var costR = 3;
    var baseCell = sh.getRange("A1");
    var middleIndex = Math.floor(branch.list.length / 2);
    branch.list.forEach(recipe => {

      // переход вправо при достижении половины рецептов
      if(branch.list.indexOf(recipe) === middleIndex) baseCell = sh.getRange("F1");
      // создани перменных для навигации
      const baseC = baseCell.getColumn();
      const baseR = baseCell.getRow();
      var row = baseR;
      const ingredientsCount = recipe.ingredients.length;

      // заполнение заголовков
      const header = sh.getRange(row, baseC, 1, 4);
      header.merge();
      header.setFormula(createLink(shCosts, shCosts.getRange(costR, costC), recipe.result));
      row++;
      sh.getRange(row, baseC, 1, 4).setValues([['Предмет', 'Штук', 'Цена за штуку', 'Подсчет для крафта']]);
      sh.getRange(baseR, baseC, 2, 4).setHorizontalAlignment("center");

      // создание заготовок формул
      const craftingCountIndex = sh.getRange(baseR + ingredientsCount + 3, baseC + 1).getA1Notation();
      const wantToCraftIndex = sh.getRange(baseR + ingredientsCount + 3, baseC + 3).getA1Notation();
      const countOfFullCrafts = `CEILING(${wantToCraftIndex}/${craftingCountIndex})`;
      const costRange = sh.getRange(baseR + 2, baseC + 2, ingredientsCount).getA1Notation();
      
      // заполнение информации о ингридиентах
      row++;
      recipe.ingredients.forEach(ingredient => {
        var ingredientR, ingredientC;
        if(!listOfRepetitions.includes(ingredient.name)) {
          if(listOfBasicItems.some(l => l === ingredient.name)) {
            ingredientC = 1;
            ingredientR = listOfBasicItems.findIndex(l => l === ingredient.name) + 3; 
          } else {
            listOfRecipes.forEach(branch1 => {
              if(branch1.list.some(b => b.result === ingredient.name)) {
                ingredientC = (listOfRecipes.indexOf(branch1) + 1) * 4 + 1;
                ingredientR = branch1.list.findIndex(b => b.result === ingredient.name) + 3; 
              } 
            });
          }
          sh.getRange(row, baseC).setFormula(createLink(shCosts, shCosts.getRange(ingredientR, ingredientC), ingredient.name));
          sh.getRange(row, baseC + 2).setFormula(`=Цены!${shCosts.getRange(ingredientR, ingredientC + 3).getA1Notation()}`);
        } else {
          sh.getRange(row, baseC).setFormula(`=HYPERLINK("#gid=${shCosts.getSheetId()}&range="&'Тех. лист'!${shTech.getRange(listOfRepetitions.indexOf(ingredient.name) + 2, maxLevel + maxRepetitionsCount + 13).getA1Notation()};"${ingredient.name}")`);
          sh.getRange(row, baseC + 2).setFormula(`='Тех. лист'!${shTech.getRange(listOfRepetitions.indexOf(ingredient.name) + 2, maxLevel + maxRepetitionsCount + 10).getA1Notation()}`);
        }
        sh.getRange(row, baseC + 1).setValue(ingredient.count);
        sh.getRange(row, baseC + 3).setFormula(createFormula(craftingCountIndex, costRange, `${countOfFullCrafts}*R[0]C[-2]`));
        row++;
      });

      // заполнение строки с энергией
      sh.getRange(row, baseC, 1, 2).setValues([["Энергия", recipe.energy]]);
      sh.getRange(row, baseC + 2).setFormula("=FAQ!D3");
      sh.getRange(row, baseC + 3).setFormula(createFormula(craftingCountIndex, costRange, `${countOfFullCrafts}*R[0]C[-2]`));

      // заполнение первой нижней строки
      row++;
      sh.getRange(row, baseC, 1, 4).setValues([['Штук за крафт', null, 'Хочу сделать', 0]]);
      var name = recipe.result;
      if(listOfRepetitions.includes(recipe.result)) name += `${repetitionsCounter.find(item => item.name === recipe.result).count}`;
      sh.getRange(row, baseC + 1).setFormula(`VLOOKUP("${name}";'Тех. лист'!${shTech.getRange(2, 4, allRecipes.length, maxLevel + 2).getA1Notation()};'FAQ'!D${listOfRecipes.indexOf(branch) + 4}+2;false)`);

      // заполнение второй нижней строки
      row++;
      sh.getRange(row, baseC, 1, 3).setValues([['Себестоимость', null, 'Цена крафта']]);
      var sumOfCraft = `SUMPRODUCT(${sh.getRange(baseR + 2, baseC + 1, ingredientsCount + 1).getA1Notation()};${sh.getRange(baseR + 2, baseC + 2, ingredientsCount + 1).getA1Notation()})`;
      sh.getRange(row, baseC + 1).setFormula(createFormula(craftingCountIndex, costRange, `CEILING(${sumOfCraft}/${craftingCountIndex})`))
      sh.getRange(row, baseC + 3).setFormula(createFormula(craftingCountIndex, costRange, `${countOfFullCrafts}*${sumOfCraft}`));
      sh.getRange(baseR, baseC, ingredientsCount + 5, 4).setBorder(true, true, true, true, false, false);

      // заполнение строки о крафте на листе "Цены"
      shCosts.getRange(costR, costC).setFormula(createLink(sh, baseCell, recipe.result));
      shCosts.getRange(costR, costC + 1).setFormula(`='${branch.name}'!${sh.getRange(baseR + ingredientsCount + 4, baseC + 1).getA1Notation()}`);
      if(listOfRepetitions.includes(recipe.result)) {
        const item = repetitionsCounter.find(item => item.name === recipe.result);
        shTech.getRange(listOfRepetitions.indexOf(recipe.result) + 2, maxLevel + item.count + 8).setFormula(`='Цены'!${shCosts.getRange(costR, costC + 1).getA1Notation()}`);
        item.count++;
        name = `"${name}"`;
      } else name = `R[0]C[-3]`;
      shCosts.getRange(costR, costC + 3).setFormula(`=IF(ISNUMBER(R[0]C[-1]); R[0]C[-1]; IF(AND(FAQ!D2;ISNUMBER(R[0]C[-2]));MAX(VLOOKUP(${name};'Тех. лист'!D2:E` + (allRecipes.length + 1) + `; 2; FALSE)*(1+(FAQ!D1)/100); R[0]C[-2]); R[0]C[-2]))`);

      // заполнение листа прокачка
      shImp.getRange(costR, costC - 4).setFormula(createLink(shCosts, shCosts.getRange(costR, costC), recipe.result))
      const cell = `'Цены'!${shCosts.getRange(costR, costC + 1).getA1Notation()}` 
      const countCell = `'${branch.name}'!${sh.getRange(row - 1, baseC + 1).getA1Notation()}`
      sumOfCraft = `SUMPRODUCT('${branch.name}'!${sh.getRange(baseR + 2, baseC + 1, ingredientsCount + 1).getA1Notation()};'${branch.name}'!${sh.getRange(baseR + 2, baseC + 2, ingredientsCount + 1).getA1Notation()})`;
      shImp.getRange(costR, costC - 2).setFormula(`IFERROR(IF(OR(${cell}="Н/Д";${cell}="Низкий уровень");"Недоступно";CEILING(IF(ISNUMBER(R[0]C[-1]);(${sumOfCraft}-${countCell}*R[0]C[-1])/${recipe.energy/10};(${sumOfCraft}-${countCell}*VLOOKUP(${name};'Тех. лист'!D2:E` + (allRecipes.length + 1) + `; 2; FALSE)*IF(FAQ!D2;1+(FAQ!D1)/100;1))/${recipe.energy/10});0,1));)`)

      // подготовка к выводу следующего рецепта
      costR++;
      baseCell = sh.getRange(baseR + ingredientsCount + 6, baseC);
    });
  });

  for(var i = 0; i < listOfRepetitions.length; i++) {
    var cells = shTech.getRange(i + 2, maxLevel + 8, 1, maxRepetitionsCount).getFormulas().flat().filter(cell => cell !== "").map(cell => cell.slice(1));
    const decrementedCells = decrementColumnInArray(cells);
    const decrementedCellsWithQuotes = '{"' + decrementedCells.map(cell => cell.slice(7)).join('"; "') + '"}';
    const incrementedCellsWithoutQuotes = '{' + incrementColumnInArray(cells).join('; ') + '}';
    cells = '{' + cells.join('; ') + '}';
    shTech.getRange(i + 2, maxLevel + maxRepetitionsCount + 9).setFormula(`=IFERROR(INDEX(${incrementedCellsWithoutQuotes}; MATCH(MIN(FILTER(${incrementedCellsWithoutQuotes};ISNUMBER(${incrementedCellsWithoutQuotes})));${incrementedCellsWithoutQuotes};0));)`);
    shTech.getRange(i + 2, maxLevel + maxRepetitionsCount + 11).setFormula(`=IF(SUMPRODUCT(--ISNUMBER(${cells}))>0;INDEX(${decrementedCellsWithQuotes}; MATCH(MIN(FILTER(${cells};ISNUMBER(${cells})));${cells};0));IF(COUNTIF(${cells};"Н/Д")>0;INDEX(${decrementedCellsWithQuotes};MATCH("Н/Д";${cells};0));"${decrementedCells[0].slice(7)}"))`);
    shTech.getRange(i + 2, maxLevel + maxRepetitionsCount + 12).setFormula(`=IFERROR(INDEX(${decrementedCellsWithQuotes}; MATCH(MIN(FILTER(${incrementedCellsWithoutQuotes};ISNUMBER(${incrementedCellsWithoutQuotes})));${incrementedCellsWithoutQuotes};0));)`);
  }

  // форматирование всех листов
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const range = sheet.getDataRange();
    range.setFontFamily("Nunito");
    range.setFontSize(12);
    range.setFontColor("black");
  });
  for(var i = 0; i < listOfRecipes.length + 1; i++) {
    shCosts.getRange(1, (i + 1) * 4, 1000).setFontColor("white");
  }
 
  function createFormula(craftingCountIndex, costRange, formula) {
    return `IF(OR(ISBLANK(${craftingCountIndex});COUNTIF(${costRange}; "Н/Д")>0); "Н/Д"; IF(OR(${craftingCountIndex}=0; COUNTIF(${costRange}; "Низкий уровень")>0); "Низкий уровень"; ${formula}))`;
  }

  function createLink(sh, cellRange, text) {
    return `=HYPERLINK("#gid=${sh.getSheetId()}&range=${cellRange.getA1Notation()}";"${text}")`;
  }

  function incrementColumnInArray(cells) {
    return cells.map(cell => {
      const [_, column, rowNumber] = cell.slice(7).match(/^([A-Z]+)(\d+)$/);
      const newColumn = numberToColumn(columnToNumber(column) + 1);
      return "'Цены'!" + newColumn + rowNumber;
    });
  }

  function decrementColumnInArray(cells) {
    return cells.map(cell => {
      const [_, column, rowNumber] = cell.slice(7).match(/^([A-Z]+)(\d+)$/);
      const newColumn = numberToColumn(columnToNumber(column) - 1);
      return "'Цены'!" + newColumn + rowNumber;
    });
  }

  function columnToNumber(column) {
    var result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
  }

  function numberToColumn(number) {
    var result = '';
    while (number > 0) {
      number--;
      result = String.fromCharCode((number % 26) + 'A'.charCodeAt(0)) + result;
      number = Math.floor(number / 26);
    }
    return result;
  }
}


