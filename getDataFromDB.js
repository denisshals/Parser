function getDataFromAPI() {
  // получение данных о названиях предметов
  const directoryUrl = 'https://api.github.com/repos/EXBO-Studio/stalcraft-database/contents/ru/items/';
  const folders = ['bullet','drink','food','grenade','medicine','misc','other'];
  const listOfAllItems = []; 
  folders.forEach(folder => {
    const response = UrlFetchApp.fetch(directoryUrl + folder);
    const files = JSON.parse(response.getContentText());
    files.forEach(file => {
      if (file.name.endsWith('.json')) { 
        const jsonResponse = UrlFetchApp.fetch(file.download_url);
        const jsonData = JSON.parse(jsonResponse.getContentText());
        listOfAllItems.push({id: jsonData.id, name: jsonData.name.lines.ru, cost: getCost(jsonData)});
      }
    });
  });

  // получение данных о рецептах
  const recipesUrl = 'https://api.github.com/repos/EXBO-Studio/stalcraft-database/contents/ru/hideout_recipes.json';
  const response = UrlFetchApp.fetch(recipesUrl);
  const content = JSON.parse(response.getContentText()).content;
  const decodedContent = Utilities.newBlob(Utilities.base64Decode(content)).getDataAsString();
  const jsonData = JSON.parse(decodedContent);

  // создание массива с разделами и рецептами
  const listOfRecipes = jsonData.perks.map(perk => ({id: perk.id, name: perk.name.lines.ru, list: []}));

  // извлечение и сортировка крафтов по веткам
  jsonData.recipes.forEach(recipe => {
    const branchId = Object.keys(recipe.requirements.perks)[0];
    const branch = listOfRecipes.find(b => b.id === branchId);
    const result = listOfAllItems.find(item => item.id === recipe.result[0].item).name;
    const ingredients = recipe.ingredients.map(ingredient => ({
      name: listOfAllItems.find(item => item.id === ingredient.item).name, 
      count: ingredient.amount
    }));
    branch.list.push({result, count: recipe.result[0].amount,ingredients, energy: recipe.energy, level: Object.values(recipe.requirements.perks)[0]});
  });

  // сортировка крафтов по алфавиту в ветках
  listOfRecipes.forEach(branch => {
    branch.list.sort((a, b) => a.result.localeCompare(b.result));
  });

  // получение перечня базовых компонентов
  const listOfBasicItems = new Set();
  const listOfCraftingItems = new Set();

  listOfRecipes.forEach(branch => {
    branch.list.forEach(recipe => {
      listOfCraftingItems.add(recipe.result);
      recipe.ingredients.forEach(ingredient => {
        listOfBasicItems.add(ingredient.name);
      });
    });
  });
  const basicItemsArray = Array.from(listOfBasicItems).filter(item => !listOfCraftingItems.has(item)).sort();

  const listOfCosts = [];
  const combinedItems = new Set([...listOfBasicItems, ...listOfCraftingItems]);
  combinedItems.forEach(itemName => {
    const item = listOfAllItems.find(i => i.name === itemName);
    listOfCosts.push({
      name: itemName,
      cost: item.cost
    });
  });

  // создание массива повторяющихся крафтов
  const counts = [];
  listOfRecipes.forEach(branch => {
    branch.list.forEach(recipe => {
      var repetition = counts.find(item => item.name === recipe.result);
      if(repetition) repetition.count += 1;
      else counts.push({name: recipe.result, count: 1});
    });
  });
  const listOfRepetitions = counts.filter(item => item.count >= 2).map(item => item.name);
  const maxRepetitionsCount = Math.max(...counts.map(item => item.count));
  listOfRepetitions.sort((a, b) => a.localeCompare(b));

  // сохранение полученных массивов в параметры документа
  const documentProperties = PropertiesService.getDocumentProperties();
  const basicItemsString = JSON.stringify(basicItemsArray);
  const recipesString = JSON.stringify(listOfRecipes);
  const costsString = JSON.stringify(listOfCosts);
  const listOfRepetitionsString = JSON.stringify(listOfRepetitions);
  const maxRepetitionsCountString = JSON.stringify(maxRepetitionsCount);
  documentProperties.setProperty('listOfBasicItems', basicItemsString);
  documentProperties.setProperty('listOfRecipes', recipesString);
  documentProperties.setProperty('listOfCosts', costsString);
  documentProperties.setProperty('listOfRepetitions', listOfRepetitionsString);
  documentProperties.setProperty('maxRepetitionsCount', maxRepetitionsCountString);

  function getCost(data) {
    var cost = 0;
    if (!Array.isArray(data.infoBlocks)) {
      return cost; 
    }
    for (const block of data.infoBlocks) {
      if (block.type === "list") {
        if (Array.isArray(block.elements)) {
          for (const element of block.elements) {
            if (element.type === "key-value" && element.value && element.value.args && element.value.args.cost) {
              cost = parseFloat(element.value.args.cost) || 0;
              return cost;
            }
          }
        }
      }
    }
    return cost;
  }
}
