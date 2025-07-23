/**
 * @OnlyCurrentDoc
 */
const SPREADSHEET_ID = "18Hx_V3lPBQYp4xcpopQPjQ5dR7cCtaWpESb1EM92gqI"; 
const SHEET_NAME = "АвтоматическийОтчет";
const VENDISTA_API_BASE_URL = "https://api.vendista.ru:99";
const VENDISTA_API_TOKEN = "fe38f8470367451f81228617"; 

const MACHINES_PER_TRIGGER_RUN = 5; 
const DELAY_AFTER_ZAPROS_MS = 3000; 
const DELAY_BETWEEN_MACHINES_MS = 500; 
const DELAY_BETWEEN_TRIGGERS_MINUTES = 1; 
const NUM_SLOTS_CONST = 12; 
const NUM_COLUMNS_MAIN_TABLE = 8;


// --- Helper для API запросов ---
function _vendistaApiFetch(endpoint, method = 'get', token = null, payload = null, queryParams = {}) {
  const apiUrl = VENDISTA_API_BASE_URL + endpoint;
  let fullUrl = apiUrl;

  if (token) {
    queryParams.token = token;
  }

  if (Object.keys(queryParams).length > 0) {
    const queryString = Object.entries(queryParams)
      .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
      .join('&');
    fullUrl += `?${queryString}`;
  }

  const options = {
    method: method.toLowerCase(),
    muteHttpExceptions: true, 
    headers: {}
  };

  if (payload) {
    options.contentType = 'application/json'; 
    options.payload = JSON.stringify(payload);
  }
  
  Logger.log(`_vendistaApiFetch: ${method.toUpperCase()} ${fullUrl} ContentType: ${options.contentType || 'N/A'} Payload: ${payload ? JSON.stringify(payload).substring(0,300)+'...' : 'N/A'}`);

  try {
    const response = UrlFetchApp.fetch(fullUrl, options);
    const responseCode = response.getResponseCode();
    const contentText = response.getContentText();
    const logContent = contentText.length > 600 ? contentText.substring(0, 600) + "..." : contentText;
    Logger.log(`_vendistaApiFetch Response: Code ${responseCode}, Content: ${logContent}`);

    if (responseCode >= 200 && responseCode < 300) {
      if (!contentText || contentText.trim() === "") { 
            if (responseCode === 204) return { success: true, message: "Operation successful (204 No Content).", raw: contentText, item: null, items: [] };
            return { success: true, message: "Operation successful with empty response body.", raw: contentText, item: null, items: [] };
      }
      try { 
        const jsonData = JSON.parse(contentText);
        if (jsonData.success === false && jsonData.error) {
             Logger.log(`_vendistaApiFetch: API returned success=false. Error: ${jsonData.error}`);
             return { error: `API Error: ${jsonData.error}`, details: jsonData, success: false, item: jsonData.item || null, items: jsonData.items || [] };
        }
        if (jsonData.success === undefined) {
            if (jsonData.item || jsonData.items !== undefined || responseCode === 200 || responseCode === 201 || responseCode === 202) {
                 jsonData.success = true;
            } else {
                Logger.log(`_vendistaApiFetch: Success field undefined, no item/items, code ${responseCode}. Assuming success for now.`);
                jsonData.success = true; 
            }
        }
        return jsonData;
      } catch (e) {
        if (contentText.toLowerCase().includes("ok") || contentText.toLowerCase().includes("success")) {
           return { success: true, message: "Operation successful, non-JSON 'ok/success' response.", raw: contentText, item: null, items: [] };
        }
        Logger.log(`_vendistaApiFetch: Failed to parse JSON (Code ${responseCode}). Error: ${e.toString()}. Content: ${contentText}`);
        return { error: `Failed to parse JSON response (Code ${responseCode}): ${e.message}`, raw: contentText, success: false, item: null, items: [] };
      }
    } else { 
      let apiErrorMsg = `API Request Failed: ${responseCode}.`;
      let errorDetailsParsed = null;
      try { errorDetailsParsed = JSON.parse(contentText); } catch (e) { /* не JSON */ }

      if (errorDetailsParsed && errorDetailsParsed.error) {
        apiErrorMsg += ` Message: ${errorDetailsParsed.error}`;
      } else if (errorDetailsParsed && typeof errorDetailsParsed === 'object' && Object.keys(errorDetailsParsed).length > 0) {
        apiErrorMsg += ` Response: ${JSON.stringify(errorDetailsParsed).substring(0,200)}...`;
      } else if (contentText) {
        apiErrorMsg += ` Response: ${contentText.substring(0,200)}...`;
      }
      Logger.log(apiErrorMsg + " Full response details: " + JSON.stringify(errorDetailsParsed || contentText));
      return { error: apiErrorMsg, details: errorDetailsParsed || contentText, responseCode: responseCode, success: false, item: null, items: [] };
    }
  } catch (e) {
    Logger.log(`_vendistaApiFetch Exception: ${e.toString()}`);
    return { error: `Exception during API call: ${e.message}`, success: false, item: null, items: [] };
  }
}

// --- Функции для веб-интерфейса ---
function showMachines(e) {
  return HtmlService.createTemplateFromFile("Index").evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Управление автоматами Vendista");
}

// --- Ingredients API ---
function getIngredients(params) { return _vendistaApiFetch('/ingredients', 'get', params.token, null, params); }
function addIngredient(params) { 
  Logger.log(`addIngredient: Payload: ${JSON.stringify({ name: params.name, measure: params.measure })}`);
  return _vendistaApiFetch('/ingredients', 'post', params.token, { name: params.name, measure: params.measure });
}
function updateIngredient(params) { 
  Logger.log(`updateIngredient: ID: ${params.id}, Payload: ${JSON.stringify({ name: params.name, measure: params.measure })}`);
  return _vendistaApiFetch(`/ingredients/${params.id}`, 'put', params.token, { name: params.name, measure: params.measure });
}
function deleteIngredient(params) {
  Logger.log(`deleteIngredient: ID ${params.id}`);
  return _vendistaApiFetch(`/ingredients/${params.id}`, 'delete', params.token);
}

// --- Recipes API ---
function getRecipes(params) { return _vendistaApiFetch('/recipes', 'get', params.token, null, params); }
function addRecipe(params) {
  Logger.log(`addRecipe: Payload: ${JSON.stringify({ name: params.name, ingredients: params.ingredients })}`);
  return _vendistaApiFetch('/recipes', 'post', params.token, { name: params.name, ingredients: params.ingredients });
}
function updateRecipe(params) {
  Logger.log(`updateRecipe: ID: ${params.id}, Payload: ${JSON.stringify({ name: params.name, ingredients: params.ingredients })}`);
  return _vendistaApiFetch(`/recipes/${params.id}`, 'put', params.token, { name: params.name, ingredients: params.ingredients });
}
function deleteRecipe(params) {
  Logger.log(`deleteRecipe: ID ${params.id}`);
  return _vendistaApiFetch(`/recipes/${params.id}`, 'delete', params.token);
}

// --- Products API ---
function getProducts(params) { return _vendistaApiFetch('/products', 'get', params.token, null, params); }
function getProductById(params) { 
  return _vendistaApiFetch(`/products/${params.id}`, 'get', params.token);
}
function addProduct(params) { 
  const payload = {
    name: params.name,
    recipe_id: params.recipe_id,
    subject_type: params.subject_type 
  };
  for (const key in payload) {
    if (payload[key] === undefined) {
      delete payload[key];
    }
  }
  Logger.log(`addProduct (called by createFullProductStack_NEW): Payload: ${JSON.stringify(payload)}`);
  return _vendistaApiFetch('/products', 'post', params.token, payload);
}
function updateProduct(params) { 
   const payload = {
    name: params.name,
    recipe_id: params.recipe_id, 
    subject_type: params.subject_type, 
    full_name: params.full_name || params.name,
    price: params.price !== undefined ? params.price : 0,
    barcode: params.barcode || null,
    gtin: params.gtin || "",
    is_ingredient: params.is_ingredient, 
    nds_type: params.nds_type || 0,
    cost_type: params.cost_type || 0,
    supplier_id: params.supplier_id || 0
  };
  Logger.log(`updateProduct: ID: ${params.id}, Payload: ${JSON.stringify(payload)}`);
  return _vendistaApiFetch(`/products/${params.id}`, 'put', params.token, payload);
}

function deleteProductCascade(params) {
    const { token, productId } = params;
    if (!token || !productId) return { error: "Token и productId обязательны.", success: false};
    Logger.log(`deleteProductCascade V3.1: Начало удаления товара ID ${productId}`);

    const productInfo = getProductById({token, id: productId});
    let recipeId = null;
    let productName = null;
    let ingredientIdFromRecipe = null;

    if (productInfo.success && productInfo.item) {
        recipeId = productInfo.item.recipe_id;
        productName = productInfo.item.name;
        Logger.log(`deleteProductCascade: Товар найден: name="${productName}", recipe_id=${recipeId}`);
        if (recipeId) {
        const recipeDetails = _vendistaApiFetch(`/recipes/${recipeId}`, 'get', token);
        if (recipeDetails.success && recipeDetails.item && recipeDetails.item.ingredients && recipeDetails.item.ingredients.length > 0) {
            ingredientIdFromRecipe = recipeDetails.item.ingredients[0].ingredient_id;
            Logger.log(`deleteProductCascade: Из рецепта ${recipeId} получен ingredient_id=${ingredientIdFromRecipe}`);
        } else {
            Logger.log(`deleteProductCascade: Рецепт ${recipeId} не найден или не содержит ингредиентов. Ошибка: ${recipeDetails.error || JSON.stringify(recipeDetails)}`);
        }
        }
    } else {
        const err = `Не удалось найти товар ID ${productId} для удаления: ${productInfo.error || 'нет данных item'}`;
        Logger.log(`deleteProductCascade: ${err}`); return { error: err, success: false };
    }

    Logger.log(`deleteProductCascade: Удаление товара ID ${productId}`);
    const deleteProdResult = _vendistaApiFetch(`/products/${productId}`, 'delete', token);
    if (!deleteProdResult.success) Logger.log(`deleteProductCascade: Ошибка при удалении товара ID ${productId}: ${deleteProdResult.error || JSON.stringify(deleteProdResult.details) || 'неизвестно'}`);
    else Logger.log(`deleteProductCascade: Товар ID ${productId} успешно удален.`);

    if (recipeId) {
        Logger.log(`deleteProductCascade: Удаление рецепта ID ${recipeId}`);
        const deleteRecipeResult = deleteRecipe({ token, id: recipeId }); 
        if (!deleteRecipeResult.success) Logger.log(`deleteProductCascade: Ошибка удаления рецепта ID ${recipeId}: ${deleteRecipeResult.error || 'неизвестно'}`);
        else Logger.log(`deleteProductCascade: Рецепт ID ${recipeId} успешно удален.`);
    }

    let finalIngredientIdToDelete = ingredientIdFromRecipe;
    if (!finalIngredientIdToDelete && productName) { 
        Logger.log(`deleteProductCascade: ingredientIdFromRecipe не найден, поиск ингредиента по имени "${productName}".`);
        const ingredSearch = getIngredients({token, FilterText: productName, ItemsOnPage: 5, FilterType: 1 }); 
        if (ingredSearch.success && ingredSearch.items && ingredSearch.items.length > 0) {
        const foundIng = ingredSearch.items.find(i => i.name === productName);
        if (foundIng) {
            finalIngredientIdToDelete = foundIng.id;
            Logger.log(`deleteProductCascade: Ингредиент найден по имени, ID: ${finalIngredientIdToDelete}`);
        } else Logger.log(`deleteProductCascade: Ингредиент с именем "${productName}" не найден при точном поиске.`);
        } else Logger.log(`deleteProductCascade: Поиск ингредиента "${productName}" не дал результатов или ошибка: ${ingredSearch.error}`);
    }

    if (finalIngredientIdToDelete) {
        Logger.log(`deleteProductCascade: Удаление ингредиента ID ${finalIngredientIdToDelete}`);
        const deleteIngResult = deleteIngredient({ token, id: finalIngredientIdToDelete }); 
        if (!deleteIngResult.success) Logger.log(`deleteProductCascade: Ошибка удаления ингредиента ID ${finalIngredientIdToDelete}: ${deleteIngResult.error || 'неизвестно'}`);
        else Logger.log(`deleteProductCascade: Ингредиент ID ${finalIngredientIdToDelete} успешно удален.`);
    } else Logger.log(`deleteProductCascade: Ingredient ID для товара "${productName}" не определен, удаление ингредиента пропущено.`);
    
    return { success: deleteProdResult.success, message: "Каскадное удаление завершено." + (!deleteProdResult.success ? " Были ошибки при удалении основного товара." : "")};
}

// --- ProductMatrix API ---
function getProductMatrices(params) { return _vendistaApiFetch('/productmatrix', 'get', params.token, null, params); }
function getProductMatrixById(params) { return _vendistaApiFetch(`/productmatrix/${params.id}`, 'get', params.token); }
function addProductMatrix(params) { 
  const payload = {
    name: params.name,
    micromarket: params.micromarket !== undefined ? params.micromarket : false, 
    products: params.products || [],
    offset: params.offset !== undefined ? params.offset : 0 
  };
  Logger.log(`addProductMatrix: Payload: ${JSON.stringify(payload)}`);
  return _vendistaApiFetch('/productmatrix', 'post', params.token, payload);
}
function updateProductMatrix(params) { 
  const payload = {
    name: params.name,
    micromarket: params.micromarket !== undefined ? params.micromarket : false,
    products: params.products || [],
    offset: params.offset !== undefined ? params.offset : 0
  };
  Logger.log(`updateProductMatrix: ID: ${params.id}, Payload: ${JSON.stringify(payload)}`);
  return _vendistaApiFetch(`/productmatrix/${params.id}`, 'put', params.token, payload);
}
function deleteProductMatrix(params) { 
  Logger.log(`deleteProductMatrix: Запрос на удаление товарной матрицы ID ${params.id}`);
  return _vendistaApiFetch(`/productmatrix/${params.id}`, 'delete', params.token);
}


// --- MachineModels API ---
function getMachineModels(params) { 
  return _vendistaApiFetch('/machinemodels', 'get', params.token, null, params);
}

// --- Terminals API ---
function getTerminals(params) { return _vendistaApiFetch('/terminals', 'get', params.token, null, params); }
function getTerminalById(params) { 
  if (!params.token || !params.id) return { error: "Token и ID терминала обязательны.", success: false };
  return _vendistaApiFetch(`/terminals/${params.id}`, 'get', params.token);
}
function updateTerminal(params) { 
  if (!params.token || !params.id || !params.payload) return { error: "Token, ID терминала и payload обязательны.", success: false };
  Logger.log(`updateTerminal: ID: ${params.id}, Payload: ${JSON.stringify(params.payload)}`);
  return _vendistaApiFetch(`/terminals/${params.id}`, 'put', params.token, params.payload);
}

// --- TID API ---
function getAvailableTIDs(params) { 
  Logger.log(`getAvailableTIDs: params: ${JSON.stringify(params)}`);
  const queryParams = { 
    ItemsOnPage: params.ItemsOnPage || 1000,
    NotUsedAsPrimary: params.NotUsedAsPrimary === true ? true : undefined, 
    NotUsedAsReserve: params.NotUsedAsReserve === true ? true : undefined 
  };
  for (const key in queryParams) {
    if (queryParams[key] === undefined) {
      delete queryParams[key];
    }
  }
  return _vendistaApiFetch('/tid', 'get', params.token, null, queryParams);
}
function updateTID(params) { 
  if (!params.token || !params.id || !params.payload) return { error: "Token, ID TID и payload обязательны.", success: false };
  
  const allowedTidUpdateFields = ["mcc", "owner_id", "bank_id", "comment", "zip_code", "region", "area", "city", "settlement", "street", "time_zone"];
  const filteredPayload = {};
  for (const key in params.payload) {
    if (allowedTidUpdateFields.includes(key)) {
      filteredPayload[key] = params.payload[key];
    }
  }
  if (Object.keys(filteredPayload).length === 0) {
    Logger.log(`updateTID: Нет разрешенных полей для обновления TID ID ${params.id}. Пропускаем.`);
    return { success: true, message: "Нет полей для обновления TID." }; 
  }
  Logger.log(`updateTID: ID: ${params.id}, Filtered Payload: ${JSON.stringify(filteredPayload)}`);
  return _vendistaApiFetch(`/tid/${params.id}`, 'put', params.token, filteredPayload);
}


// --- Machines API ---
function getMachinesList(token) {
  if (!token) return { error: "Token обязателен (machines).", success: false, items:[] };
  return _vendistaApiFetch('/machines', 'get', token, null, { ItemsOnPage: 1000 });
}
function getMachineById(params){ 
  return _vendistaApiFetch(`/machines/${params.id}`, 'get', params.token);
}
function addMachine(params) {
  const payload = {
    name: params.name, model_id: params.model_id, address: params.address,
    number: params.number || "", terminal_id: params.terminal_id,
    product_matrix_id: params.product_matrix_id,
    micromarket: params.micromarket 
  };
  if (params.latitude !== undefined) payload.latitude = params.latitude;
  if (params.longitude !== undefined) payload.longitude = params.longitude;
  Logger.log(`addMachine: Payload: ${JSON.stringify(payload)}`);
  return _vendistaApiFetch('/machines', 'post', params.token, payload);
 }
function updateMachine(params) {
  const payload = {
    name: params.name, model_id: params.model_id, address: params.address,
    number: params.number || "", terminal_id: params.terminal_id,
    product_matrix_id: params.product_matrix_id,
    micromarket: params.micromarket
  };
  if (params.latitude !== undefined) payload.latitude = params.latitude;
  if (params.longitude !== undefined) payload.longitude = params.longitude;
  Logger.log(`updateMachine: ID: ${params.id}, Payload: ${JSON.stringify(payload)}`);
  return _vendistaApiFetch(`/machines/${params.id}`, 'put', params.token, payload);
}
function deleteMachine(params){ 
  Logger.log(`deleteMachine: Запрос на удаление автомата ID ${params.id}`);
  return _vendistaApiFetch(`/machines/${params.id}`, 'delete', params.token);
}


// --- Сложные операции ---
function createFullProductStack_NEW(params) { 
  const { token, productNameFromUser } = params;
  const logPrefix = "createFullProductStack_NEW V14: "; 

  Logger.log(`${logPrefix}Запуск. productNameFromUser: "${productNameFromUser}", token присутствует: ${!!token}`);

  if (!token || !productNameFromUser || productNameFromUser.trim() === "") {
    Logger.log(`${logPrefix}Ошибка: Токен или имя продукта не предоставлены.`);
    return { success: false, error: "Токен и непустое имя продукта обязательны.", step: "initial_validation", details: null };
  }

  let ingredientId, recipeId, finalProductNameFromIngredientApi;
  let createdEntitiesForRollback = { ingredientId: null, recipeId: null, productId: null };

  Logger.log(`${logPrefix}Шаг 1 - Создание ингредиента "${productNameFromUser}"`);
  const ingredPayload = { name: productNameFromUser, measure: 3 }; 
  const ingredResult = _vendistaApiFetch('/ingredients', 'post', token, ingredPayload);
  Logger.log(`${logPrefix}Шаг 1 - Ответ API (addIngredient): ${JSON.stringify(ingredResult)}`);

  if (!ingredResult || ingredResult.success === false || !ingredResult.item || !ingredResult.item.id) {
    const errMsg = `Ошибка создания ингредиента: ${ingredResult.error || (ingredResult.details ? JSON.stringify(ingredResult.details) : 'Некорректный ответ API - нет item.id')}`;
    Logger.log(`${logPrefix}${errMsg}`);
    return { success: false, error: errMsg, details: ingredResult, step: "add_ingredient" };
  }
  ingredientId = ingredResult.item.id;
  createdEntitiesForRollback.ingredientId = ingredientId;
  finalProductNameFromIngredientApi = ingredResult.item.name; 
  Logger.log(`${logPrefix}Ингредиент создан. ID: ${ingredientId}, Имя от API: "${finalProductNameFromIngredientApi}"`);

  Logger.log(`${logPrefix}Шаг 2 - Создание рецепта для "${finalProductNameFromIngredientApi}" (ингредиент ID: ${ingredientId})`);
  const recipePayload = {
    name: finalProductNameFromIngredientApi,
    ingredients: [{ ingredient_id: ingredientId, count: 1, ingredient_name: finalProductNameFromIngredientApi }]
  };
  const recipeResult = _vendistaApiFetch('/recipes', 'post', token, recipePayload);
  Logger.log(`${logPrefix}Шаг 2 - Ответ API (addRecipe): ${JSON.stringify(recipeResult)}`);

  if (!recipeResult || recipeResult.success === false || !recipeResult.item || !recipeResult.item.id) {
    const errMsg = `Ошибка создания рецепта: ${recipeResult.error || (recipeResult.details ? JSON.stringify(recipeResult.details) : 'Некорректный ответ API - нет item.id')}`;
    Logger.log(`${logPrefix}${errMsg}. Откат ингредиента ID: ${createdEntitiesForRollback.ingredientId}`);
    if (createdEntitiesForRollback.ingredientId) _vendistaApiFetch(`/ingredients/${createdEntitiesForRollback.ingredientId}`, 'delete', token); 
    return { success: false, error: errMsg, details: recipeResult, step: "add_recipe" };
  }
  recipeId = recipeResult.item.id;
  createdEntitiesForRollback.recipeId = recipeId;
  Logger.log(`${logPrefix}Рецепт создан. ID: ${recipeId}`);

  Logger.log(`${logPrefix}Шаг 3 - Создание товара для "${finalProductNameFromIngredientApi}" (рецепт ID: ${recipeId})`);
  const productPayload = {
    name: finalProductNameFromIngredientApi,
    recipe_id: recipeId,        
    subject_type: 1,          
  };
  Logger.log(`${logPrefix}Шаг 3 - Запрос addProduct с: ${JSON.stringify(productPayload)} и токеном.`);
  const productResult = _vendistaApiFetch('/products', 'post', token, productPayload); 
  Logger.log(`${logPrefix}Шаг 3 - Ответ API (addProduct): ${JSON.stringify(productResult)}`);

  if (!productResult || productResult.success === false || !productResult.item || !productResult.item.id) {
    const errMsg = `Ошибка создания товара: ${productResult.error || (productResult.details ? JSON.stringify(productResult.details) : 'Некорректный ответ API - нет item.id')}`;
    Logger.log(`${logPrefix}${errMsg}. Откат рецепта ID: ${createdEntitiesForRollback.recipeId} и ингредиента ID: ${createdEntitiesForRollback.ingredientId}`);
    if (createdEntitiesForRollback.recipeId) _vendistaApiFetch(`/recipes/${createdEntitiesForRollback.recipeId}`, 'delete', token);    
    if (createdEntitiesForRollback.ingredientId) _vendistaApiFetch(`/ingredients/${createdEntitiesForRollback.ingredientId}`, 'delete', token); 
    return { success: false, error: errMsg, details: productResult, step: "add_product" };
  }
  createdEntitiesForRollback.productId = productResult.item.id;
  Logger.log(`${logPrefix}Товар успешно создан. ID: ${productResult.item.id}, Имя от API: "${productResult.item.name}"`);

  return {
    success: true,
    message: "Полный стек товара успешно создан!",
    ingredient: ingredResult.item,
    recipe: recipeResult.item,
    product: productResult.item 
  };
}

function updateFullProductStackName(params) {
  const { token, productId, newProductName } = params;
  const logPrefix = "updateFullProductStackName V2.2: ";
  Logger.log(`${logPrefix}Начало обновления имени для товара ID ${productId} на "${newProductName}"`);

  if (!token || !productId || !newProductName || newProductName.trim() === "") {
    return { success: false, error: "Токен, ID товара и новое непустое имя обязательны." };
  }

  Logger.log(`${logPrefix}Шаг 1 - Получение товара ID ${productId}`);
  const productInfo = getProductById({ token, id: productId });
  if (!productInfo.success || !productInfo.item) {
    return { success: false, error: `Не удалось получить товар ID ${productId}: ${productInfo.error || 'нет item'}`, step: "get_product" };
  }
  const currentProduct = productInfo.item;
  const oldProductName = currentProduct.name;
  const recipeId = currentProduct.recipe_id;
  let ingredientIdToUpdate = null;
  Logger.log(`${logPrefix}Товар получен: старое имя "${oldProductName}", recipe_id ${recipeId}, is_ingredient ${currentProduct.is_ingredient}`);

  if (recipeId) {
    const recipeInfo = _vendistaApiFetch(`/recipes/${recipeId}`, 'get', token);
    if (recipeInfo.success && recipeInfo.item && recipeInfo.item.ingredients && recipeInfo.item.ingredients.length > 0) {
      ingredientIdToUpdate = recipeInfo.item.ingredients[0].ingredient_id;
      Logger.log(`${logPrefix}Найден ingredient_id ${ingredientIdToUpdate} через рецепт.`);
    } else {
      Logger.log(`${logPrefix}Не удалось получить ingredient_id из рецепта ${recipeId}. Поиск по старому имени "${oldProductName}".`);
    }
  }
  if (!ingredientIdToUpdate && oldProductName) { 
    const ingredSearch = getIngredients({ token, FilterText: oldProductName, ItemsOnPage: 5, FilterType: 1 }); 
    if (ingredSearch.success && ingredSearch.items && ingredSearch.items.length > 0) {
      const foundIng = ingredSearch.items.find(i => i.name === oldProductName);
      if (foundIng) {
         ingredientIdToUpdate = foundIng.id;
         Logger.log(`${logPrefix}Найден ingredient_id ${ingredientIdToUpdate} поиском по имени "${oldProductName}".`);
      }
    }
  }
  
  let finalUpdatedNameFromAPI = newProductName; 

  if (ingredientIdToUpdate) {
    Logger.log(`${logPrefix}Шаг 2 - Обновление ингредиента ID ${ingredientIdToUpdate} на имя "${newProductName}" (measure:3)`);
    const updateIngredResult = updateIngredient({ token, id: ingredientIdToUpdate, name: newProductName, measure: 3 }); 
    Logger.log(`${logPrefix}Ответ API (updateIngredient): ${JSON.stringify(updateIngredResult)}`);
    if (!updateIngredResult.success || !updateIngredResult.item) {
      Logger.log(`${logPrefix}Предупреждение: Ошибка обновления ингредиента: ${updateIngredResult.error || 'нет item'}. Продолжаем с запрошенным именем.`);
    } else {
      finalUpdatedNameFromAPI = updateIngredResult.item.name; 
      Logger.log(`${logPrefix}Ингредиент обновлен. Новое имя от API: "${finalUpdatedNameFromAPI}"`);
    }
  } else {
    Logger.log(`${logPrefix}Шаг 2 - Ингредиент для обновления не найден. Пропускаем обновление ингредиента.`);
  }

  if (recipeId) {
    Logger.log(`${logPrefix}Шаг 3 - Обновление рецепта ID ${recipeId} на имя "${finalUpdatedNameFromAPI}"`);
    let ingredientsForRecipeUpdate = [];
    if (ingredientIdToUpdate) {
        ingredientsForRecipeUpdate = [{ ingredient_id: ingredientIdToUpdate, count: 1, ingredient_name: finalUpdatedNameFromAPI }];
    } else if (currentProduct.recipe_ingredients && currentProduct.recipe_ingredients.length > 0) { 
        ingredientsForRecipeUpdate = currentProduct.recipe_ingredients.map(ing => ({...ing, ingredient_name: finalUpdatedNameFromAPI})); 
    } 

    const recipeUpdatePayload = { name: finalUpdatedNameFromAPI, ingredients: ingredientsForRecipeUpdate };
    const updateRecipeResult = updateRecipe({ token, id: recipeId, ...recipeUpdatePayload });
    Logger.log(`${logPrefix}Ответ API (updateRecipe): ${JSON.stringify(updateRecipeResult)}`);
    if (!updateRecipeResult.success || !updateRecipeResult.item) {
      Logger.log(`${logPrefix}Предупреждение: Ошибка обновления рецепта: ${updateRecipeResult.error || 'нет item'}.`);
    } else {
      Logger.log(`${logPrefix}Рецепт обновлен.`);
    }
  } else {
    Logger.log(`${logPrefix}Шаг 3 - Recipe ID (${recipeId}) не найден. Пропускаем обновление рецепта.`);
  }

  Logger.log(`${logPrefix}Шаг 4 - Обновление товара ID ${productId} на имя "${finalUpdatedNameFromAPI}"`);
  const productUpdatePayload = { 
    ...currentProduct, 
    name: finalUpdatedNameFromAPI, 
    full_name: finalUpdatedNameFromAPI 
  };
  delete productUpdatePayload.id; 
  delete productUpdatePayload.token; 
  
  const updateProductResult = updateProduct({ token, id: productId, ...productUpdatePayload });
  Logger.log(`${logPrefix}Ответ API (updateProduct): ${JSON.stringify(updateProductResult)}`);
  if (!updateProductResult.success || !updateProductResult.item) {
    return { success: false, error: `Ошибка обновления товара: ${updateProductResult.error || 'нет item'}`, step: "update_product", details: updateProductResult };
  }
  Logger.log(`${logPrefix}Товар обновлен.`);

  return { success: true, message: `Товар ID ${productId} успешно переименован в "${finalUpdatedNameFromAPI}" (и связанные сущности).`, product: updateProductResult.item };
}


function getOrCreateProductMatrixByName(params) {
  const { token, matrixName } = params;
  if (!token || !matrixName) return { error: "Токен и имя матрицы обязательны.", success: false };
  Logger.log(`getOrCreateProductMatrixByName: Поиск/создание матрицы "${matrixName}"`);
  
  const searchParams = { token, FilterText: matrixName, ItemsOnPage: 5, FilterType: 1 }; 
  const searchResult = getProductMatrices(searchParams);

  if (!searchResult.success && searchResult.error) {
    Logger.log(`getOrCreateProductMatrixByName: Ошибка поиска матрицы API: ${searchResult.error}`);
    return { error: `Ошибка API при поиске матрицы: ${searchResult.error}`, details: searchResult, success: false };
  }
  if (searchResult.success && searchResult.items && searchResult.items.length > 0) {
    const foundMatrix = searchResult.items.find(m => m.name === matrixName);
    if (foundMatrix) {
      Logger.log(`getOrCreateProductMatrixByName: Матрица найдена ID: ${foundMatrix.id}`);
      return { success: true, matrix: foundMatrix, existed: true };
    }
  }
  
  Logger.log(`getOrCreateProductMatrixByName: Матрица "${matrixName}" не найдена, создаем новую.`);
  const addMatrixParams = { token, name: matrixName, products: [], micromarket: false, offset: 0 }; 
  const addResult = addProductMatrix(addMatrixParams);
  if (!addResult.success || !addResult.item || !addResult.item.id) {
    Logger.log(`getOrCreateProductMatrixByName: Ошибка создания матрицы: ${addResult.error || JSON.stringify(addResult.details) || 'нет ID'}`);
    return { error: `Ошибка создания матрицы: ${addResult.error || 'нет ID'}`, details: addResult, success: false };
  }
  Logger.log(`getOrCreateProductMatrixByName: Новая матрица создана ID: ${addResult.item.id}`);
  return { success: true, matrix: addResult.item, existed: false };
}

function getMachineDetailsForView(params) {
  const { token, machineId } = params;
  if (!token || !machineId) return { error: "Токен и ID автомата обязательны.", success: false };

  Logger.log(`getMachineDetailsForView: Запрос деталей для автомата ID ${machineId}`);
  const machineDetails = getMachineById({ token, id: machineId });
  if (!machineDetails.success || !machineDetails.item) {
    Logger.log(`getMachineDetailsForView: Ошибка получения автомата: ${JSON.stringify(machineDetails)}`);
    return { error: `Не удалось загрузить автомат: ${machineDetails.error || 'нет данных'}`, success: false, details: machineDetails };
  }
  
  let terminalData = null;
  let terminalError = null;
  if (machineDetails.item.terminal_id) {
    const terminalResult = getTerminalById({token, id: machineDetails.item.terminal_id});
    if (terminalResult.success && terminalResult.item) {
      terminalData = terminalResult.item;
    } else {
      terminalError = `Терминал ID ${machineDetails.item.terminal_id}: ${terminalResult.error || 'не загружен'}`;
      Logger.log(`getMachineDetailsForView: ${terminalError}`);
    }
  }

  const productMatrixId = machineDetails.item.product_matrix_id;
  let productMatrixData = { id: null, name: "", products: [] }; 
  let matrixError = null;
  if (productMatrixId) {
    const matrixResult = getProductMatrixById({ token, id: productMatrixId });
    if (matrixResult.success && matrixResult.item) {
      productMatrixData = matrixResult.item; 
    } else {
      matrixError = `Матрица ${productMatrixId}: ${matrixResult.error || 'не загружена/нет item'}`;
      Logger.log(`getMachineDetailsForView: ${matrixError}`);
    }
  } else {
     matrixError = "ID товарной матрицы не указан у автомата.";
     Logger.log(`getMachineDetailsForView: У автомата ${machineId} нет product_matrix_id.`);
  }

  const machineIngredientsResult = loadMachineIngredients(machineId, token);
  let ingredientsItems = [];
  let ingredientsError = null;
  if (machineIngredientsResult.success && Array.isArray(machineIngredientsResult.items)) {
    ingredientsItems = machineIngredientsResult.items;
  } else if (!machineIngredientsResult.success) {
    ingredientsError = machineIngredientsResult.error || "Ошибка загрузки ингредиентов автомата";
    Logger.log(`getMachineDetailsForView: Ошибка ингредиентов автомата ID ${machineId}: ${ingredientsError}`);
  }

  return {
    success: true,
    machine: machineDetails.item,
    terminal: terminalData, 
    terminalError: terminalError,
    productMatrix: productMatrixData, 
    matrixError: matrixError, 
    ingredients: ingredientsItems,
    ingredientsError: ingredientsError
  };
}

function loadMachineIngredients(machineId, token) {
  if (!machineId || !token) return { error: "ID машины и токен обязательны.", success: false, items: [] };
  return _vendistaApiFetch(`/machines/${encodeURIComponent(machineId)}/ingredients`, 'get', token);
}

function updateMachineIngredients(params) {
  const { machineId, token, ingredientsPayload } = params;
  if (!machineId || !token || !ingredientsPayload || !Array.isArray(ingredientsPayload.ingredients)) {
    return { success: false, error: "ID автомата, токен и корректный payload (ingredients) обязательны." };
  }
  return _vendistaApiFetch(`/machines/${encodeURIComponent(machineId)}/ingredients`, 'put', token, ingredientsPayload);
}
