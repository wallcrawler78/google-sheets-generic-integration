/**
 * Category Favorites Management
 * Extension to CategoryManager for managing favorite categories
 */

/**
 * Gets favorite category GUIDs
 * @return {Array<string>} Array of favorite category GUIDs
 */
function getFavoriteCategories() {
  var json = PropertiesService.getScriptProperties().getProperty(PROP_FAVORITE_CATEGORIES);
  if (!json) {
    return [];
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    return [];
  }
}

/**
 * Saves favorite category GUIDs
 * @param {Array<string>} favorites - Array of category GUIDs
 */
function saveFavoriteCategories(favorites) {
  PropertiesService.getScriptProperties().setProperty(
    PROP_FAVORITE_CATEGORIES,
    JSON.stringify(favorites)
  );
}

/**
 * Adds a category to favorites
 * @param {string} categoryGuid - Category GUID to add
 */
function addFavoriteCategory(categoryGuid) {
  var favorites = getFavoriteCategories();

  // Check if already in favorites
  if (favorites.indexOf(categoryGuid) === -1) {
    favorites.push(categoryGuid);
    saveFavoriteCategories(favorites);
  }
}

/**
 * Removes a category from favorites
 * @param {string} categoryGuid - Category GUID to remove
 */
function removeFavoriteCategory(categoryGuid) {
  var favorites = getFavoriteCategories();
  var index = favorites.indexOf(categoryGuid);

  if (index !== -1) {
    favorites.splice(index, 1);
    saveFavoriteCategories(favorites);
  }
}

/**
 * Toggles a category's favorite status
 * @param {string} categoryGuid - Category GUID to toggle
 * @return {boolean} New favorite status
 */
function toggleFavoriteCategory(categoryGuid) {
  var favorites = getFavoriteCategories();
  var index = favorites.indexOf(categoryGuid);

  if (index !== -1) {
    favorites.splice(index, 1);
    saveFavoriteCategories(favorites);
    return false;
  } else {
    favorites.push(categoryGuid);
    saveFavoriteCategories(favorites);
    return true;
  }
}

/**
 * Checks if a category is a favorite
 * @param {string} categoryGuid - Category GUID to check
 * @return {boolean} True if favorite
 */
function isFavoriteCategory(categoryGuid) {
  var favorites = getFavoriteCategories();
  return favorites.indexOf(categoryGuid) !== -1;
}

/**
 * Gets all categories with favorite status
 * @return {Array<Object>} Array of category objects with isFavorite property
 */
function getCategoriesWithFavorites() {
  var categories = getArenaCategories();
  var favorites = getFavoriteCategories();

  return categories.map(function(cat) {
    return {
      guid: cat.guid,
      name: cat.name,
      path: cat.path,
      fullPath: cat.fullPath,
      isFavorite: favorites.indexOf(cat.guid) !== -1
    };
  });
}

/**
 * Gets only favorite categories
 * @return {Array<Object>} Array of favorite category objects
 */
function getOnlyFavoriteCategories() {
  var categories = getCategoriesWithFavorites();
  return categories.filter(function(cat) {
    return cat.isFavorite;
  });
}
