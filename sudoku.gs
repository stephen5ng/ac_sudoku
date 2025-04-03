// Sudoku puzzle data
const sudoku = [
  [null, 6, null, 5, null, null],
  [3, null, null, null, null, null],
  [2, 3, null, 6, 5, null],
  [null, null, 1, 4, null, null],
  [null, null, null, 1, null, 4],
  [1, null, null, null, null, null],
];

const answers = [
  [4, 6, 2, 5, 1, 3],
  [3, 1, 5, 2, 4, 6],
  [2, 3, 4, 6, 5, 1],
  [6, 5, 1, 4, 3, 2],
  [5, 2, 3, 1, 6, 4],
  [1, 4, 6, 3, 2, 5],
];

// Constants
const GRID_SIZE = 6;
const GROUP_SIZE = 2;
const GROUP_COUNT = 3;
const SPREADSHEET_ID = '1t9mwKfa_aPzJwx6qUOgO54N9-1XBQJGCPKY3PpwF-BE';

// Image configuration
const IMAGE_CONFIG = {
  header: {
    width: 256,
    height: 256
  },
  content: {
    width: 100,
    height: 100
  }
};

// Group boundaries for 6x6 grid with 2x3 groups
const GROUP_BOUNDARIES = [
  { rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 2 }, // Group 1 (top left)
  { rowStart: 0, rowEnd: 1, colStart: 3, colEnd: 5 }, // Group 2 (top right)
  { rowStart: 2, rowEnd: 3, colStart: 0, colEnd: 2 }, // Group 3 (middle left)
  { rowStart: 2, rowEnd: 3, colStart: 3, colEnd: 5 }, // Group 4 (middle right)
  { rowStart: 4, rowEnd: 5, colStart: 0, colEnd: 2 }, // Group 5 (bottom left)
  { rowStart: 4, rowEnd: 5, colStart: 3, colEnd: 5 }  // Group 6 (bottom right)
];


/**
 * Gets the active spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The active sheet
 * @throws {Error} If spreadsheet cannot be accessed
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  } catch (error) {
    throw new Error(`Failed to access spreadsheet: ${error.message}`);
  }
}

/**
 * Gets the image URL from a spreadsheet cell
 * @param {number} num - The number to map (1-6)
 * @returns {string} The image URL from the corresponding cell
 * @throws {Error} If the number is out of range or cell doesn't contain an image
 */
function getImageFromCell(num) {
  if (num < 1 || num > GRID_SIZE) {
    throw new Error(`Invalid number: ${num}. Must be between 1 and ${GRID_SIZE}`);
  }
  
  const sheet = getSpreadsheet();
  const cell = sheet.getRange(1, num);
  const formula = cell.getFormula();
  
  if (!formula.toLowerCase().startsWith('=image(')) {
    throw new Error(`Cell ${String.fromCharCode(65 + num)}1 does not contain an image formula`);
  }
  
  // Extract the content from the formula (could be a URL or cell reference)
  const match = formula.match(/=image\(([^)]+)\)/i);
  if (!match) {
    throw new Error(`Invalid image formula in cell ${String.fromCharCode(65 + num)}1`);
  }
  
  const content = match[1];
  
  // If it's a quoted string, it's a direct URL
  if (content.startsWith('"') && content.endsWith('"')) {
    return content.slice(1, -1);
  }
  
  // Otherwise, it's a cell reference
  try {
    const referencedCell = sheet.getParent().getRange(content);
    const url = referencedCell.getValue();
    
    if (!url || typeof url !== 'string') {
      throw new Error(`Referenced cell ${content} does not contain a valid URL`);
    }
    
    return url;
  } catch (error) {
    throw new Error(`Failed to get URL from referenced cell ${content}: ${error.message}`);
  }
}

/**
 * Creates a new Google Doc and returns its body
 * @param {string} title - The title of the document
 * @returns {GoogleAppsScript.Document.Body} The document body
 */
function createDocument(title) {
  try {
    const doc = DocumentApp.create(title);
    doc.setMarginTop(36);
    doc.setMarginBottom(18);
    return doc.getBody();
  } catch (error) {
    throw new Error(`Failed to create document: ${error.message}`);
  }
}

/**
 * Inserts an image from a URL into the document
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {string} url - The URL of the image
 * @param {Object} config - Image configuration (width and height)
 * @returns {GoogleAppsScript.Document.InlineImage} The inserted image
 * @throws {Error} If image insertion fails
 */
function insertImageFromUrl(body, url, config) {
  try {
    const image = UrlFetchApp.fetch(url).getBlob();
    const resizedImage = image.setContentType('image/png');
    const paragraph = body.appendParagraph('');
    const insertedImage = paragraph.appendInlineImage(resizedImage);
    insertedImage.setWidth(config.width);
    insertedImage.setHeight(config.height);
    return insertedImage;
  } catch (error) {
    throw new Error(`Failed to insert image from URL: ${error.message}`);
  }
}

/**
 * Creates a section header in the document
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {string} text - The header text
 */
function createSectionHeader(body, text) {
  const header = body.appendParagraph(text);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

/**
 * Validates the sudoku array
 * @param {Array<Array<number|null>>} array - The sudoku array to validate
 * @throws {Error} If the array is invalid
 */
function validateSudokuArray(array) {
  if (!array || !Array.isArray(array) || array.length !== GRID_SIZE) {
    throw new Error(`Invalid array dimensions. Expected ${GRID_SIZE}x${GRID_SIZE}`);
  }
  if (!array.every(row => Array.isArray(row) && row.length === GRID_SIZE)) {
    throw new Error(`Invalid row dimensions. Each row must be ${GRID_SIZE} elements`);
  }
  if (!array.every(row => row.every(cell => cell === null || (cell >= 1 && cell <= GRID_SIZE)))) {
    throw new Error(`Invalid cell values. Each cell must be null or between 1 and ${GRID_SIZE}`);
  }
}

/**
 * Gets non-null values from a 2D array slice
 * @param {Array<Array<number|null>>} array - The 2D array
 * @param {Object} boundaries - The boundaries to slice
 * @returns {Array<string>} Array of image URLs from corresponding cells
 */
function getNonNullValues(array, boundaries) {
  const values = [];
  for (let i = boundaries.rowStart; i <= boundaries.rowEnd; i++) {
    for (let j = boundaries.colStart; j <= boundaries.colEnd; j++) {
      if (array[i][j] !== null) {
        values.push(getImageFromCell(array[i][j]));
      }
    }
  }
  return values;
}

/**
 * Outputs rows to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function outputRows(sudokuArray, body) {
  validateSudokuArray(sudokuArray);
  
  createSectionHeader(body, 'ROWS must not contain any of these values');
  sudokuArray.forEach((row, index) => {
    const paragraph = body.appendParagraph(`ROW ${index + 1}: `);
    row
      .filter(value => value !== null)
      .forEach(value => {
        const url = getImageFromCell(value);
        const image = UrlFetchApp.fetch(url).getBlob();
        const resizedImage = image.setContentType('image/png');
        const insertedImage = paragraph.appendInlineImage(resizedImage);
        insertedImage.setWidth(IMAGE_CONFIG.content.width);
        insertedImage.setHeight(IMAGE_CONFIG.content.height);
      });
    body.appendParagraph(''); // Add spacing between rows
  });
  body.appendPageBreak();
}

/**
 * Outputs columns to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function outputColumns(sudokuArray, body) {
  validateSudokuArray(sudokuArray);
  
  createSectionHeader(body, 'COLUMNS must not contain any of these values');
  for (let j = 0; j < GRID_SIZE; j++) {
    const paragraph = body.appendParagraph(`COLUMN ${j + 1}: `);
    sudokuArray
      .map(row => row[j])
      .filter(value => value !== null)
      .forEach(value => {
        const url = getImageFromCell(value);
        const image = UrlFetchApp.fetch(url).getBlob();
        const resizedImage = image.setContentType('image/png');
        const insertedImage = paragraph.appendInlineImage(resizedImage);
        insertedImage.setWidth(IMAGE_CONFIG.content.width);
        insertedImage.setHeight(IMAGE_CONFIG.content.height);
      });
    body.appendParagraph(''); // Add spacing between columns
  }
  body.appendPageBreak();
}

/**
 * Outputs groups to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function outputGroups(sudokuArray, body) {
  validateSudokuArray(sudokuArray);
  
  createSectionHeader(body, 'GROUPS must not contain any of these values');
  GROUP_BOUNDARIES.forEach((boundaries, index) => {
    const paragraph = body.appendParagraph(`GROUP ${index + 1}: `);
    const groupValues = getNonNullValues(sudokuArray, boundaries);
    groupValues.forEach(url => {
      const image = UrlFetchApp.fetch(url).getBlob();
      const resizedImage = image.setContentType('image/png');
      const insertedImage = paragraph.appendInlineImage(resizedImage);
      insertedImage.setWidth(IMAGE_CONFIG.content.width);
      insertedImage.setHeight(IMAGE_CONFIG.content.height);
    });
    body.appendParagraph(''); // Add spacing between groups
  });
}

/**
 * Creates a reference page with all images
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function createReferencePage(body) {
  body.appendPageBreak();
  createSectionHeader(body, 'Reference Images');
  
  // Create a paragraph for each number (1-6)
  for (let num = 1; num <= GRID_SIZE; num++) {
    const paragraph = body.appendParagraph('');
    const url = getImageFromCell(num);
    const image = UrlFetchApp.fetch(url).getBlob();
    const resizedImage = image.setContentType('image/png');
    
    // Insert 6 copies of the same image
    for (let i = 0; i < 6; i++) {
      const insertedImage = paragraph.appendInlineImage(resizedImage);
      insertedImage.setWidth(IMAGE_CONFIG.content.width);
      insertedImage.setHeight(IMAGE_CONFIG.content.height);
    }
    
    body.appendParagraph(''); // Add spacing between rows
  }
}

/**
 * Creates the answers sheet with the complete solution
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function createAnswersSheet(body) {
  body.appendPageBreak();
  createSectionHeader(body, 'Solution');
  
  // Create a row for each answer array
  answers.forEach((row, index) => {
    const paragraph = body.appendParagraph(``);
    row.forEach(value => {
      const url = getImageFromCell(value);
      const image = UrlFetchApp.fetch(url).getBlob();
      const resizedImage = image.setContentType('image/png');
      const insertedImage = paragraph.appendInlineImage(resizedImage);
      insertedImage.setWidth(IMAGE_CONFIG.content.width);
      insertedImage.setHeight(IMAGE_CONFIG.content.height);
    });
    body.appendParagraph(''); // Add spacing between rows
  });
}

/**
 * Main function to run all outputs
 */
function main() {
  try {
    const body = createDocument('Mint Hulzo Coin');
    
    outputRows(sudoku, body);
    outputColumns(sudoku, body);
    outputGroups(sudoku, body);
    createReferencePage(body);
    createAnswersSheet(body);
  } catch (error) {
    console.error('Error processing sudoku:', error.message);
    throw error; // Re-throw to show in Apps Script execution log
  }
}
