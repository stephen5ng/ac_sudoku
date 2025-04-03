// Constants
const GRID_SIZE = 6;
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

// Cache for image blobs
const imageCache = new Map();

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
 * Gets an image blob from cache or fetches it
 * @param {string} url - The image URL
 * @returns {GoogleAppsScript.Base.Blob} The image blob
 */
function getImageBlob(url) {
  if (imageCache.has(url)) {
    return imageCache.get(url);
  }
  
  const image = UrlFetchApp.fetch(url).getBlob();
  const resizedImage = image.setContentType('image/png');
  imageCache.set(url, resizedImage);
  return resizedImage;
}

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
  console.log(`Processing image formula for number ${num}: ${formula}`);
  
  // If it's a quoted string, it's a direct URL
  if (content.startsWith('"') && content.endsWith('"')) {
    return content.slice(1, -1);
  }
  
  // If it contains a sheet reference (e.g., "Sheet2!"), it's a cell reference
  if (content.includes('!')) {
    try {
      console.log(`Looking up cell reference: ${content}`);
      const spreadsheet = sheet.getParent();
      const referencedCell = spreadsheet.getRange(content);
      const url = referencedCell.getValue();
      
      console.log(`Found URL in referenced cell: ${url}`);
      
      if (!url || typeof url !== 'string') {
        throw new Error(`Referenced cell ${content} does not contain a valid URL`);
      }
      
      return url;
    } catch (error) {
      console.error(`Error processing cell reference ${content}:`, error);
      throw new Error(`Failed to get URL from referenced cell ${content}: ${error.message}`);
    }
  }
  
  // If we get here, assume it's a direct URL
  return content;
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
        const image = getImageBlob(url);
        const insertedImage = paragraph.appendInlineImage(image);
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
        const image = getImageBlob(url);
        const insertedImage = paragraph.appendInlineImage(image);
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
      const image = getImageBlob(url);
      const insertedImage = paragraph.appendInlineImage(image);
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
    const image = getImageBlob(url);
    
    // Insert 6 copies of the same image
    for (let i = 0; i < 6; i++) {
      const insertedImage = paragraph.appendInlineImage(image);
      insertedImage.setWidth(IMAGE_CONFIG.content.width);
      insertedImage.setHeight(IMAGE_CONFIG.content.height);
    }
    
    body.appendParagraph(''); // Add spacing between rows
  }
}

/**
 * Gets the name of the answers sheet from the images sheet
 * @returns {string} The name of the answers sheet
 * @throws {Error} If the sheet name cannot be read
 */
function getAnswersSheetName() {
  try {
    const sheet = getSpreadsheet();
    const sheetName = sheet.getRange('G1').getValue();
    
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('Cell G1 does not contain a valid sheet name');
    }
    
    return sheetName;
  } catch (error) {
    throw new Error(`Failed to get answers sheet name: ${error.message}`);
  }
}

/**
 * Gets the Sudoku puzzle from the answers sheet
 * @returns {Array<Array<number|null>>} The Sudoku puzzle array
 * @throws {Error} If puzzle cannot be read from the sheet
 */
function getSudokuPuzzle() {
  try {
    const answersSheetName = getAnswersSheetName();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(answersSheetName);
    if (!sheet) {
      throw new Error(`Sheet "${answersSheetName}" not found`);
    }
    
    const range = sheet.getRange('A1:F6');
    const values = range.getValues();
    
    // Create the Sudoku puzzle array
    const puzzle = [];
    
    // Process each row
    for (let i = 0; i < GRID_SIZE; i++) {
      const row = [];
      for (let j = 0; j < GRID_SIZE; j++) {
        const value = values[i][j];
        
        // Check if the cell is bold by getting its font weight
        let isBold = false;
        try {
          const cell = sheet.getRange(i + 1, j + 1);
          const fontWeight = cell.getFontWeight();
          // Check if the font weight is "bold" (string comparison)
          isBold = fontWeight === "bold";
        } catch (e) {
          console.log(`Could not get font weight for cell (${i+1}, ${j+1}): ${e.message}`);
        }
        
        // If the cell is bold and contains a number between 1 and 6, use it
        // Otherwise, use null
        if (isBold && Number.isInteger(value) && value >= 1 && value <= GRID_SIZE) {
          row.push(value);
        } else {
          row.push(null);
        }
      }
      puzzle.push(row);
    }
    
    return puzzle;
  } catch (error) {
    throw new Error(`Failed to get Sudoku puzzle: ${error.message}`);
  }
}

/**
 * Gets the answers from the answers sheet
 * @returns {Array<Array<number>>} The answers array
 * @throws {Error} If answers cannot be read from the sheet
 */
function getAnswers() {
  try {
    const answersSheetName = getAnswersSheetName();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(answersSheetName);
    if (!sheet) {
      throw new Error(`Sheet "${answersSheetName}" not found`);
    }
    
    const range = sheet.getRange('A1:F6');
    const values = range.getValues();
    
    // Validate that all values are numbers between 1 and 6
    if (!values.every(row => row.every(cell => Number.isInteger(cell) && cell >= 1 && cell <= GRID_SIZE))) {
      throw new Error(`Invalid values in "${answersSheetName}" sheet. All values must be integers between 1 and 6`);
    }
    
    return values;
  } catch (error) {
    throw new Error(`Failed to get answers: ${error.message}`);
  }
}

/**
 * Creates the answers sheet with the complete solution
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function createAnswersSheet(body) {
  body.appendPageBreak();
  createSectionHeader(body, 'Solution');
  
  // Get answers from the sheet
  const answers = getAnswers();
  
  // Create a row for each answer array
  answers.forEach((row, index) => {
    const paragraph = body.appendParagraph(``);
    row.forEach(value => {
      const url = getImageFromCell(value);
      const image = getImageBlob(url);
      const insertedImage = paragraph.appendInlineImage(image);
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
    
    // Get the Sudoku puzzle from the sheet
    const puzzle = getSudokuPuzzle();
    
    outputRows(puzzle, body);
    outputColumns(puzzle, body);
    outputGroups(puzzle, body);
    createReferencePage(body);
    createAnswersSheet(body);
  } catch (error) {
    console.error('Error processing sudoku:', error.message);
    throw error; // Re-throw to show in Apps Script execution log
  }
}
