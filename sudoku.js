// Constants
const GRID_SIZE_6 = 6;
const GRID_SIZE_4 = 4;
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
const GROUP_BOUNDARIES_6 = [
  { rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 2 }, // Group 1 (top left)
  { rowStart: 0, rowEnd: 1, colStart: 3, colEnd: 5 }, // Group 2 (top right)
  { rowStart: 2, rowEnd: 3, colStart: 0, colEnd: 2 }, // Group 3 (middle left)
  { rowStart: 2, rowEnd: 3, colStart: 3, colEnd: 5 }, // Group 4 (middle right)
  { rowStart: 4, rowEnd: 5, colStart: 0, colEnd: 2 }, // Group 5 (bottom left)
  { rowStart: 4, rowEnd: 5, colStart: 3, colEnd: 5 }  // Group 6 (bottom right)
];

// Group boundaries for 4x4 grid with 2x2 groups
const GROUP_BOUNDARIES_4 = [
  { rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 1 }, // Group 1 (top left)
  { rowStart: 0, rowEnd: 1, colStart: 2, colEnd: 3 }, // Group 2 (top right)
  { rowStart: 2, rowEnd: 3, colStart: 0, colEnd: 1 }, // Group 3 (bottom left)
  { rowStart: 2, rowEnd: 3, colStart: 2, colEnd: 3 }, // Group 4 (bottom right)
];

/**
 * Helper function to handle spreadsheet errors
 * @param {Function} operation - The operation to perform
 * @param {string} errorMessage - The error message prefix
 * @returns {any} The result of the operation
 * @throws {Error} If the operation fails
 */
function handleSpreadsheetError(operation, errorMessage) {
  try {
    return operation();
  } catch (error) {
    throw new Error(`${errorMessage}: ${error.message}`);
  }
}

/**
 * Gets the active spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The active sheet
 */
function getSpreadsheet() {
  return handleSpreadsheetError(
    () => SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet(),
    'Failed to access spreadsheet'
  );
}

/**
 * Gets a sheet by name
 * @param {string} sheetName - The name of the sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
 */
function getSheetByName(sheetName) {
  return handleSpreadsheetError(
    () => SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName),
    `Failed to access sheet "${sheetName}"`
  );
}

/**
 * Gets the grid size based on the number of images in the first row
 * @returns {number} The grid size (4 or 6)
 */
function getGridSize() {
  const sheet = getSpreadsheet();
  let count = 0;
  
  // Count images in the first row, starting from column D (index 4)
  for (let i = 4; i <= 9; i++) {
    const cell = sheet.getRange(1, i);
    const formula = cell.getFormula();
    if (formula.toLowerCase().startsWith('=image(')) {
      count++;
    }
  }
  
  // If we have 4 images in columns D-G, use 4x4 grid
  if (count === 4) {
    return GRID_SIZE_4;
  }
  
  // If we have 6 images in columns D-I, use 6x6 grid
  if (count === 6) {
    return GRID_SIZE_6;
  }
  
  throw new Error(`Invalid number of images found: ${count}. Expected 4 or 6 images.`);
}

/**
 * Gets the group boundaries based on the grid size
 * @param {number} gridSize - The grid size (4 or 6)
 * @returns {Array<Object>} The group boundaries
 */
function getGroupBoundaries(gridSize) {
  return gridSize === GRID_SIZE_4 ? GROUP_BOUNDARIES_4 : GROUP_BOUNDARIES_6;
}

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
 * Inserts an image into a paragraph
 * @param {GoogleAppsScript.Document.Paragraph} paragraph - The paragraph to insert into
 * @param {string} url - The image URL
 */
function insertImage(paragraph, url) {
  const image = getImageBlob(url);
  const insertedImage = paragraph.appendInlineImage(image);
  insertedImage.setWidth(IMAGE_CONFIG.content.width);
  insertedImage.setHeight(IMAGE_CONFIG.content.height);
}

/**
 * Gets the image URL from a spreadsheet cell
 * @param {number} num - The number to map (1-6 or 1-4)
 * @returns {string} The image URL from the corresponding cell
 * @throws {Error} If the number is out of range, cell doesn't contain an image formula, or referenced cell is empty
 */
function getImageFromCell(num) {
  const gridSize = getGridSize();
  if (num < 1 || num > gridSize) {
    throw new Error(`Invalid number: ${num}. Must be between 1 and ${gridSize}`);
  }
  
  const sheet = getSpreadsheet();
  // Add 3 to the column to account for the triple shift (A=shortname, B=longname, C=sheetname)
  const cell = sheet.getRange(1, num + 3);
  const formula = cell.getFormula();
  
  if (!formula.toLowerCase().startsWith('=image(')) {
    throw new Error(`Cell ${String.fromCharCode(65 + num + 2)}1 does not contain an image formula`);
  }
  
  // Extract the content from the formula (could be a URL or cell reference)
  const match = formula.match(/=image\(([^)]+)\)/i);
  if (!match) {
    throw new Error(`Invalid image formula in cell ${String.fromCharCode(65 + num + 2)}1`);
  }
  
  const content = match[1];
  console.log(`Processing image formula for number ${num}: ${formula}`);
  console.log(`Extracted content: ${content}`);
  
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
      console.log(`Referenced cell range: ${referencedCell.getA1Notation()}`);
      
      const value = referencedCell.getValue();
      console.log(`Value in referenced cell: ${value} (type: ${typeof value})`);
      
      if (!value || typeof value !== 'string' || value.trim() === '') {
        throw new Error(`Referenced cell ${content} is empty - please fill in the missing image URL`);
      }
      
      return value;
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
  return handleSpreadsheetError(
    () => {
      const doc = DocumentApp.create(title);
      doc.setMarginTop(36);
      doc.setMarginBottom(18);
      return doc.getBody();
    },
    'Failed to create document'
  );
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
  const gridSize = getGridSize();
  if (!array || !Array.isArray(array) || array.length !== gridSize) {
    throw new Error(`Invalid array dimensions. Expected ${gridSize}x${gridSize}`);
  }
  if (!array.every(row => Array.isArray(row) && row.length === gridSize)) {
    throw new Error(`Invalid row dimensions. Each row must be ${gridSize} elements`);
  }
  if (!array.every(row => row.every(cell => cell === null || (cell >= 1 && cell <= gridSize)))) {
    throw new Error(`Invalid cell values. Each cell must be null or between 1 and ${gridSize}`);
  }
}

/**
 * Gets non-null values from a 2D array slice
 * @param {Array<Array<number|null>>} array - The 2D array
 * @param {Object} boundaries - The boundaries to slice
 * @returns {Array<{url: string, value: number}>} Array of image URLs and their original values
 */
function getNonNullValues(array, boundaries) {
  const values = [];
  for (let i = boundaries.rowStart; i <= boundaries.rowEnd; i++) {
    for (let j = boundaries.colStart; j <= boundaries.colEnd; j++) {
      if (array[i][j] !== null) {
        values.push({
          url: getImageFromCell(array[i][j]),
          value: array[i][j]
        });
      }
    }
  }
  return values;
}

/**
 * Outputs a section with images from values
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {string} title - The section title
 * @param {Array<Array<{url: string, value: number}>>} sections - Array of arrays of image data
 * @param {string} prefix - The prefix for each section (e.g., "ROW", "COLUMN", "GROUP")
 */
function outputSection(body, title, sections, prefix) {
  createSectionHeader(body, title);
  sections.forEach((section, index) => {
    const paragraph = body.appendParagraph(`${prefix} ${index + 1}: `);
    // Sort by the original number value
    section.sort((a, b) => a.value - b.value).forEach(item => insertImage(paragraph, item.url));
    body.appendParagraph(''); // Add spacing between sections
    body.appendHorizontalRule(); // Add horizontal line between sections
  });
  body.appendPageBreak();
}

/**
 * Outputs rows to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function outputRows(sudokuArray, body) {
  validateSudokuArray(sudokuArray);
  const sections = sudokuArray.map(row => 
    row
      .map(value => {
        if (value === null) return null;
        const url = getImageFromCell(value);
        return url ? { url, value } : null;
      })
      .filter(Boolean)
  );
  outputSection(body, 'ROWS must not contain any of these values', sections, 'ROW');
}

/**
 * Outputs columns to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function outputColumns(sudokuArray, body) {
  validateSudokuArray(sudokuArray);
  const gridSize = getGridSize();
  const sections = Array.from({ length: gridSize }, (_, j) =>
    sudokuArray
      .map(row => {
        const value = row[j];
        if (value === null) return null;
        const url = getImageFromCell(value);
        return url ? { url, value } : null;
      })
      .filter(Boolean)
  );
  outputSection(body, 'COLUMNS must not contain any of these values', sections, 'COLUMN');
}

/**
 * Outputs groups to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function outputGroups(sudokuArray, body) {
  validateSudokuArray(sudokuArray);
  const gridSize = getGridSize();
  const groupBoundaries = getGroupBoundaries(gridSize);
  const sections = groupBoundaries.map(boundaries => {
    const values = [];
    for (let i = boundaries.rowStart; i <= boundaries.rowEnd; i++) {
      for (let j = boundaries.colStart; j <= boundaries.colEnd; j++) {
        const value = sudokuArray[i][j];
        if (value !== null) {
          const url = getImageFromCell(value);
          if (url) {
            values.push({ url, value });
          }
        }
      }
    }
    return values;
  });
  outputSection(body, 'GROUPS must not contain any of these values', sections, 'GROUP');
}

/**
 * Creates a reference page with all images
 * @param {GoogleAppsScript.Document.Body} body - The document body
 */
function createReferencePage(body) {
  body.appendPageBreak();
  createSectionHeader(body, 'Reference Images');
  
  const gridSize = getGridSize();
  
  // Create a paragraph for each number (1-4 or 1-6)
  for (let num = 1; num <= gridSize; num++) {
    const paragraph = body.appendParagraph('');
    const url = getImageFromCell(num);
    
    // Insert copies of the same image
    for (let i = 0; i < gridSize; i++) {
      insertImage(paragraph, url);
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
  const sheet = getSpreadsheet();
  const sheetName = sheet.getRange(1, 3).getValue();
  
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('Cell C1 does not contain a valid sheet name');
  }
  
  return sheetName;
}

/**
 * Gets data from a sheet range
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to get data from
 * @returns {Array<Array<any>>} The data from the range
 */
function getSheetData(sheetName, range) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return sheet.getRange(range).getValues();
}

/**
 * Gets the Sudoku puzzle from the answers sheet
 * @returns {Array<Array<number|null>>} The Sudoku puzzle array
 */
function getSudokuPuzzle() {
  const answersSheetName = getAnswersSheetName();
  const gridSize = getGridSize();
  const range = `A1:${String.fromCharCode(64 + gridSize)}${gridSize}`;
  const values = getSheetData(answersSheetName, range);
  const sheet = getSheetByName(answersSheetName);
  
  // Create the Sudoku puzzle array
  const puzzle = [];
  
  // Process each row
  for (let i = 0; i < gridSize; i++) {
    const row = [];
    for (let j = 0; j < gridSize; j++) {
      const value = values[i][j];
      
      // Check if the cell is bold by getting its font weight
      let isBold = false;
      try {
        const cell = sheet.getRange(i + 1, j + 1);
        const fontWeight = cell.getFontWeight();
        isBold = fontWeight === "bold";
      } catch (e) {
        console.log(`Could not get font weight for cell (${i+1}, ${j+1}): ${e.message}`);
      }
      
      // If the cell is bold and contains a number between 1 and gridSize, use it
      // Otherwise, use null
      if (isBold && Number.isInteger(value) && value >= 1 && value <= gridSize) {
        row.push(value);
      } else {
        row.push(null);
      }
    }
    puzzle.push(row);
  }
  
  return puzzle;
}

/**
 * Gets the answers from the answers sheet
 * @returns {Array<Array<number>>} The answers array
 */
function getAnswers() {
  const answersSheetName = getAnswersSheetName();
  const gridSize = getGridSize();
  const range = `A1:${String.fromCharCode(64 + gridSize)}${gridSize}`;
  const values = getSheetData(answersSheetName, range);
  
  // Validate that all values are numbers between 1 and gridSize
  if (!values.every(row => row.every(cell => Number.isInteger(cell) && cell >= 1 && cell <= gridSize))) {
    throw new Error(`Invalid values in "${answersSheetName}" sheet. All values must be integers between 1 and ${gridSize}`);
  }
  
  return values;
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
  answers.forEach(row => {
    const paragraph = body.appendParagraph('');
    row.forEach(value => {
      const url = getImageFromCell(value);
      insertImage(paragraph, url);
    });
    body.appendParagraph(''); // Add spacing between rows
  });
}

/**
 * Creates a SudokuGrid sheet with X's for bold numbers
 */
function createSudokuGrid() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const gridSize = getGridSize();
    const templateName = gridSize === GRID_SIZE_4 ? 'Template4' : 'Template6';
    const templateSheet = spreadsheet.getSheetByName(templateName);
    if (!templateSheet) {
      throw new Error(`${templateName} sheet not found`);
    }

    // Get the shortname from column A and longname from column B
    const shortname = getSpreadsheet().getRange(1, 1).getValue();
    if (!shortname || typeof shortname !== 'string') {
      throw new Error('Cell A1 does not contain a valid shortname');
    }

    const longname = getSpreadsheet().getRange(1, 2).getValue();
    if (!longname || typeof longname !== 'string') {
      throw new Error('Cell B1 does not contain a valid longname');
    }

    // Delete the sheet if it already exists
    const existingSheet = spreadsheet.getSheetByName(shortname);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }

    // Create a copy of the template and name it using the shortname
    const sudokuGrid = templateSheet.copyTo(spreadsheet);
    sudokuGrid.setName(shortname);

    // Get the named range from the spreadsheet to find the correct cell
    const namedRange = gridSize === GRID_SIZE_4 ? 'Template4Name' : 'Template6Name';
    const templateRange = spreadsheet.getRangeByName(namedRange);
    if (!templateRange) {
      throw new Error(`Named range "${namedRange}" not found`);
    }

    // Get the cell coordinates from the template range
    const row = templateRange.getRow();
    const col = templateRange.getColumn();

    // Set the longname in the corresponding cell of the new sheet
    sudokuGrid.getRange(row, col).setValue(longname);

    // Get the answers sheet name and data
    const answersSheetName = getAnswersSheetName();
    const answersSheet = spreadsheet.getSheetByName(answersSheetName);
    if (!answersSheet) {
      throw new Error(`Sheet "${answersSheetName}" not found`);
    }

    const answersRange = `A1:${String.fromCharCode(64 + gridSize)}${gridSize}`;
    const values = answersSheet.getRange(answersRange).getValues();

    // Process each cell in the answers sheet
    for (let i = 0; i < gridSize; i++) {
      for (let j = 0; j < gridSize; j++) {
        const value = values[i][j];
        
        // Check if the cell is bold
        let isBold = false;
        try {
          const cell = answersSheet.getRange(i + 1, j + 1);
          const fontWeight = cell.getFontWeight();
          isBold = fontWeight === "bold";
        } catch (e) {
          console.log(`Could not get font weight for cell (${i+1}, ${j+1}): ${e.message}`);
        }
        
        // If the cell is bold and contains a number, set the corresponding cell in SudokuGrid to 'X'
        if (isBold && Number.isInteger(value) && value >= 1 && value <= gridSize) {
          // Add 1 to row and column to account for the offset in SudokuGrid
          const sudokuCell = sudokuGrid.getRange(i + 2, j + 2);
          sudokuCell.setValue('X');
        }
      }
    }
  } catch (error) {
    console.error('Error creating SudokuGrid:', error.message);
    throw error;
  }
}

/**
 * Main function to run all outputs
 */
function main() {
  try {
    // Get the shortname from column A
    const shortname = getSpreadsheet().getRange(1, 1).getValue();
    if (!shortname || typeof shortname !== 'string') {
      throw new Error('Cell A1 does not contain a valid shortname');
    }

    const body = createDocument(`Mint Hulzo Coin - ${shortname}`);
    const puzzle = getSudokuPuzzle();
    
    outputRows(puzzle, body);
    outputColumns(puzzle, body);
    outputGroups(puzzle, body);
    createReferencePage(body);
    createAnswersSheet(body);
    createSudokuGrid();
  } catch (error) {
    console.error('Error processing sudoku:', error.message);
    throw error; // Re-throw to show in Apps Script execution log
  }
}
