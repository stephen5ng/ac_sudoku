// Constants
const GRID_SIZE_6 = 6;
const GRID_SIZE_4 = 4;
const SPREADSHEET_ID = '1t9mwKfa_aPzJwx6qUOgO54N9-1XBQJGCPKY3PpwF-BE';
const NEGATE_DECLARATIONS = false;

// Section title text
const MUST_NOT_CONTAIN = 'must not contain any of these values';
const MAY_ONLY_CONTAIN = 'may only contain one of these values';

// Section types
const SECTION_TYPES = {
  ROWS: 'ROWS',
  COLUMNS: 'COLUMNS',
  GROUPS: 'GROUPS'
};

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
 * Gets the Sudokus sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Sudokus sheet
 */
function getSpreadsheet() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Sudokus');
    if (!sheet) {
      throw new Error('Sudokus sheet not found');
    }
    return sheet;
  } catch (error) {
    throw new Error(`Failed to access Sudokus sheet: ${error.message}`);
  }
}

/**
 * Gets a sheet by name
 * @param {string} sheetName - The name of the sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
 */
function getSheetByName(sheetName) {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  } catch (error) {
    throw new Error(`Failed to access sheet "${sheetName}": ${error.message}`);
  }
}

/**
 * Gets the grid size based on the number of images in the specified row
 * @param {number} row - The row number to check for images
 * @returns {number} The grid size (4 or 6)
 */
function getGridSize(row) {
  const sheet = getSpreadsheet();
  let count = 0;
  
  // Count images in the specified row, starting from column D (index 4)
  for (let i = 4; i <= 9; i++) {
    const formula = sheet.getRange(row, i).getFormula();
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
  
  throw new Error(`Invalid number of images found: ${count}. Expected 4 or 6 images in row ${row}.`);
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
 * @param {string} answersSheetName - The name of the answers sheet to determine grid size
 * @param {number} row - The row number to get the image from
 * @returns {string} The image URL from the corresponding cell
 * @throws {Error} If the number is out of range, cell doesn't contain an image formula, or referenced cell is empty
 */
function getImageFromCell(num, answersSheetName, row) {
  console.log(`getImageFromCell called with num=${num}, answersSheetName=${answersSheetName}, row=${row}`);
  const gridSize = answersSheetName.includes('6') ? GRID_SIZE_6 : GRID_SIZE_4;
  console.log(`Determined gridSize=${gridSize} from answersSheetName`);
  if (num < 1 || num > gridSize) {
    throw new Error(`Invalid number: ${num}. Must be between 1 and ${gridSize} (answersSheetName=${answersSheetName})`);
  }
  
  const sheet = getSpreadsheet();
  // Add 3 to the column to account for the triple shift (A=shortname, B=longname, C=sheetname)
  const cell = sheet.getRange(row, num + 3);
  const formula = cell.getFormula();
  
  if (!formula.toLowerCase().startsWith('=image(')) {
    throw new Error(`Cell ${String.fromCharCode(65 + num + 2)}${row} does not contain an image formula`);
  }
  
  // Extract the content from the formula (could be a URL or cell reference)
  const match = formula.match(/=image\(([^)]+)\)/i);
  if (!match) {
    throw new Error(`Invalid image formula in cell ${String.fromCharCode(65 + num + 2)}${row}`);
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
 * @param {number} row - The row number to determine grid size
 * @throws {Error} If the array is invalid
 */
function validateSudokuArray(array, row) {
  const gridSize = getGridSize(row);
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
          url: getImageFromCell(array[i][j], getAnswersSheetName(i + 1), i + 1),
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
 * Gets the section title based on type and declaration mode
 * @param {string} sectionType - The type of section (ROWS, COLUMNS, GROUPS)
 * @returns {string} The formatted section title
 */
function getSectionTitle(sectionType) {
  const declaration = NEGATE_DECLARATIONS ? MUST_NOT_CONTAIN : MAY_ONLY_CONTAIN;
  return `${sectionType} ${declaration}`;
}

/**
 * Outputs rows to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {number} currentRow - The row number to get the answers from
 */
function outputRows(sudokuArray, body, currentRow) {
  console.log(`outputRows called with currentRow=${currentRow}, type=${typeof currentRow}`);
  validateSudokuArray(sudokuArray, currentRow);
  const sections = sudokuArray.map(sudokuRow => 
    sudokuRow
      .map(value => {
        if (value === null) return null;
        const url = getImageFromCell(value, getAnswersSheetName(currentRow), currentRow);
        return url ? { url, value } : null;
      })
      .filter(Boolean)
  );
  outputSection(body, getSectionTitle(SECTION_TYPES.ROWS), sections, 'ROW');
}

/**
 * Outputs columns to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {number} currentRow - The row number to get the answers from
 */
function outputColumns(sudokuArray, body, currentRow) {
  validateSudokuArray(sudokuArray, currentRow);
  const gridSize = getGridSize(currentRow);
  const sections = Array.from({ length: gridSize }, (_, j) =>
    sudokuArray
      .map(sudokuRow => {
        const value = sudokuRow[j];
        if (value === null) return null;
        const url = getImageFromCell(value, getAnswersSheetName(currentRow), currentRow);
        return url ? { url, value } : null;
      })
      .filter(Boolean)
  );
  outputSection(body, getSectionTitle(SECTION_TYPES.COLUMNS), sections, 'COLUMN');
}

/**
 * Outputs groups to the document
 * @param {Array<Array<number|null>>} sudokuArray - The sudoku array
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {number} currentRow - The row number to get the answers from
 */
function outputGroups(sudokuArray, body, currentRow) {
  validateSudokuArray(sudokuArray, currentRow);
  const gridSize = getGridSize(currentRow);
  const groupBoundaries = getGroupBoundaries(gridSize);
  const sections = groupBoundaries.map(boundaries => {
    const values = [];
    for (let i = boundaries.rowStart; i <= boundaries.rowEnd; i++) {
      for (let j = boundaries.colStart; j <= boundaries.colEnd; j++) {
        const value = sudokuArray[i][j];
        if (value !== null) {
          const url = getImageFromCell(value, getAnswersSheetName(currentRow), currentRow);
          if (url) {
            values.push({ url, value });
          }
        }
      }
    }
    return values;
  });
  outputSection(body, getSectionTitle(SECTION_TYPES.GROUPS), sections, 'GROUP');
}

/**
 * Creates a reference page with all images
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {number} row - The row number to get the images from
 */
function createReferencePage(body, row) {
  body.appendPageBreak();
  createSectionHeader(body, 'Reference Images');
  
  const gridSize = getGridSize(row);
  
  // Create a paragraph for each number (1-4 or 1-6)
  for (let num = 1; num <= gridSize; num++) {
    const paragraph = body.appendParagraph('');
    const url = getImageFromCell(num, getAnswersSheetName(row), row);
    
    // Insert copies of the same image
    for (let i = 0; i < gridSize; i++) {
      insertImage(paragraph, url);
    }
    
    body.appendParagraph(''); // Add spacing between rows
  }
}

/**
 * Gets the name of the answers sheet from the Sudokus sheet
 * @param {number} row - The row number to get the sheet name from
 * @returns {string} The name of the answers sheet
 * @throws {Error} If the sheet name cannot be read
 */
function getAnswersSheetName(row) {
  console.log(`getAnswersSheetName called with row=${row}, type=${typeof row}`);
  if (!row || typeof row !== 'number' || row < 1) {
    throw new Error(`Invalid row number: ${row} (type: ${typeof row})`);
  }
  
  const sheet = getSpreadsheet();
  const sheetName = sheet.getRange(row, 3).getValue();
  
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error(`Cell C${row} does not contain a valid sheet name`);
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
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    return sheet.getRange(range).getValues();
  } catch (error) {
    throw new Error(`Failed to get data from range "${range}" in sheet "${sheetName}": ${error.message}`);
  }
}

/**
 * Gets the Sudoku puzzle from the answers sheet, returning either bold or non-bold numbers based on NEGATE_DECLARATIONS
 * @param {number} row - The row number to get the puzzle from
 * @returns {Array<Array<number|null>>} The Sudoku puzzle array
 */
function getSudokuPuzzle(row) {
  const answersSheetName = getAnswersSheetName(row);
  const gridSize = getGridSize(row);
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
      
      // If NEGATE_DECLARATIONS is false, we want non-bold numbers
      // If NEGATE_DECLARATIONS is true, we want bold numbers (original behavior)
      const shouldInclude = NEGATE_DECLARATIONS ? isBold : !isBold;
      
      // Include the value if it matches our criteria and is a valid number
      if (shouldInclude && Number.isInteger(value) && value >= 1 && value <= gridSize) {
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
 * @param {number} row - The row number to get the answers from
 * @returns {Array<Array<number>>} The answers array
 */
function getAnswers(row) {
  const answersSheetName = getAnswersSheetName(row);
  const gridSize = answersSheetName.includes('6') ? GRID_SIZE_6 : GRID_SIZE_4;
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
 * @param {number} row - The row number to get the answers from
 */
function createAnswersSheet(body, row) {
  body.appendPageBreak();
  createSectionHeader(body, 'Solution');
  
  // Get answers from the sheet
  const answersSheetName = getAnswersSheetName(row);
  console.log(`createAnswersSheet: answersSheetName=${answersSheetName}, row=${row}`);
  const answers = getAnswers(row);
  console.log(`createAnswersSheet: answers=${JSON.stringify(answers)}`);
  
  // Create a row for each answer array
  answers.forEach((answerRow, index) => {
    console.log(`Processing answer row ${index + 1}: ${JSON.stringify(answerRow)}`);
    const paragraph = body.appendParagraph('');
    answerRow.forEach(value => {
      console.log(`Processing value ${value} in row ${index + 1}`);
      const url = getImageFromCell(value, answersSheetName, row);
      insertImage(paragraph, url);
    });
    body.appendParagraph(''); // Add spacing between rows
  });
}

/**
 * Creates a SudokuGrid sheet with X's for bold numbers
 * @param {number} row - The row number to get the data from
 */
function createSudokuGrid(row) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const gridSize = getGridSize(row);
    const templateName = gridSize === GRID_SIZE_4 ? 'Template4' : 'Template6';
    const templateSheet = spreadsheet.getSheetByName(templateName);
    if (!templateSheet) {
      throw new Error(`${templateName} sheet not found`);
    }

    // Get the shortname from column A and longname from column B
    const shortname = getSpreadsheet().getRange(row, 1).getValue();
    if (!shortname || typeof shortname !== 'string') {
      throw new Error(`Cell A${row} does not contain a valid shortname`);
    }

    const longname = getSpreadsheet().getRange(row, 2).getValue();
    if (!longname || typeof longname !== 'string') {
      throw new Error(`Cell B${row} does not contain a valid longname`);
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
    const templateRow = templateRange.getRow();
    const templateCol = templateRange.getColumn();

    // Set the longname in the corresponding cell of the new sheet
    sudokuGrid.getRange(templateRow, templateCol).setValue(longname);

    // Get the answers sheet name and data
    const answersSheetName = getAnswersSheetName(row);
    const answersSheet = spreadsheet.getSheetByName(answersSheetName);
    if (!answersSheet) {
      throw new Error(`Sheet "${answersSheetName}" not found for row ${row}`);
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
    console.error(`Error creating SudokuGrid for row ${row}:`, error.message);
    throw error;
  }
}

/**
 * Main function to run all outputs
 */
function main() {
  try {
    const sheet = getSpreadsheet();
    const lastRow = sheet.getLastRow();
    console.log(`Starting main with lastRow=${lastRow}`);
    
    // Process each row that has data
    for (let row = 1; row <= lastRow; row++) {
      console.log(`Processing row ${row} (type: ${typeof row})`);
      // Check if this row has a shortname (column A)
      const shortname = sheet.getRange(row, 1).getValue();
      if (!shortname || typeof shortname !== 'string' || shortname.trim() === '') {
        console.log(`Skipping row ${row} - no shortname found`);
        continue; // Skip rows without a shortname
      }
      
      console.log(`Processing row ${row} with shortname: ${shortname}`);
      
      // Create document for this row
      const body = createDocument(`Mint Hulzo Coin - ${shortname}`);
      const puzzle = getSudokuPuzzle(row);
      
      outputRows(puzzle, body, row);
      outputColumns(puzzle, body, row);
      outputGroups(puzzle, body, row);
      createReferencePage(body, row);
      createAnswersSheet(body, row);
      createSudokuGrid(row);
    }
  } catch (error) {
    console.error('Error processing sudoku:', error.message);
    throw error; // Re-throw to show in Apps Script execution log
  }
}
