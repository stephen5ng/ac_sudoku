// Constants
const GRID_SIZE_6 = 6;
const GRID_SIZE_4 = 4;
const SPREADSHEET_ID = '1t9mwKfa_aPzJwx6qUOgO54N9-1XBQJGCPKY3PpwF-BE';

// Menu configuration
const MENU_NAME = 'Sudoku';
const MENU_ITEMS = [
  {name: 'Generate Puzzles', functionName: 'main'}
];

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
 * Gets or creates the "Generated Files" folder
 * @returns {GoogleAppsScript.Drive.Folder} The folder where generated files should be stored
 */
function getGeneratedFilesFolder() {
  const parentFolderId = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next().getId();
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  
  // Try to find existing "Generated Files" folder
  const folderIterator = parentFolder.getFoldersByName('Generated Files');
  if (folderIterator.hasNext()) {
    return folderIterator.next();
  }
  
  // Create new folder if it doesn't exist
  return parentFolder.createFolder('Generated Files');
}

/**
 * Creates a custom menu in the spreadsheet when it opens
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu(MENU_NAME);
    MENU_ITEMS.forEach(item => menu.addItem(item.name, item.functionName));
    menu.addToUi();
  } catch (error) {
    console.error(`Failed to create menu: ${error.message}`);
  }
}

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
  
  // Count images in the specified row, starting from column E (index 5)
  for (let i = 5; i <= 10; i++) {
    const cell = sheet.getRange(row, i);
    const value = cell.getValue();
    const formula = cell.getFormula();
    
    // Count if cell is an image or has an image formula
    if (String(value) == "CellImage" || formula.toLowerCase().startsWith('=image(')) {
      count++;
    }
  }
  
  // If we have 4 images in columns E-H, use 4x4 grid
  if (count === 4) {
    return GRID_SIZE_4;
  }
  
  // If we have 6 images in columns E-J, use 6x6 grid
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
 * Inserts an image into a paragraph while maintaining aspect ratio
 * @param {GoogleAppsScript.Document.Paragraph} paragraph - The paragraph to insert the image into
 * @param {string} url - The URL of the image to insert
 * @returns {GoogleAppsScript.Document.InlineImage} The inserted image
 */
function insertImage(paragraph, url) {
  try {
    const blob = UrlFetchApp.fetch(url).getBlob();
    const image = paragraph.appendInlineImage(blob);
    
    // Set height to 50 points while maintaining aspect ratio
    const originalWidth = image.getWidth();
    const originalHeight = image.getHeight();
    const aspectRatio = originalWidth / originalHeight;
    const targetHeight = 50; // points
    const targetWidth = targetHeight * aspectRatio;
    
    image.setHeight(targetHeight);
    image.setWidth(targetWidth);
    
    return image;
  } catch (error) {
    console.error(`Error inserting image from URL ${url}:`, error);
    throw new Error(`Failed to insert image: ${error.message}`);
  }
}

/**
 * Gets the name and starting cell of the answers sheet from the Sudokus sheet
 * @param {number} row - The row number to get the sheet info from
 * @returns {{sheetName: string, startCell: string}} Object containing sheet name and starting cell reference
 * @throws {Error} If the sheet info cannot be read
 */
function getAnswersSheetInfo(row) {
  console.log(`getAnswersSheetInfo called with row=${row}, type=${typeof row}`);
  if (!row || typeof row !== 'number' || row < 2) {
    throw new Error(`Invalid row number: ${row} (type: ${typeof row}). Must be row 2 or greater.`);
  }
  
  const sheet = getSpreadsheet();
  const cellValue = sheet.getRange(row, 3).getValue();
  
  if (!cellValue || typeof cellValue !== 'string') {
    throw new Error(`Cell C${row} does not contain a valid sheet reference`);
  }
  
  // Split into sheet name and cell reference (e.g. "Answers6.1!A1" -> ["Answers6.1", "A1"])
  const [sheetName, startCell] = cellValue.split('!');
  if (!sheetName || !startCell) {
    throw new Error(`Cell C${row} does not contain a valid sheet reference format (expected "SheetName!CellReference")`);
  }
  
  return { sheetName, startCell };
}

/**
 * Gets the image URL from a spreadsheet cell
 * @param {number} num - The number to map (1-6 or 1-4)
 * @param {string} answersSheetName - The name of the answers sheet to determine grid size
 * @param {number} row - The row number to get the image from
 * @returns {string} The image URL from the corresponding cell
 * @throws {Error} If the number is out of range, cell doesn't contain an image formula, or referenced cell is empty
 */
function getImageFromCell(num, row) {
  console.log(`getImageFromCell called with num=${num}, row=${row}`);
  const { sheetName: answersSheetName } = getAnswersSheetInfo(row);
  const gridSize = answersSheetName.includes('6') ? GRID_SIZE_6 : GRID_SIZE_4;
  console.log(`Determined gridSize=${gridSize} from answersSheetName`);
  if (num < 1 || num > gridSize) {
    throw new Error(`Invalid number: ${num}. Must be between 1 and ${gridSize} (answersSheetName=${answersSheetName})`);
  }
  
  const sheet = getSpreadsheet();
  // Add 4 to the column to account for the quadruple shift (A=shortname, B=longname, C=sheetname, D=new column)
  const cell = sheet.getRange(row, num + 4);
  
  // Check if the cell contains an image
  const formula = cell.getFormula();
  if (formula == "") {
    return cell.getValue().getContentUrl();
  }

  if (!formula.toLowerCase().startsWith('=image(')) {
    throw new Error(`Cell ${String.fromCharCode(65 + num + 3)}${row} does not contain an image formula`);
  }
  
  // Extract the content from the formula (could be a URL or cell reference)
  const match = formula.match(/=image\(([^)]+)\)/i);
  if (!match) {
    throw new Error(`Invalid image formula in cell ${String.fromCharCode(65 + num + 3)}${row}`);
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
    
    // Move the document to the Generated Files folder
    const folder = getGeneratedFilesFolder();
    const docFile = DriveApp.getFileById(doc.getId());
    const parents = docFile.getParents();
    while (parents.hasNext()) {
      parents.next().removeFile(docFile);
    }
    folder.addFile(docFile);
    
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
        const { sheetName: answersSheetName } = getAnswersSheetInfo(i + 1);
        values.push({
          url: getImageFromCell(array[i][j], i + 1),
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
 * Gets whether the puzzle may only contain the specified values
 * @param {number} row - The row number to check
 * @returns {boolean} Whether the puzzle may only contain the specified values
 */
function getMayOnlyContain(row) {
  const sheet = getSpreadsheet();
  const value = sheet.getRange(row, 4).getValue(); // Column D is the mayOnlyContain column
  return Boolean(value); // Convert any value to boolean
}

/**
 * Gets the answers from the answers sheet
 * @param {string} answersSheetName - The name of the answers sheet
 * @param {string} startCell - The starting cell reference (e.g. "A1")
 * @returns {Array<Array<number>>} The answers array
 */
function getAnswers(answersSheetName, startCell) {
  const gridSize = answersSheetName.includes('6') ? GRID_SIZE_6 : GRID_SIZE_4;
  
  // Parse the starting cell to get row and column offsets
  const startCol = startCell.match(/[A-Z]+/)[0];
  const startRow = parseInt(startCell.match(/\d+/)[0]);
  
  // Calculate the end cell based on grid size and starting position
  const endCol = String.fromCharCode(startCol.charCodeAt(0) + gridSize - 1);
  const endRow = startRow + gridSize - 1;
  const range = `${startCell}:${endCol}${endRow}`;
  
  const values = getSheetData(answersSheetName, range);
  
  // Validate that all values are numbers between 1 and gridSize
  if (!values.every(row => row.every(cell => Number.isInteger(cell) && cell >= 1 && cell <= gridSize))) {
    throw new Error(`Invalid values in "${answersSheetName}" sheet. All values must be integers between 1 and ${gridSize}`);
  }
  
  return values;
}

/**
 * Gets the Sudoku puzzle from the answers sheet, returning either the specified or non-specified values based on mayOnlyContain value
 * @param {number} row - The row number to get the puzzle from
 * @returns {Array<Array<number|null>>} The Sudoku puzzle array
 */
function getSudokuPuzzle(row) {
  const { sheetName: answersSheetName, startCell } = getAnswersSheetInfo(row);
  const gridSize = getGridSize(row);
  
  // Parse the starting cell to get row and column offsets
  const startCol = startCell.match(/[A-Z]+/)[0];
  const startRow = parseInt(startCell.match(/\d+/)[0]);
  
  // Calculate the end cell based on grid size and starting position
  const endCol = String.fromCharCode(startCol.charCodeAt(0) + gridSize - 1);
  const endRow = startRow + gridSize - 1;
  const range = `${startCell}:${endCol}${endRow}`;
  
  const values = getSheetData(answersSheetName, range);
  const sheet = getSheetByName(answersSheetName);
  const mayOnlyContain = getMayOnlyContain(row);
  
  // Create the Sudoku puzzle array
  const puzzle = [];
  
  // Process each row
  for (let i = 0; i < gridSize; i++) {
    const row = [];
    for (let j = 0; j < gridSize; j++) {
      const value = values[i][j];
      
      // Check if the cell is specified by getting its font weight
      let isSpecified = false;
      try {
        // Adjust cell coordinates based on starting position
        const cell = sheet.getRange(startRow + i, startCol.charCodeAt(0) - 64 + j);
        const fontWeight = cell.getFontWeight();
        isSpecified = fontWeight === "bold";
      } catch (e) {
        console.log(`Could not get font weight for cell (${startRow + i}, ${startCol.charCodeAt(0) - 64 + j}): ${e.message}`);
      }
      
      // If mayOnlyContain is true, we want specified values
      // If mayOnlyContain is false, we want non-specified values
      const shouldInclude = mayOnlyContain ? isSpecified : !isSpecified;
      
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
 * Gets the section title based on type and mayOnlyContain value
 * @param {string} sectionType - The type of section (ROWS, COLUMNS, GROUPS)
 * @param {number} row - The row number to get the mayOnlyContain value from
 * @returns {string} The formatted section title
 */
function getSectionTitle(sectionType, row) {
  const mayOnlyContain = getMayOnlyContain(row);
  const declaration = mayOnlyContain ? MAY_ONLY_CONTAIN : MUST_NOT_CONTAIN;
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
  const { sheetName: answersSheetName } = getAnswersSheetInfo(currentRow);
  const sections = sudokuArray.map(sudokuRow => 
    sudokuRow
      .map(value => {
        if (value === null) return null;
        const url = getImageFromCell(value, currentRow);
        return url ? { url, value } : null;
      })
      .filter(Boolean)
  );
  outputSection(body, getSectionTitle(SECTION_TYPES.ROWS, currentRow), sections, 'ROW');
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
  const { sheetName: answersSheetName } = getAnswersSheetInfo(currentRow);
  const sections = Array.from({ length: gridSize }, (_, j) =>
    sudokuArray
      .map(sudokuRow => {
        const value = sudokuRow[j];
        if (value === null) return null;
        const url = getImageFromCell(value, currentRow);
        return url ? { url, value } : null;
      })
      .filter(Boolean)
  );
  outputSection(body, getSectionTitle(SECTION_TYPES.COLUMNS, currentRow), sections, 'COLUMN');
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
  const { sheetName: answersSheetName } = getAnswersSheetInfo(currentRow);
  const sections = groupBoundaries.map(boundaries => {
    const values = [];
    for (let i = boundaries.rowStart; i <= boundaries.rowEnd; i++) {
      for (let j = boundaries.colStart; j <= boundaries.colEnd; j++) {
        const value = sudokuArray[i][j];
        if (value !== null) {
          const url = getImageFromCell(value, currentRow);
          if (url) {
            values.push({ url, value });
          }
        }
      }
    }
    return values;
  });
  outputSection(body, getSectionTitle(SECTION_TYPES.GROUPS, currentRow), sections, 'GROUP');
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
  const { sheetName: answersSheetName } = getAnswersSheetInfo(row);
  
  // Create a paragraph for each number (1-4 or 1-6)
  for (let num = 1; num <= gridSize; num++) {
    const paragraph = body.appendParagraph('');
    const url = getImageFromCell(num, row);
    
    // Insert copies of the same image
    for (let i = 0; i < gridSize; i++) {
      insertImage(paragraph, url);
    }
    
    body.appendParagraph(''); // Add spacing between rows
  }
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
 * Creates the answers sheet with the complete solution
 * @param {GoogleAppsScript.Document.Body} body - The document body
 * @param {number} row - The row number to get the answers from
 */
function createAnswersSheet(body, row) {
  body.appendPageBreak();
  createSectionHeader(body, 'Solution');
  
  // Get answers from the sheet
  const { sheetName: answersSheetName, startCell } = getAnswersSheetInfo(row);
  console.log(`createAnswersSheet: answersSheetName=${answersSheetName}, startCell=${startCell}, row=${row}`);
  const answers = getAnswers(answersSheetName, startCell);
  console.log(`createAnswersSheet: answers=${JSON.stringify(answers)}`);
  
  // Create a row for each answer array
  answers.forEach((answerRow, index) => {
    console.log(`Processing answer row ${index + 1}: ${JSON.stringify(answerRow)}`);
    const paragraph = body.appendParagraph('');
    answerRow.forEach(value => {
      console.log(`Processing value ${value} in row ${index + 1}`);
      const url = getImageFromCell(value, row);
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
    const targetSpreadsheetId = '1JB2VLOx1DuzSHr4FdMfGfLMmaXiXkGxb1jM3_SwZStM';
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
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

    // Delete the sheet if it already exists in the target spreadsheet
    const existingSheet = targetSpreadsheet.getSheetByName(shortname);
    if (existingSheet) {
      targetSpreadsheet.deleteSheet(existingSheet);
    }

    // Create a copy of the template in the target spreadsheet
    const sudokuGrid = templateSheet.copyTo(targetSpreadsheet);
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
    const answersSheetName = getAnswersSheetInfo(row).sheetName;
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
    
    // Process each row until we hit an empty row, starting after the header
    for (let row = 2; row <= lastRow; row++) {
      console.log(`Processing row ${row} (type: ${typeof row})`);
      // Check if this row has a shortname (column A)
      const shortname = sheet.getRange(row, 1).getValue();
      if (!shortname || typeof shortname !== 'string' || shortname.trim() === '') {
        console.log(`Stopping at row ${row} - no shortname found`);
        break; // Stop processing when we hit an empty row
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
