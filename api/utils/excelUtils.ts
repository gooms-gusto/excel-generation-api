import ExcelJS from 'exceljs';

// Interface definitions
export interface MappingConfig {
  sheet: string;
  cell: string;
  fieldName: string;
  style?: {
    bgColor?: string;
    fontColor?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
  };
  formula?: string;
  format?: string;
}

export interface TableConfig {
  sheet: string;
  tableName: string;
  startCell: string;
  columns: string[];
  style?: {
    headerBgColor?: string;
    headerFontColor?: string;
    rowBgColor?: string;
    alternateRowBgColor?: string;
  };
}

export interface ExcelOptions {
  includeHeaders?: boolean;
  autoFitColumns?: boolean;
  freezeFirstRow?: boolean;
  protectSheet?: boolean;
  password?: string;
}

// Utility function to get nested object value by path
const getNestedValue = (obj: any, path: string): any => {
  return path.split('.').reduce((current, key) => {
    return current && current[key] !== undefined ? current[key] : null;
  }, obj);
};

// Utility function to convert column letter to number
const columnLetterToNumber = (letter: string): number => {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result;
};

// Utility function to convert number to column letter
const numberToColumnLetter = (num: number): string => {
  let result = '';
  while (num > 0) {
    num--;
    result = String.fromCharCode('A'.charCodeAt(0) + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
};

// Apply cell styling
const applyCellStyle = (cell: ExcelJS.Cell, style: MappingConfig['style']) => {
  if (!style) return;

  // Background color
  if (style.bgColor) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${style.bgColor}` }
    };
  }

  // Font styling
  const fontOptions: Partial<ExcelJS.Font> = {};
  if (style.fontColor) fontOptions.color = { argb: `FF${style.fontColor}` };
  if (style.fontSize) fontOptions.size = style.fontSize;
  if (style.bold) fontOptions.bold = true;
  if (style.italic) fontOptions.italic = true;
  if (style.underline) fontOptions.underline = true;

  if (Object.keys(fontOptions).length > 0) {
    cell.font = { ...cell.font, ...fontOptions };
  }
};

// Apply cell format
const applyCellFormat = (cell: ExcelJS.Cell, format?: string) => {
  if (!format) return;

  switch (format.toLowerCase()) {
    case 'currency':
      cell.numFmt = '$#,##0.00';
      break;
    case 'percentage':
      cell.numFmt = '0.00%';
      break;
    case 'date':
      cell.numFmt = 'mm/dd/yyyy';
      break;
    case 'datetime':
      cell.numFmt = 'mm/dd/yyyy hh:mm:ss';
      break;
    case 'number':
      cell.numFmt = '#,##0.00';
      break;
    case 'integer':
      cell.numFmt = '#,##0';
      break;
    default:
      cell.numFmt = format;
  }
};

// Create table from JSON data
const createTable = (
  worksheet: ExcelJS.Worksheet,
  tableConfig: TableConfig,
  tableData: any[]
) => {
  if (!tableData || tableData.length === 0) return;

  const startCellMatch = tableConfig.startCell.match(/^([A-Z]+)([0-9]+)$/);
  if (!startCellMatch) throw new Error(`Invalid start cell: ${tableConfig.startCell}`);

  const startCol = columnLetterToNumber(startCellMatch[1]);
  const startRow = parseInt(startCellMatch[2]);

  // Add headers
  tableConfig.columns.forEach((header, index) => {
    const cell = worksheet.getCell(startRow, startCol + index);
    cell.value = header;
    
    // Apply header styling
    if (tableConfig.style?.headerBgColor) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${tableConfig.style.headerBgColor}` }
      };
    }
    if (tableConfig.style?.headerFontColor) {
      cell.font = {
        color: { argb: `FF${tableConfig.style.headerFontColor}` },
        bold: true
      };
    } else {
      cell.font = { bold: true };
    }
  });

  // Add data rows
  tableData.forEach((row, rowIndex) => {
    const currentRow = startRow + rowIndex + 1;
    
    if (Array.isArray(row)) {
      // Handle array data
      row.forEach((cellValue, colIndex) => {
        if (colIndex < tableConfig.columns.length) {
          const cell = worksheet.getCell(currentRow, startCol + colIndex);
          cell.value = cellValue;
          
          // Apply alternating row colors
          if (tableConfig.style?.rowBgColor || tableConfig.style?.alternateRowBgColor) {
            const bgColor = rowIndex % 2 === 0 
              ? tableConfig.style?.rowBgColor 
              : tableConfig.style?.alternateRowBgColor || tableConfig.style?.rowBgColor;
            
            if (bgColor) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: `FF${bgColor}` }
              };
            }
          }
        }
      });
    } else if (typeof row === 'object') {
      // Handle object data
      tableConfig.columns.forEach((column, colIndex) => {
        const cell = worksheet.getCell(currentRow, startCol + colIndex);
        cell.value = row[column] || '';
        
        // Apply alternating row colors
        if (tableConfig.style?.rowBgColor || tableConfig.style?.alternateRowBgColor) {
          const bgColor = rowIndex % 2 === 0 
            ? tableConfig.style?.rowBgColor 
            : tableConfig.style?.alternateRowBgColor || tableConfig.style?.rowBgColor;
          
          if (bgColor) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: `FF${bgColor}` }
            };
          }
        }
      });
    }
  });

  // Create Excel table
  const endCol = numberToColumnLetter(startCol + tableConfig.columns.length - 1);
  const endRow = startRow + tableData.length;
  const tableRef = `${tableConfig.startCell}:${endCol}${endRow}`;

  worksheet.addTable({
    name: tableConfig.tableName,
    ref: tableRef,
    headerRow: true,
    style: {
      theme: 'TableStyleMedium9',
      showRowStripes: true,
    },
    columns: tableConfig.columns.map(col => ({ name: col })),
    rows: tableData.map(row => 
      Array.isArray(row) ? row : tableConfig.columns.map(col => row[col] || '')
    )
  });
};

// Auto-fit columns
const autoFitColumns = (worksheet: ExcelJS.Worksheet) => {
  try {
    if (!worksheet) return;
    
    // Simple approach: set reasonable default widths for first few columns
    for (let colNumber = 1; colNumber <= 10; colNumber++) {
      const column = worksheet.getColumn(colNumber);
      if (column) {
        column.width = 15; // Set a reasonable default width
      }
    }
  } catch (error) {
    console.warn('Auto-fit columns failed:', error);
  }
};

// Main function to generate workbook
export const generateWorkbook = async (
  jsonData: any,
  mappingConfig: MappingConfig[],
  tableConfigs: TableConfig[] = [],
  templateBuffer?: Buffer,
  options: ExcelOptions = {}
): Promise<Buffer> => {
  const workbook = new ExcelJS.Workbook();

  try {
    // Load template or create new workbook
    if (templateBuffer) {
      await workbook.xlsx.load(templateBuffer);
    } else {
      workbook.addWorksheet('Sheet1');
    }

    // Process cell mappings
    for (const mapping of mappingConfig) {
      let worksheet = workbook.getWorksheet(mapping.sheet);
      if (!worksheet) {
        worksheet = workbook.addWorksheet(mapping.sheet);
      }

      const cell = worksheet.getCell(mapping.cell);
      
      // Set cell value
      const value = getNestedValue(jsonData, mapping.fieldName);
      
      if (mapping.formula) {
        // Set formula
        cell.value = { formula: mapping.formula, result: value || 0 };
      } else {
        cell.value = value !== null ? value : '';
      }

      // Apply styling
      applyCellStyle(cell, mapping.style);
      
      // Apply formatting
      applyCellFormat(cell, mapping.format);
    }

    // Create tables
    for (const tableConfig of tableConfigs) {
      let worksheet = workbook.getWorksheet(tableConfig.sheet);
      if (!worksheet) {
        worksheet = workbook.addWorksheet(tableConfig.sheet);
      }

      const tableData = getNestedValue(jsonData, 'tableData') || jsonData.tableData;
      if (tableData) {
        createTable(worksheet, tableConfig, tableData);
      }
    }

    // Apply options
    workbook.eachSheet((worksheet) => {
      if (!worksheet) return;
      
      // Auto-fit columns
      if (options.autoFitColumns !== false) {
        try {
          autoFitColumns(worksheet);
        } catch (error) {
          console.warn('Auto-fit columns failed:', error);
        }
      }

      // Freeze first row
      if (options.freezeFirstRow) {
        worksheet.views = [{ state: 'frozen', ySplit: 1 }];
      }

      // Protect sheet
      if (options.protectSheet) {
        worksheet.protect(options.password || '', {
          selectLockedCells: true,
          selectUnlockedCells: true
        });
      }
    });

    // Generate buffer
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);

  } catch (error) {
    console.error('Error generating workbook:', error);
    throw new Error(`Failed to generate Excel workbook: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
};

// Function to validate template
export const validateTemplate = async (templateBuffer: Buffer): Promise<{
  isValid: boolean;
  sheets: string[];
  error?: string;
}> => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBuffer);
    
    const sheets: string[] = [];
    workbook.eachSheet((worksheet) => {
      sheets.push(worksheet.name);
    });

    return {
      isValid: true,
      sheets
    };
  } catch (error) {
    return {
      isValid: false,
      sheets: [],
      error: error instanceof Error ? error.message : 'Invalid template file'
    };
  }
};

// Function to get template info
export const getTemplateInfo = async (templateBuffer: Buffer): Promise<{
  sheets: Array<{
    name: string;
    rowCount: number;
    columnCount: number;
    hasData: boolean;
  }>;
  totalSheets: number;
}> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  
  const sheets: Array<{
    name: string;
    rowCount: number;
    columnCount: number;
    hasData: boolean;
  }> = [];

  workbook.eachSheet((worksheet) => {
    const rowCount = worksheet.rowCount;
    const columnCount = worksheet.columnCount;
    let hasData = false;

    // Check if sheet has any data
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
          hasData = true;
        }
      });
    });

    sheets.push({
      name: worksheet.name,
      rowCount,
      columnCount,
      hasData
    });
  });

  return {
    sheets,
    totalSheets: sheets.length
  };
};
