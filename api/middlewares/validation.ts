import { Request, Response, NextFunction } from 'express';
import Joi from 'joi';

// Schema for mapping configuration
const mappingConfigSchema = Joi.object({
  sheet: Joi.string().required().description('Sheet name'),
  cell: Joi.string().required().pattern(/^[A-Z]+[0-9]+$/).description('Cell reference (e.g., A1, B2)'),
  fieldName: Joi.string().required().description('JSON field name to map'),
  style: Joi.object({
    bgColor: Joi.string().pattern(/^[0-9A-Fa-f]{6}$/).description('Background color in hex (without #)'),
    fontColor: Joi.string().pattern(/^[0-9A-Fa-f]{6}$/).description('Font color in hex (without #)'),
    fontSize: Joi.number().min(8).max(72).description('Font size'),
    bold: Joi.boolean().description('Bold text'),
    italic: Joi.boolean().description('Italic text'),
    underline: Joi.boolean().description('Underline text')
  }).optional(),
  formula: Joi.string().optional().description('Excel formula (e.g., SUM(A1:A10))'),
  format: Joi.string().optional().description('Cell format (e.g., currency, percentage)')
});

// Schema for table configuration
const tableConfigSchema = Joi.object({
  sheet: Joi.string().required().description('Sheet name'),
  tableName: Joi.string().required().description('Table name'),
  startCell: Joi.string().required().pattern(/^[A-Z]+[0-9]+$/).description('Starting cell for table'),
  columns: Joi.array().items(Joi.string()).required().description('Column headers'),
  style: Joi.object({
    headerBgColor: Joi.string().pattern(/^[0-9A-Fa-f]{6}$/),
    headerFontColor: Joi.string().pattern(/^[0-9A-Fa-f]{6}$/),
    rowBgColor: Joi.string().pattern(/^[0-9A-Fa-f]{6}$/),
    alternateRowBgColor: Joi.string().pattern(/^[0-9A-Fa-f]{6}$/)
  }).optional()
});

// Main request schema
const excelRequestSchema = Joi.object({
  jsonData: Joi.object().required().description('JSON data to be mapped to Excel'),
  mappingConfig: Joi.array().items(mappingConfigSchema).required().description('Array of mapping configurations'),
  tables: Joi.array().items(tableConfigSchema).optional().description('Table configurations'),
  mode: Joi.string().valid('download', 'email').default('download').description('Output mode'),
  emailAddress: Joi.when('mode', {
    is: 'email',
    then: Joi.string().email().required(),
    otherwise: Joi.string().email().optional()
  }).description('Email address (required for email mode)'),
  fileName: Joi.string().optional().default('generated.xlsx').description('Output file name'),
  options: Joi.object({
    includeHeaders: Joi.boolean().default(true),
    autoFitColumns: Joi.boolean().default(true),
    freezeFirstRow: Joi.boolean().default(false),
    protectSheet: Joi.boolean().default(false),
    password: Joi.string().optional()
  }).optional()
});

// Validation middleware
export const validateExcelRequest = (req: Request, res: Response, next: NextFunction) => {
  try {
    // Parse JSON strings if they come from form-data
    if (typeof req.body.jsonData === 'string') {
      req.body.jsonData = JSON.parse(req.body.jsonData);
    }
    if (typeof req.body.mappingConfig === 'string') {
      req.body.mappingConfig = JSON.parse(req.body.mappingConfig);
    }
    if (typeof req.body.tables === 'string') {
      req.body.tables = JSON.parse(req.body.tables);
    }
    if (typeof req.body.options === 'string') {
      req.body.options = JSON.parse(req.body.options);
    }

    const { error, value } = excelRequestSchema.validate(req.body, {
      abortEarly: false,
      stripUnknown: true
    });

    if (error) {
      return res.status(400).json({
        error: 'Validation Error',
        details: error.details.map(detail => ({
          field: detail.path.join('.'),
          message: detail.message,
          value: detail.context?.value
        }))
      });
    }

    // Replace request body with validated and sanitized data
    req.body = value;
    next();
  } catch (parseError) {
    return res.status(400).json({
      error: 'Invalid JSON format',
      message: 'Please ensure all JSON fields are properly formatted'
    });
  }
};

// Template validation middleware
export const validateTemplate = (req: Request, res: Response, next: NextFunction) => {
  if (!req.file) {
    return res.status(400).json({
      error: 'Template file required',
      message: 'Please upload an Excel template file (.xlsx or .xls)'
    });
  }

  // Additional file validation
  const maxSize = parseInt(process.env.MAX_FILE_SIZE || '10485760'); // 10MB
  if (req.file.size > maxSize) {
    return res.status(413).json({
      error: 'File too large',
      message: `File size must be less than ${Math.round(maxSize / 1024 / 1024)}MB`
    });
  }

  next();
};

// Email validation middleware
export const validateEmailRequest = (req: Request, res: Response, next: NextFunction) => {
  const emailSchema = Joi.object({
    emailAddress: Joi.string().email().required().description('Recipient email address')
  });

  const { error } = emailSchema.validate({ emailAddress: req.body.emailAddress });
  
  if (error) {
    return res.status(400).json({
      error: 'Invalid email address',
      message: 'Please provide a valid email address'
    });
  }

  next();
};

export { mappingConfigSchema, tableConfigSchema, excelRequestSchema };
