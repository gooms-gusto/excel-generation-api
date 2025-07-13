import { Request, Response } from 'express';
import { generateWorkbook, validateTemplate as validateTemplateUtil, getTemplateInfo } from '../utils/excelUtils';
import { sendEmailWithAttachment, testEmailConfiguration } from '../services/emailService';

// Generate Excel file (main endpoint)
export const generateExcel = async (req: Request, res: Response) => {
  try {
    const { 
      jsonData, 
      mappingConfig, 
      tables = [], 
      mode = 'download', 
      emailAddress, 
      fileName = 'generated.xlsx',
      options = {}
    } = req.body;

    const templateFile = req.file ? req.file.buffer : undefined;

    // Generate the Excel workbook
    const workbookBuffer = await generateWorkbook(
      jsonData,
      mappingConfig,
      tables,
      templateFile,
      options
    );

    if (mode === 'download') {
      // Set headers for file download
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="${fileName}"`
      );
      res.setHeader('Content-Length', workbookBuffer.length);
      
      return res.send(workbookBuffer);
    }

    if (mode === 'email') {
      if (!emailAddress) {
        return res.status(400).json({ 
          error: 'Email address required for email mode',
          message: 'Please provide a valid email address in the emailAddress field'
        });
      }

      // Send email with attachment
      await sendEmailWithAttachment(emailAddress, workbookBuffer, fileName);
      
      return res.json({ 
        success: true,
        message: 'Excel file generated and sent successfully',
        details: {
          recipient: emailAddress,
          fileName: fileName,
          fileSize: `${Math.round(workbookBuffer.length / 1024)} KB`,
          timestamp: new Date().toISOString()
        }
      });
    }

    // Default: return file info
    return res.json({ 
      success: true,
      message: 'Excel file generated successfully',
      data: {
        fileName: fileName,
        fileSize: `${Math.round(workbookBuffer.length / 1024)} KB`,
        base64: workbookBuffer.toString('base64')
      }
    });

  } catch (error) {
    console.error('Excel generation error:', error);
    return res.status(500).json({ 
      error: 'Failed to generate Excel file',
      message: error instanceof Error ? error.message : 'Unknown error occurred',
      timestamp: new Date().toISOString()
    });
  }
};

// Upload and validate template
export const uploadTemplate = async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        error: 'No template file provided',
        message: 'Please upload an Excel template file (.xlsx or .xls)'
      });
    }

    const templateBuffer = req.file.buffer;
    const validation = await validateTemplateUtil(templateBuffer);
    
    if (!validation.isValid) {
      return res.status(400).json({
        error: 'Invalid template file',
        message: validation.error,
        details: 'Please ensure the file is a valid Excel workbook'
      });
    }

    const templateInfo = await getTemplateInfo(templateBuffer);

    return res.json({
      success: true,
      message: 'Template uploaded and validated successfully',
      template: {
        fileName: req.file.originalname,
        fileSize: `${Math.round(req.file.size / 1024)} KB`,
        sheets: validation.sheets,
        details: templateInfo,
        uploadedAt: new Date().toISOString()
      }
    });

  } catch (error) {
    console.error('Template upload error:', error);
    return res.status(500).json({
      error: 'Failed to process template',
      message: error instanceof Error ? error.message : 'Unknown error occurred'
    });
  }
};

// Generate and download Excel (dedicated download endpoint)
export const downloadExcel = async (req: Request, res: Response) => {
  // Set mode to download and call main generate function
  req.body.mode = 'download';
  return generateExcel(req, res);
};

// Generate and email Excel (dedicated email endpoint)
export const emailExcel = async (req: Request, res: Response) => {
  // Set mode to email and call main generate function
  req.body.mode = 'email';
  
  if (!req.body.emailAddress) {
    return res.status(400).json({
      error: 'Email address required',
      message: 'Please provide a valid email address'
    });
  }

  return generateExcel(req, res);
};

// Validate template file
export const validateTemplate = async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        error: 'No template file provided',
        message: 'Please upload an Excel template file for validation'
      });
    }

    const templateBuffer = req.file.buffer;
    const validation = await validateTemplateUtil(templateBuffer);
    
    if (!validation.isValid) {
      return res.status(400).json({
        success: false,
        error: 'Template validation failed',
        message: validation.error,
        details: {
          fileName: req.file.originalname,
          fileSize: `${Math.round(req.file.size / 1024)} KB`,
          validatedAt: new Date().toISOString()
        }
      });
    }

    const templateInfo = await getTemplateInfo(templateBuffer);

    return res.json({
      success: true,
      message: 'Template is valid and ready to use',
      validation: {
        isValid: true,
        sheets: validation.sheets,
        details: templateInfo,
        fileName: req.file.originalname,
        fileSize: `${Math.round(req.file.size / 1024)} KB`,
        validatedAt: new Date().toISOString()
      }
    });

  } catch (error) {
    console.error('Template validation error:', error);
    return res.status(500).json({
      success: false,
      error: 'Validation failed',
      message: error instanceof Error ? error.message : 'Unknown error occurred'
    });
  }
};

// Get Excel information and capabilities
export const getExcelInfo = async (req: Request, res: Response) => {
  try {
    const { templateId } = req.params;

    // Test email configuration
    const emailTest = await testEmailConfiguration();

    const info: any = {
      service: 'Excel Generation API',
      version: '1.0.0',
      capabilities: {
        fileFormats: ['.xlsx', '.xls'],
        features: [
          'Flexible JSON to Excel mapping',
          'Multiple sheet support',
          'Custom cell styling (colors, fonts, formatting)',
          'Formula support',
          'Table creation from arrays',
          'Template-based generation',
          'Download and email delivery',
          'Auto-fit columns',
          'Sheet protection',
          'Freeze panes'
        ],
        styling: {
          supportedColors: 'Hex colors (6-digit format)',
          fontSizes: '8-72 points',
          fontStyles: ['bold', 'italic', 'underline'],
          cellFormats: ['currency', 'percentage', 'date', 'datetime', 'number', 'integer', 'custom']
        },
        limits: {
          maxFileSize: process.env.MAX_FILE_SIZE || '10MB',
          maxSheets: 'Unlimited',
          maxCells: 'Excel limits apply',
          maxTableRows: 'Memory dependent'
        }
      },
      emailConfiguration: {
        configured: emailTest.success,
        status: emailTest.message
      },
      endpoints: {
        generate: 'POST /api/excel/generate',
        download: 'POST /api/excel/download',
        email: 'POST /api/excel/email',
        uploadTemplate: 'POST /api/excel/upload-template',
        validateTemplate: 'POST /api/excel/validate-template',
        examples: 'GET /api/excel/examples'
      },
      timestamp: new Date().toISOString()
    };

    if (templateId) {
      info.templateId = templateId;
      info.message = `Information for template: ${templateId}`;
    }

    return res.json(info);

  } catch (error) {
    console.error('Get info error:', error);
    return res.status(500).json({
      error: 'Failed to retrieve service information',
      message: error instanceof Error ? error.message : 'Unknown error occurred'
    });
  }
};

// Bulk generation (bonus endpoint for multiple files)
export const bulkGenerate = async (req: Request, res: Response) => {
  try {
    const { requests } = req.body; // Array of generation requests
    
    if (!Array.isArray(requests) || requests.length === 0) {
      return res.status(400).json({
        error: 'Invalid bulk request',
        message: 'Please provide an array of generation requests'
      });
    }

    const results = [];
    const errors = [];

    for (let i = 0; i < requests.length; i++) {
      try {
        const request = requests[i];
        const workbookBuffer = await generateWorkbook(
          request.jsonData,
          request.mappingConfig,
          request.tables || [],
          undefined, // No template support in bulk for now
          request.options || {}
        );

        results.push({
          index: i,
          success: true,
          fileName: request.fileName || `generated_${i + 1}.xlsx`,
          fileSize: `${Math.round(workbookBuffer.length / 1024)} KB`,
          data: workbookBuffer.toString('base64')
        });

      } catch (error) {
        errors.push({
          index: i,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    return res.json({
      success: true,
      message: `Bulk generation completed. ${results.length} successful, ${errors.length} failed.`,
      results,
      errors,
      summary: {
        total: requests.length,
        successful: results.length,
        failed: errors.length,
        timestamp: new Date().toISOString()
      }
    });

  } catch (error) {
    console.error('Bulk generation error:', error);
    return res.status(500).json({
      error: 'Bulk generation failed',
      message: error instanceof Error ? error.message : 'Unknown error occurred'
    });
  }
};

export default {
  generateExcel,
  uploadTemplate,
  downloadExcel,
  emailExcel,
  validateTemplate,
  getExcelInfo,
  bulkGenerate
};
