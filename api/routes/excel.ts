import { Router } from 'express';
import { 
  generateExcel, 
  uploadTemplate, 
  downloadExcel, 
  emailExcel,
  validateTemplate,
  getExcelInfo,
  bulkGenerate
} from '../controllers/excelController';
import upload from '../middlewares/upload';
import { validateExcelRequest } from '../middlewares/validation';

const router = Router();

/**
 * @route POST /api/excel/generate
 * @desc Generate Excel file from JSON data with flexible mapping
 * @access Public
 * @example
 * POST /api/excel/generate
 * Content-Type: multipart/form-data or application/json
 * 
 * Body (JSON):
 * {
 *   "jsonData": { "name": "John", "age": 30 },
 *   "mappingConfig": [
 *     {
 *       "sheet": "Sheet1",
 *       "cell": "A1",
 *       "fieldName": "name",
 *       "style": { "bgColor": "FFFF00", "fontColor": "000000" }
 *     }
 *   ],
 *   "mode": "download", // or "email"
 *   "emailAddress": "user@example.com" // required if mode is "email"
 * }
 * 
 * With template file:
 * - Include template file as 'template' in form-data
 */
router.post('/generate', upload.single('template'), validateExcelRequest, generateExcel);

/**
 * @route POST /api/excel/upload-template
 * @desc Upload and validate Excel template
 * @access Public
 * @example
 * POST /api/excel/upload-template
 * Content-Type: multipart/form-data
 * 
 * Form data:
 * - template: Excel file (.xlsx, .xls)
 */
router.post('/upload-template', upload.single('template'), uploadTemplate);

/**
 * @route POST /api/excel/download
 * @desc Generate and download Excel file
 * @access Public
 * @example
 * POST /api/excel/download
 * Content-Type: application/json
 * 
 * Body: Same as /generate but mode is automatically set to "download"
 */
router.post('/download', upload.single('template'), validateExcelRequest, downloadExcel);

/**
 * @route POST /api/excel/email
 * @desc Generate Excel file and send via email
 * @access Public
 * @example
 * POST /api/excel/email
 * Content-Type: application/json
 * 
 * Body: Same as /generate but mode is automatically set to "email"
 * emailAddress is required
 */
router.post('/email', upload.single('template'), validateExcelRequest, emailExcel);

/**
 * @route POST /api/excel/validate-template
 * @desc Validate uploaded Excel template
 * @access Public
 * @example
 * POST /api/excel/validate-template
 * Content-Type: multipart/form-data
 * 
 * Form data:
 * - template: Excel file (.xlsx, .xls)
 */
router.post('/validate-template', upload.single('template'), validateTemplate);

/**
 * @route GET /api/excel/info/:templateId?
 * @desc Get information about Excel template or general Excel capabilities
 * @access Public
 * @example
 * GET /api/excel/info
 * GET /api/excel/info/template123
 */
router.get('/info/:templateId?', getExcelInfo);

/**
 * @route POST /api/excel/bulk
 * @desc Generate multiple Excel files in one request
 * @access Public
 * @example
 * POST /api/excel/bulk
 * Content-Type: application/json
 * 
 * Body:
 * {
 *   "requests": [
 *     {
 *       "jsonData": {...},
 *       "mappingConfig": [...],
 *       "fileName": "file1.xlsx"
 *     },
 *     {
 *       "jsonData": {...},
 *       "mappingConfig": [...],
 *       "fileName": "file2.xlsx"
 *     }
 *   ]
 * }
 */
router.post('/bulk', bulkGenerate);

/**
 * @route GET /api/excel/examples
 * @desc Get example requests and responses
 * @access Public
 */
router.get('/examples', (req, res) => {
  res.json({
    examples: {
      basicGeneration: {
        url: '/api/excel/generate',
        method: 'POST',
        body: {
          jsonData: {
            name: 'John Doe',
            age: 30,
            email: 'john@example.com',
            tableData: [
              ['Product', 'Price', 'Quantity'],
              ['Laptop', 999.99, 5],
              ['Mouse', 29.99, 10]
            ]
          },
          mappingConfig: [
            {
              sheet: 'Sheet1',
              cell: 'A1',
              fieldName: 'name',
              style: { bgColor: 'FFFF00', fontColor: '000000' }
            },
            {
              sheet: 'Sheet1',
              cell: 'B1',
              fieldName: 'age',
              formula: 'SUM(B2:B10)'
            }
          ],
          tables: [
            {
              sheet: 'Sheet1',
              tableName: 'ProductTable',
              startCell: 'A3',
              columns: ['Product', 'Price', 'Quantity']
            }
          ],
          mode: 'download'
        }
      },
      withTemplate: {
        url: '/api/excel/generate',
        method: 'POST',
        contentType: 'multipart/form-data',
        formData: {
          template: 'Excel file',
          jsonData: '{"name": "Jane Doe", "company": "ABC Corp"}',
          mappingConfig: '[{"sheet": "Sheet1", "cell": "A1", "fieldName": "name"}]',
          mode: 'email',
          emailAddress: 'jane@example.com'
        }
      },
      emailGeneration: {
        url: '/api/excel/email',
        method: 'POST',
        body: {
          jsonData: { name: 'Alice', department: 'Engineering' },
          mappingConfig: [
            {
              sheet: 'Report',
              cell: 'A1',
              fieldName: 'name',
              style: { bgColor: 'E6F3FF', fontColor: '0066CC' }
            }
          ],
          emailAddress: 'alice@company.com'
        }
      }
    },
    supportedFeatures: [
      'Flexible cell mapping',
      'Multiple sheets support',
      'Custom styling (background color, font color)',
      'Formula support',
      'Table creation from JSON arrays',
      'Template-based generation',
      'Download or email delivery',
      'Multiple file formats (.xlsx, .xls)'
    ]
  });
});

export default router;
