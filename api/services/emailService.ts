import nodemailer from 'nodemailer';
import { createTransport, Transporter } from 'nodemailer';
import dotenv from 'dotenv';

dotenv.config();

// Email configuration interface
interface EmailConfig {
  host: string;
  port: number;
  secure: boolean;
  auth: {
    user: string;
    pass: string;
  };
}

// Email options interface
interface EmailOptions {
  to: string;
  subject?: string;
  text?: string;
  html?: string;
  attachments?: Array<{
    filename: string;
    content: Buffer;
    contentType?: string;
  }>;
}

// Create email transporter
const createEmailTransporter = (): Transporter => {
  const config: EmailConfig = {
    host: process.env.EMAIL_HOST || 'smtp.gmail.com',
    port: parseInt(process.env.EMAIL_PORT || '587'),
    secure: process.env.EMAIL_SECURE === 'true',
    auth: {
      user: process.env.EMAIL_USER || '',
      pass: process.env.EMAIL_PASS || ''
    }
  };

  if (!config.auth.user || !config.auth.pass) {
    throw new Error('Email credentials not configured. Please set EMAIL_USER and EMAIL_PASS in .env file');
  }

  return createTransport(config);
};

// Send email with Excel attachment
export const sendEmailWithAttachment = async (
  recipient: string,
  attachmentBuffer: Buffer,
  fileName: string = 'generated.xlsx',
  customSubject?: string,
  customMessage?: string
): Promise<void> => {
  try {
    const transporter = createEmailTransporter();

    const subject = customSubject || 'Your Generated Excel File';
    const textMessage = customMessage || `
Hello,

Please find attached the Excel file generated from your JSON data.

The file contains:
- Your mapped data in the specified cells
- Custom styling and formatting as requested
- Any tables or formulas you configured

If you have any questions or need modifications, please don't hesitate to reach out.

Best regards,
Excel Generation Service
    `.trim();

    const htmlMessage = `
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header { background-color: #f8f9fa; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
        .content { background-color: #ffffff; padding: 20px; border: 1px solid #dee2e6; border-radius: 5px; }
        .footer { margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; font-size: 12px; color: #6c757d; }
        .highlight { background-color: #e7f3ff; padding: 10px; border-left: 4px solid #007bff; margin: 15px 0; }
        ul { padding-left: 20px; }
        li { margin-bottom: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2 style="margin: 0; color: #007bff;">📊 Your Excel File is Ready!</h2>
        </div>
        
        <div class="content">
            <p>Hello,</p>
            
            <p>Your Excel file has been successfully generated from your JSON data and is attached to this email.</p>
            
            <div class="highlight">
                <strong>📋 File Details:</strong>
                <ul>
                    <li><strong>Filename:</strong> ${fileName}</li>
                    <li><strong>Generated:</strong> ${new Date().toLocaleString()}</li>
                    <li><strong>Format:</strong> Microsoft Excel (.xlsx)</li>
                </ul>
            </div>
            
            <p><strong>✨ Your file includes:</strong></p>
            <ul>
                <li>Your mapped data in the specified cells and sheets</li>
                <li>Custom styling and formatting as requested</li>
                <li>Any tables, formulas, or advanced features you configured</li>
                <li>Professional formatting optimized for readability</li>
            </ul>
            
            <p>If you need any modifications or have questions about the generated file, please don't hesitate to reach out.</p>
            
            <p>Thank you for using our Excel Generation Service!</p>
        </div>
        
        <div class="footer">
            <p><strong>Excel Generation API Service</strong><br>
            Automated Excel file generation from JSON data<br>
            <em>This is an automated message. Please do not reply to this email.</em></p>
        </div>
    </div>
</body>
</html>
    `;

    const mailOptions = {
      from: `"Excel Generator" <${process.env.EMAIL_USER}>`,
      to: recipient,
      subject: subject,
      text: textMessage,
      html: htmlMessage,
      attachments: [
        {
          filename: fileName,
          content: attachmentBuffer,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      ]
    };

    const info = await transporter.sendMail(mailOptions);
    console.log('Email sent successfully:', info.messageId);
    
  } catch (error) {
    console.error('Error sending email:', error);
    throw new Error(`Failed to send email: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
};

// Send notification email (without attachment)
export const sendNotificationEmail = async (
  recipient: string,
  subject: string,
  message: string,
  isHtml: boolean = false
): Promise<void> => {
  try {
    const transporter = createEmailTransporter();

    const mailOptions = {
      from: `"Excel Generator" <${process.env.EMAIL_USER}>`,
      to: recipient,
      subject: subject,
      ...(isHtml ? { html: message } : { text: message })
    };

    const info = await transporter.sendMail(mailOptions);
    console.log('Notification email sent:', info.messageId);
    
  } catch (error) {
    console.error('Error sending notification email:', error);
    throw new Error(`Failed to send notification: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
};

// Test email configuration
export const testEmailConfiguration = async (): Promise<{
  success: boolean;
  message: string;
}> => {
  try {
    const transporter = createEmailTransporter();
    await transporter.verify();
    
    return {
      success: true,
      message: 'Email configuration is valid and ready to use'
    };
  } catch (error) {
    return {
      success: false,
      message: `Email configuration error: ${error instanceof Error ? error.message : 'Unknown error'}`
    };
  }
};

// Send bulk emails (for multiple recipients)
export const sendBulkEmails = async (
  recipients: string[],
  attachmentBuffer: Buffer,
  fileName: string = 'generated.xlsx',
  customSubject?: string,
  customMessage?: string
): Promise<{
  successful: string[];
  failed: Array<{ email: string; error: string }>;
}> => {
  const successful: string[] = [];
  const failed: Array<{ email: string; error: string }> = [];

  for (const recipient of recipients) {
    try {
      await sendEmailWithAttachment(recipient, attachmentBuffer, fileName, customSubject, customMessage);
      successful.push(recipient);
    } catch (error) {
      failed.push({
        email: recipient,
        error: error instanceof Error ? error.message : 'Unknown error'
      });
    }
  }

  return { successful, failed };
};

// Email templates
export const emailTemplates = {
  success: {
    subject: '✅ Excel File Generated Successfully',
    getHtml: (fileName: string) => `
      <h2>Success! Your Excel file is ready</h2>
      <p>Your file <strong>${fileName}</strong> has been generated and is attached to this email.</p>
      <p>Thank you for using our service!</p>
    `
  },
  
  error: {
    subject: '❌ Excel Generation Failed',
    getHtml: (error: string) => `
      <h2>Generation Failed</h2>
      <p>We encountered an error while generating your Excel file:</p>
      <p style="color: red; font-family: monospace;">${error}</p>
      <p>Please check your request and try again.</p>
    `
  },
  
  templateValidation: {
    subject: '📋 Template Validation Results',
    getHtml: (isValid: boolean, details: string) => `
      <h2>Template Validation ${isValid ? 'Successful' : 'Failed'}</h2>
      <p>${details}</p>
    `
  }
};

export default {
  sendEmailWithAttachment,
  sendNotificationEmail,
  testEmailConfiguration,
  sendBulkEmails,
  emailTemplates
};
