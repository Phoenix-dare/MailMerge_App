const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const cors = require('cors');
const { PDFDocument, rgb } = require('pdf-lib');
require('dotenv').config();

// Ensure necessary folders exist with absolute paths
const uploadsDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'output');

[uploadsDir, outputDir].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

const app = express();
app.use(express.json());
const port = process.env.PORT || 5000;

// Configure CORS
const corsOptions = {
  origin: process.env.CORS_ORIGIN || 'http://localhost:5173',
  optionsSuccessStatus: 200
};
app.use(cors(corsOptions));

// Configure multer with absolute paths
const upload = multer({ 
  dest: uploadsDir,
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
}).fields([
  { name: 'template', maxCount: 1 },
  { name: 'datafile', maxCount: 1 }
]);

// Email configuration with enhanced debugging
const transporter = nodemailer.createTransport({
  service: 'gmail',
  host: 'smtp.gmail.com',
  port: 465,
  secure: true,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
  },
  debug: true,
  logger: true
});

// Verify email configuration on startup with detailed logging
transporter.verify((error, success) => {
  if (error) {
    console.error('Email configuration error:', error);
    console.error('Email settings:', {
      user: process.env.EMAIL_USER,
      host: 'smtp.gmail.com',
      port: 465
    });
  } else {
    console.log('Email server is ready to send messages');
    // Test email connection
    transporter.sendMail({
      from: process.env.EMAIL_USER,
      to: process.env.EMAIL_USER,
      subject: 'Mail Merge System Test',
      text: 'This is a test email to verify the email system is working.'
    }).then(info => {
      console.log('Test email sent successfully:', info.messageId);
    }).catch(err => {
      console.error('Test email failed:', err);
    });
  }
});

function validateTemplate(templatePath) {
  try {
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
    new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    return true;
  } catch (error) {
    console.error('Invalid template file:', error.message);
    if (error.properties && error.properties.errors) {
      error.properties.errors.forEach((err) => {
        console.error('Template error detail:', err);
      });
    }
    return false;
  }
}

async function parseExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];

  const records = [];
  const headers = worksheet.getRow(1).values.slice(1);

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header
    const values = row.values.slice(1);
    const record = {};
    headers.forEach((header, index) => {
      record[header] = values[index];
    });
    records.push(record);
  });

  return records;
}

// Function to validate records
function validateRecords(records) {
  if (!Array.isArray(records) || records.length === 0) {
    console.error('No valid records found in the data file.');
    return false;
  }
  const requiredFields = ['to', 'email']; // Add other required fields if necessary
  for (const record of records) {
    for (const field of requiredFields) {
      if (!record[field]) {
        console.error(`Missing required field "${field}" in record:`, record);
        return false;
      }
    }
  }
  return true;
}

// Function to ensure output directory exists
function ensureOutputDirectory() {
  const outputDir = path.join(__dirname, 'output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
}

// New email endpoint
app.post('/send-email', async (req, res) => {
  const { to, files } = req.body;
  
  if (!to || !files) {
    return res.status(400).json({ error: 'Missing required parameters' });
  }

  try {
    // Read the files
    const docxBuffer = fs.readFileSync(path.join(__dirname, 'output', files.docx));
    const pdfBuffer = fs.readFileSync(path.join(__dirname, 'output', files.pdf));

    // Send email
    const mailOptions = {
      from: {
        name: 'Mail Merge System',
        address: process.env.EMAIL_USER
      },
      to: to,
      subject: 'Your Generated Letter',
      text: 'Please find your generated letter attached in both DOCX and PDF formats.',
      attachments: [
        {
          filename: files.docx,
          content: docxBuffer,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        },
        {
          filename: files.pdf,
          content: pdfBuffer,
          contentType: 'application/pdf'
        }
      ]
    };

    const info = await transporter.sendMail(mailOptions);
    res.json({ 
      message: 'Email sent successfully',
      messageId: info.messageId 
    });
  } catch (error) {
    console.error('Error sending email:', error);
    res.status(500).json({ 
      error: 'Failed to send email',
      details: error.message 
    });
  }
});

// Modified upload endpoint
app.post('/upload', (req, res) => {
  upload(req, res, async function(err) {
    if (err) {
      console.error('Upload error:', err);
      return res.status(400).json({ error: 'File upload error: ' + err.message });
    }

    try {
      if (!req.files || !req.files.template || !req.files.datafile) {
        return res.status(400).json({ error: 'Both template and data files are required' });
      }

      const templatePath = req.files.template[0].path;
      const dataPath = req.files.datafile[0].path;
      const records = await parseExcel(dataPath);
      const results = [];
      
      for (const record of records) {
        try {
          // Generate DOCX
          const content = fs.readFileSync(templatePath, 'binary');
          const zip = new PizZip(content);
          const doc = new Docxtemplater(zip, { 
            paragraphLoop: true, 
            linebreaks: true 
          });

          await doc.resolveData(record);
          doc.render();
          
          const docxBuffer = doc.getZip().generate({ type: 'nodebuffer' });
          const docxPath = path.join('output', `${record.to}_letter.docx`);
          const pdfPath = path.join('output', `${record.to}_letter.pdf`);

          // Save DOCX file
          fs.writeFileSync(docxPath, docxBuffer);
          console.log(`DOCX file saved: ${docxPath}`);

          // Create PDF with proper letter formatting
          const pdfDoc = await PDFDocument.create();
          const page = pdfDoc.addPage([595.276, 841.890]); // A4 size
          const { width, height } = page.getSize();
          
          // Header section (sender info)
          let currentY = height - 50;
          const leftMargin = 50;
          
          page.drawText(`${record.title} ${record.to}`, {
            x: leftMargin,
            y: currentY,
            size: 12,
            color: rgb(0, 0, 0),
          });
          
          currentY -= 20;
          page.drawText(record.position, {
            x: leftMargin,
            y: currentY,
            size: 12,
            color: rgb(0, 0, 0),
          });
          
          currentY -= 20;
          page.drawText(record.company, {
            x: leftMargin,
            y: currentY,
            size: 12,
            color: rgb(0, 0, 0),
          });
          
          currentY -= 40;
          page.drawText(record.date, {
            x: leftMargin,
            y: currentY,
            size: 12,
            color: rgb(0, 0, 0),
          });
          
          // Salutation
          currentY -= 40;
          page.drawText(`Dear ${record.title} ${record.to},`, {
            x: leftMargin,
            y: currentY,
            size: 12,
            color: rgb(0, 0, 0),
          });
          
          // Letter body
          currentY -= 30;
          const bodyText = `I hope this letter finds you well. I am writing to inform you about our upcoming technology conference that will be held next month.

As a respected ${record.position} at ${record.company}, we believe your expertise and insights would be invaluable to our event. We would be honored to have you join us as a guest speaker.

The conference will focus on emerging technologies and their impact on business operations. Your experience in implementing innovative solutions would provide our attendees with valuable real-world perspectives.

Please let us know if you would be interested in participating. We can schedule a call to discuss the details further.`;

          // Split body text into lines for proper formatting
          const words = bodyText.split(' ');
          let line = '';
          const lineHeight = 20;
          const maxWidth = width - (leftMargin * 2);
          
          for (const word of words) {
            const testLine = line + (line ? ' ' : '') + word;
            if (testLine.endsWith('\n\n')) {
              page.drawText(line, {
                x: leftMargin,
                y: currentY,
                size: 12,
                color: rgb(0, 0, 0),
              });
              currentY -= lineHeight * 2; // Double space for paragraphs
              line = '';
            } else if (testLine.length * 7 > maxWidth) { // Approximate character width
              page.drawText(line, {
                x: leftMargin,
                y: currentY,
                size: 12,
                color: rgb(0, 0, 0),
              });
              currentY -= lineHeight;
              line = word;
            } else {
              line = testLine;
            }
          }
          
          // Draw remaining text if any
          if (line) {
            page.drawText(line, {
              x: leftMargin,
              y: currentY,
              size: 12,
              color: rgb(0, 0, 0),
            });
          }
          
          // Closing
          currentY -= lineHeight * 3;
          page.drawText('Best regards,', {
            x: leftMargin,
            y: currentY,
            size: 12,
            color: rgb(0, 0, 0),
          });
          
          currentY -= lineHeight;
          page.drawText('Conference Organizing Committee', {
            x: leftMargin,
            y: currentY,
            size: 12,
            color: rgb(0, 0, 0),
          });

          const pdfBuffer = await pdfDoc.save();
          fs.writeFileSync(pdfPath, pdfBuffer);
          console.log(`PDF file saved: ${pdfPath}`);

          // Send email with better error handling
          try {
            console.log(`Attempting to send email to: ${record.email}`);
            const mailOptions = {
              from: {
                name: 'Mail Merge System',
                address: process.env.EMAIL_USER
              },
              to: record.email,
              subject: 'Your Generated Letter',
              text: 'Please find your generated letter attached in both DOCX and PDF formats.',
              attachments: [
                {
                  filename: `${record.to}_letter.docx`,
                  content: docxBuffer,
                  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                },
                {
                  filename: `${record.to}_letter.pdf`,
                  content: pdfBuffer,
                  contentType: 'application/pdf'
                }
              ]
            };

            const info = await transporter.sendMail(mailOptions);
            console.log('Email sent successfully:', info.messageId);
            
            results.push({ 
              to: record.email, 
              status: 'success',
              messageId: info.messageId,
              files: {
                docx: `${record.to}_letter.docx`,
                pdf: `${record.to}_letter.pdf`
              }
            });
          } catch (emailError) {
            console.error('Error sending email:', emailError);
            throw new Error(`Failed to send email: ${emailError.message}`);
          }
        } catch (error) {
          console.error('Processing error for record:', record, error);
          results.push({ 
            to: record.email, 
            status: 'error', 
            error: error.message 
          });
        }
      }

      res.json({ 
        message: 'Processing complete', 
        results 
      });

    } catch (error) {
      console.error('Server error:', error);
      res.status(500).json({ 
        error: 'Server error', 
        message: error.message 
      });
    } finally {
      // Cleanup uploaded files
      if (req.files) {
        Object.values(req.files).forEach(files => {
          files.forEach(file => {
            fs.unlink(file.path, () => {});
          });
        });
      }
    }
  });
});

// Add routes to serve generated files
app.get('/files/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'output', filename);
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found' });
  }

  res.download(filePath);
});

app.get('/preview/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'output', filename);
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found' });
  }

  // For PDF files, stream for preview
  if (filename.endsWith('.pdf')) {
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'inline');
    fs.createReadStream(filePath).pipe(res);
  } else {
    // For other files, trigger download
    res.download(filePath);
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
