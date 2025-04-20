const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const fs = require('fs');

async function generateSampleFiles() {
    // Generate sample Excel data
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Recipients');
    
    // Add headers
    worksheet.columns = [
        { header: 'to', key: 'to' },
        { header: 'email', key: 'email' },
        { header: 'title', key: 'title' },
        { header: 'company', key: 'company' },
        { header: 'position', key: 'position' },
        { header: 'date', key: 'date' }
    ];

    // Add sample data
    const sampleData = [
        {
            to: 'John Doe',
            email: 'john.doe@example.com',
            title: 'Mr.',
            company: 'Tech Solutions Inc.',
            position: 'Software Engineer',
            date: '20 April, 2025'
        },
        {
            to: 'Jane Smith',
            email: 'jane.smith@example.com',
            title: 'Ms.',
            company: 'Digital Innovations Ltd.',
            position: 'Project Manager',
            date: '20 April, 2025'
        }
    ];

    // Add rows to worksheet
    sampleData.forEach(data => {
        worksheet.addRow(data);
    });

    // Style the headers
    worksheet.getRow(1).font = { bold: true };
    worksheet.columns.forEach(column => {
        column.width = 20;
    });

    // Save the Excel file
    await workbook.xlsx.writeFile('test_data.xlsx');
    console.log('Sample Excel file created: test_data.xlsx');

    // Create sample DOCX template
    const template = `
{title} {to}
{position}
{company}
{date}

Dear {title} {to},

I hope this letter finds you well. I am writing to inform you about our upcoming technology conference that will be held next month.

As a respected {position} at {company}, we believe your expertise and insights would be invaluable to our event. We would be honored to have you join us as a guest speaker.

The conference will focus on emerging technologies and their impact on business operations. Your experience in implementing innovative solutions would provide our attendees with valuable real-world perspectives.

Please let us know if you would be interested in participating. We can schedule a call to discuss the details further.

Best regards,
Conference Organizing Committee`;

    // Create a zip of a minimal docx file
    const zip = new PizZip();
    
    // Add required files for a basic DOCX
    zip.file('word/document.xml', 
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>${template}</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>`);

    zip.file('word/_rels/document.xml.rels',
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

    zip.file('[Content_Types].xml',
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

    zip.file('_rels/.rels',
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

    // Generate the DOCX file
    const docx = zip.generate({ type: 'nodebuffer' });
    fs.writeFileSync('test_template.docx', docx);
    console.log('Sample DOCX template created: test_template.docx');
}

generateSampleFiles().catch(console.error);