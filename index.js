const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const jsonData = require('./rule.json');

// const jsonData = [
//   {
//     name: 'Font declarations should contain at least one generic font family',
//     data: {
//       issues: [
//         { project: 'Project A', line: 123, severity: 'MAJOR' },
//         { project: 'Project B', line: 456, severity: 'MINOR' },
//       ],
//     },
//   },
//   {
//     name: 'Unnecessary character escapes should be removed',
//     data: {
//       issues: [{ project: 'Project C', line: 789, severity: 'CRITICAL' }],
//     },
//   },
// ];

// Transform the data into the format expected by the template
const transformedData = {
  rule: jsonData.map((rule) => ({
    name: rule.name,
    issues: rule.data.issues.map((issue) => ({
      project: issue.project,
      line: issue.line,
      severity: issue.severity,
    })),
  })),
};

console.log('transform', transformedData);

// Load the Word template
const templatePath = 'template.docx';
const templateContent = fs.readFileSync(templatePath, 'binary');
const zip = new PizZip(templateContent);
const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

try {
  // Set data for the template
  doc.setData(transformedData);
  // Render the document with the provided data
  doc.render();

  // Save the generated document
  const outputPath = 'output.docx';
  const buffer = doc.getZip().generate({ type: 'nodebuffer' });
  fs.writeFileSync(outputPath, buffer);

  console.log(`Document generated successfully at ${outputPath}`);
} catch (error) {
  const errorDetails = error.getErrors ? error.getErrors() : [error];
  console.error(
    'Rendering Error Details:',
    JSON.stringify(errorDetails, null, 2)
  );
  throw error;
}
