import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType } from 'docx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Read the markdown file
const mdPath = path.join(__dirname, '../docs/MANUALE_UTENTE.md');
const mdContent = fs.readFileSync(mdPath, 'utf-8');

// Parse markdown and create docx elements
const children = [];
const lines = mdContent.split('\n');

let inTable = false;
let tableRows = [];
let currentTableAlignments = [];

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  
  // Skip empty lines
  if (line.trim() === '') {
    if (inTable && tableRows.length > 0) {
      // End table
      children.push(new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      }));
      tableRows = [];
      inTable = false;
    }
    continue;
  }
  
  // Headers
  if (line.startsWith('# ')) {
    children.push(new Paragraph({
      text: line.substring(2).trim(),
      heading: HeadingLevel.TITLE,
      spacing: { after: 400 },
    }));
    continue;
  }
  
  if (line.startsWith('## ')) {
    children.push(new Paragraph({
      text: line.substring(3).trim(),
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }));
    continue;
  }
  
  if (line.startsWith('### ')) {
    children.push(new Paragraph({
      text: line.substring(4).trim(),
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 300, after: 150 },
    }));
    continue;
  }
  
  if (line.startsWith('#### ')) {
    children.push(new Paragraph({
      text: line.substring(5).trim(),
      heading: HeadingLevel.HEADING_3,
      spacing: { before: 200, after: 100 },
    }));
    continue;
  }
  
  // Horizontal rule
  if (line.match(/^---+$/)) {
    children.push(new Paragraph({
      text: '',
      spacing: { before: 200, after: 200 },
      border: {
        bottom: { color: "auto", space: 1, style: BorderStyle.SINGLE, size: 6 },
      },
    }));
    continue;
  }
  
  // Table detection
  if (line.startsWith('|')) {
    inTable = true;
    const cells = line.split('|').filter(c => c.trim() !== '').map(c => c.trim());
    
    // Check if it's a separator line
    if (cells.every(c => c.match(/^-+$/))) {
      currentTableAlignments = cells.map(c => c.includes(':') ? AlignmentType.CENTER : AlignmentType.LEFT);
      continue;
    }
    
    const tableRow = new TableRow({
      children: cells.map((cellText, idx) => {
        // Clean up markdown formatting
        let cleanText = cellText
          .replace(/\*\*(.+?)\*\*/g, '$1')
          .replace(/\*(.+?)\*/g, '$1')
          .replace(/`(.+?)`/g, '$1');
        
        return new TableCell({
          children: [new Paragraph({ 
            text: cleanText,
            alignment: currentTableAlignments[idx] || AlignmentType.LEFT,
          })],
          width: { size: 100 / cells.length, type: WidthType.PERCENTAGE },
        });
      }),
    });
    
    tableRows.push(tableRow);
    continue;
  }
  
  // End table if we were in one
  if (inTable && tableRows.length > 0) {
    children.push(new Table({
      rows: tableRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
    }));
    tableRows = [];
    inTable = false;
  }
  
  // Blockquotes
  if (line.startsWith('> ')) {
    const text = line.substring(2).trim();
    let cleanText = text
      .replace(/\*\*(.+?)\*\*/g, '$1')
      .replace(/\*(.+?)\*/g, '$1');
    
    children.push(new Paragraph({
      children: [
        new TextRun({
          text: cleanText,
          italics: true,
        }),
      ],
      indent: { left: 720 },
      spacing: { before: 100, after: 100 },
    }));
    continue;
  }
  
  // List items
  if (line.match(/^[-*]\s/)) {
    const text = line.replace(/^[-*]\s/, '').trim();
    let cleanText = text
      .replace(/\*\*(.+?)\*\*/g, '$1')
      .replace(/\*(.+?)\*/g, '$1')
      .replace(/`(.+?)`/g, '$1');
    
    children.push(new Paragraph({
      text: '• ' + cleanText,
      indent: { left: 360 },
      spacing: { before: 50, after: 50 },
    }));
    continue;
  }
  
  // Numbered list
  if (line.match(/^\d+\.\s/)) {
    const text = line.replace(/^\d+\.\s/, '').trim();
    let cleanText = text
      .replace(/\*\*(.+?)\*\*/g, '$1')
      .replace(/\*(.+?)\*/g, '$1')
      .replace(/`(.+?)`/g, '$1');
    
    children.push(new Paragraph({
      text: line.match(/^\d+/)[0] + '. ' + cleanText,
      indent: { left: 360 },
      spacing: { before: 50, after: 50 },
    }));
    continue;
  }
  
  // Regular paragraph
  let cleanText = line
    .replace(/\*\*(.+?)\*\*/g, '$1')
    .replace(/\*(.+?)\*/g, '$1')
    .replace(/`(.+?)`/g, '$1');
  
  children.push(new Paragraph({
    text: cleanText,
    spacing: { before: 100, after: 100 },
  }));
}

// Create document
const doc = new Document({
  title: 'Manuale Utente - App Totem ALFA ENG',
  description: 'Guida all\'utilizzo dell\'applicazione Totem Magazzino',
  creator: 'ALFA ENG',
  sections: [{
    properties: {},
    children: children,
  }],
});

// Save document
const outputPath = path.join(__dirname, '../docs/MANUALE_UTENTE.docx');
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log('Document created successfully:', outputPath);
}).catch((err) => {
  console.error('Error creating document:', err);
});
