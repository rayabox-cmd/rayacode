const express = require('express');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require('docx');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(express.static(path.join(__dirname, 'public')));

app.post('/convert', upload.array('pdfs', 20), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'Aucun fichier PDF reçu.' });
    }

    const sections = [];

    for (const file of req.files) {
      let text = '';
      try {
        const data = await pdfParse(file.buffer);
        text = data.text || '';
      } catch {
        text = `[Erreur de lecture du fichier ${file.originalname}]`;
      }

      // Title for each PDF
      sections.push(
        new Paragraph({
          text: file.originalname.replace(/\.pdf$/i, ''),
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 400, after: 200 },
        })
      );

      // Split text into paragraphs
      const lines = text.split(/\n+/).map(l => l.trim()).filter(l => l.length > 0);
      for (const line of lines) {
        sections.push(
          new Paragraph({
            children: [new TextRun({ text: line, font: 'Calibri', size: 22 })],
            spacing: { after: 100 },
          })
        );
      }

      // Page break between documents
      sections.push(new Paragraph({ children: [], pageBreakBefore: true }));
    }

    // Remove trailing page break
    if (sections.length > 0) sections.pop();

    const doc = new Document({
      sections: [{ properties: {}, children: sections }],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="documents_medicaux.docx"');
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Erreur lors de la conversion.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Serveur démarré sur http://localhost:${PORT}`));
