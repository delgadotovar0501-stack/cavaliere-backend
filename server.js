const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun, AlignmentType, UnderlineType } = require('docx');

const app = express();
app.use(cors());
app.use(express.json());

function buildProposalDoc({ clientName, clientPhone, clientEmail, estimateRef, jobAddress, date, lineItems, total, optional, note }) {

  const t = (text, opts = {}) => new TextRun({ text, size: 22, font: 'Times New Roman', ...opts });
  const tb = (text, opts = {}) => t(text, { bold: true, ...opts });
  const u = { underline: { type: UnderlineType.SINGLE } };

  const p = (children, opts = {}) => new Paragraph({
    children: Array.isArray(children) ? children : [children],
    spacing: { after: 0 },
    ...opts
  });

  const children = [
    // DATE — right aligned
    p([tb(`DATE:   ${date}`)], { alignment: AlignmentType.RIGHT, spacing: { after: 200 } }),

    // TO:
    p([tb('TO:'), t('    '), t(clientName, u)], { spacing: { after: 0 } }),
    ...(clientPhone ? [p([t(clientPhone)], { indent: { left: 864 }, spacing: { after: 0 } })] : []),
    ...(clientEmail ? [p([t(clientEmail)], { indent: { left: 864 }, spacing: { after: 0 } })] : []),
    ...(estimateRef ? [p([t(estimateRef)], { indent: { left: 864 }, spacing: { after: 240 } })] : [p([t('')], { spacing: { after: 240 } })]),

    // JOB SITE
    p([
      tb('JOB SITE:'),
      t('  '),
      t(jobAddress.split(' ')[0], u),
      t(' ' + jobAddress.split(' ').slice(1).join(' ')),
    ], { spacing: { after: 200 } }),

    // SCOPE OF WORK
    p([tb('SCOPE OF WORK:')], { spacing: { after: 0 } }),
    p([t('Based on site visit, Cavaliere Electric & Sons is pleased to quote the above-mentioned job. We will provide all labor and materials for a complete job. Work is to be done in accordance with all local and national electrical codes and guaranteed for one year.')],
      { indent: { left: 720 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 200 } }),

    // INCLUDING
    p([tb('INCLUDING:')], { spacing: { after: 80 } }),

    // Line items
    ...lineItems.map((item, i) => p(
      [tb(`${i + 1}.  ${item.desc}${item.price ? '   $' + Number(item.price).toLocaleString('en-US', { minimumFractionDigits: 2 }) : ''}`)],
      { indent: { left: 720 }, spacing: { after: 60 } }
    )),

    p([t('')], { spacing: { after: 120 } }),

    // TOTAL
    p([
      tb('TOTAL COST OF JOB:'),
      t('  '),
      tb('$' + Number(total).toLocaleString('en-US', { minimumFractionDigits: 2 }), u),
    ], { spacing: { after: 200 } }),

    // Optional (if provided)
    ...(optional ? [p([
      tb('Optional:'),
      t('  '),
      tb(optional),
    ], { spacing: { after: 200 } })] : []),

    // NOTE
    ...(note ? [p([
      tb('NOTE:', u),
      tb('  ' + note),
    ], { spacing: { after: 0 } })] : []),
  ];

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 720, bottom: 720, left: 1080, right: 1080 }
        }
      },
      children
    }]
  });
}

app.post('/generate', async (req, res) => {
  try {
    const {
      clientName, clientPhone, clientEmail, estimateRef,
      jobAddress, date, lineItems, total, optional, note
    } = req.body;

    if (!clientName || !jobAddress || !lineItems?.length) {
      return res.status(400).json({ error: 'Missing required fields: clientName, jobAddress, lineItems' });
    }

    const doc = buildProposalDoc({
      clientName, clientPhone, clientEmail, estimateRef,
      jobAddress,
      date: date || new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }),
      lineItems,
      total: total || lineItems.reduce((s, i) => s + (parseFloat(i.price) || 0), 0),
      optional,
      note
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = `Proposal_${clientName.replace(/\s+/g, '_')}_${date || 'draft'}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (_, res) => res.json({ status: 'ok' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Cavaliere backend running on port ${PORT}`));
