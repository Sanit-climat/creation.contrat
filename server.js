const express = require('express');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

app.post('/generate', (req, res) => {
  const data = req.body;
  const tmpDir = os.tmpdir();
  const dataFile = path.join(tmpDir, `sc_data_${Date.now()}.json`);
  const outFile = path.join(tmpDir, `sc_contract_${Date.now()}.docx`);

  try {
    fs.writeFileSync(dataFile, JSON.stringify(data, null, 2));
    const result = execSync(`node ${path.join(__dirname, 'generate_contract.js')} "${dataFile}" "${outFile}"`, {
      encoding: 'utf8',
      timeout: 30000
    });

    if (!fs.existsSync(outFile)) {
      throw new Error('Fichier non généré: ' + result);
    }

    const buffer = fs.readFileSync(outFile);
    const clientNom = (data.clientNom || 'CONTRAT').replace(/[^a-zA-Z0-9_-]/g, '_').toUpperCase();
    const date = (data.dateDebut || '').replace(/\//g, '-');
    const filename = `CONTRAT_${clientNom}_${date}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

    // Cleanup
    try { fs.unlinkSync(dataFile); } catch(e) {}
    try { fs.unlinkSync(outFile); } catch(e) {}

  } catch (e) {
    console.error('Generation error:', e.message);
    res.status(500).send('Erreur de génération: ' + e.message);
    try { fs.unlinkSync(dataFile); } catch(err) {}
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Serveur démarré sur http://localhost:${PORT}`);
});
