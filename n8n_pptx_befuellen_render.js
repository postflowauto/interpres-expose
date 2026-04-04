// PPTX befüllen via Render.com Service
const exposeData = $input.first().json.exposeData;
const templateBase64 = $input.first().binary.template.data;

const projektName = (exposeData.projekt_name || 'Projekt').replace(/[^a-zA-Z0-9]/g, '_');
const filename = `INTERPRÉS_Expose_${projektName}.pptx`;

const response = await fetch('https://interpres-pptx-filler.onrender.com/fill-pptx', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'X-API-Token': 'interpres-secret-2026'
  },
  body: JSON.stringify({
    template_base64: templateBase64,
    expose_data: exposeData,
    filename: filename
  })
});

if (!response.ok) {
  const err = await response.text();
  throw new Error(`Render Service Fehler ${response.status}: ${err}`);
}

const result = await response.json();

if (!result.success) {
  throw new Error('Render Service: ' + result.error);
}

return [{
  json: { filename: result.filename, size_bytes: result.size_bytes },
  binary: {
    pptx: {
      data: result.pptx_base64,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      fileName: result.filename,
      fileExtension: 'pptx'
    }
  }
}];
