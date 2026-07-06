const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const root = __dirname;
const port = Number(process.env.PORT || 8787);
const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.xls': 'application/vnd.ms-excel',
  '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.zip': 'application/zip'
};

function send(res, code, body, type = 'text/plain; charset=utf-8') {
  res.writeHead(code, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(body);
}

function safeName(name) {
  return path.basename(String(name || 'upload.xls')).replace(/[\\/:*?"<>|]/g, '_');
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function parseMultipart(buffer, contentType) {
  const match = /boundary=(?:(?:"([^"]+)")|([^;]+))/i.exec(contentType || '');
  if (!match) throw new Error('Missing multipart boundary');
  const boundary = Buffer.from('--' + (match[1] || match[2]));
  const parts = [];
  let start = buffer.indexOf(boundary);
  while (start >= 0) {
    start += boundary.length;
    if (buffer[start] === 45 && buffer[start + 1] === 45) break;
    if (buffer[start] === 13 && buffer[start + 1] === 10) start += 2;
    const headerEnd = buffer.indexOf(Buffer.from('\r\n\r\n'), start);
    if (headerEnd < 0) break;
    const headers = buffer.slice(start, headerEnd).toString('utf8');
    let next = buffer.indexOf(boundary, headerEnd + 4);
    if (next < 0) break;
    let dataEnd = next;
    if (buffer[dataEnd - 2] === 13 && buffer[dataEnd - 1] === 10) dataEnd -= 2;
    const disposition = /content-disposition:[^\r\n]*/i.exec(headers);
    if (disposition) {
      const line = disposition[0];
      const nameMatch = /name="([^"]+)"/i.exec(line);
      const fileMatch = /filename="([^"]*)"/i.exec(line);
      const fileStarMatch = /filename\*=([^;]+)/i.exec(line);
      let filename = fileMatch ? fileMatch[1] : '';
      if (!filename && fileStarMatch) {
        filename = fileStarMatch[1].replace(/^UTF-8''/i, '').trim();
        try { filename = decodeURIComponent(filename); } catch {}
      }
      if (nameMatch) parts.push({ name: nameMatch[1], filename, data: buffer.slice(headerEnd + 4, dataEnd), headers });
    }
    start = next;
  }
  return parts;
}

async function handleHtriConvert(req, res) {
  try {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'datasheet-pilot-'));
    const uploadDir = path.join(tempRoot, 'uploads');
    const outputDir = path.join(tempRoot, 'out');
    fs.mkdirSync(uploadDir, { recursive: true });
    const parts = parseMultipart(await readBody(req), req.headers['content-type']);
    const manifest = { files: [] };
    const overrides = {};
    for (const part of parts) {
      if (!part.filename && part.name && part.name.startsWith('override__')) {
        const indexText = part.name.split('__')[1];
        try { overrides[indexText] = JSON.parse(part.data.toString('utf8') || '{}'); } catch { overrides[indexText] = {}; }
        continue;
      }
      if (!part.filename) continue;
      const [prefix, indexText, kindText] = part.name.split('__');
      if (prefix !== 'htri') continue;
      const filename = safeName(part.filename);
      const filePath = path.join(uploadDir, `${indexText}_${filename}`);
      fs.writeFileSync(filePath, part.data);
      manifest.files.push({ name: filename, path: filePath, kind: kindText === 'GPHE' ? 'GPHE' : 'BPHE', overrides: overrides[indexText] || {} });
    }
    if (!manifest.files.length) throw new Error('No HTRI files were uploaded');
    const manifestPath = path.join(tempRoot, 'manifest.json');
    fs.writeFileSync(manifestPath, JSON.stringify(manifest), 'utf8');
    const psPath = path.join(root, 'htri_excel_copy.ps1');
    const templateDir = fs.existsSync(path.join(root, 'templates')) ? path.join(root, 'templates') : path.join(root, 'template');
    const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', psPath, '-ManifestPath', manifestPath, '-OutputDir', outputDir, '-TemplateDir', templateDir];
    const child = spawn('powershell.exe', args, { windowsHide: true });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', data => stdout += data.toString());
    child.stderr.on('data', data => stderr += data.toString());
    child.on('close', code => {
      try {
        if (code !== 0) throw new Error(stderr || stdout || `PowerShell exited ${code}`);
        const line = stdout.trim().split(/\r?\n/).filter(Boolean).pop();
        const result = JSON.parse(line);
        const outPath = result.path;
        const filename = path.basename(outPath);
        res.writeHead(200, {
          'Content-Type': result.type === 'zip' ? 'application/zip' : 'application/vnd.ms-excel.sheet.macroEnabled.12',
          'Content-Disposition': `attachment; filename="${encodeURIComponent(filename)}"`,
          'Cache-Control': 'no-store'
        });
        fs.createReadStream(outPath).pipe(res);
      } catch (error) {
        send(res, 500, String(error.message || error));
      }
    });
  } catch (error) {
    send(res, 500, String(error.message || error));
  }
}

http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api/htri-convert') {
    handleHtriConvert(req, res);
    return;
  }
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(req.url, `http://localhost:${port}`).pathname);
  } catch (error) {
    send(res, 400, 'Bad request');
    return;
  }
  if (pathname === '/') pathname = '/DataSheet Pilot.html';
  const filePath = path.normalize(path.join(root, pathname));
  if (!filePath.startsWith(root)) {
    send(res, 403, 'Forbidden');
    return;
  }
  fs.readFile(filePath, (error, data) => {
    if (error) {
      send(res, 404, 'not found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': mime[path.extname(filePath).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': 'no-store'
    });
    res.end(data);
  });
}).listen(port, '127.0.0.1', () => {
  console.log(`DataSheet Pilot running at http://127.0.0.1:${port}/`);
});
