// KRONOS offline: a tiny web server for the app folder, using only what Node.js already has.
// Nothing leaves this computer: it listens on localhost only.
import { createServer } from 'node:http'
import { readFile, stat } from 'node:fs/promises'
import { extname, join, normalize, resolve, sep } from 'node:path'
import { fileURLToPath } from 'node:url'
import { exec } from 'node:child_process'

const ROOT = resolve(fileURLToPath(new URL('../app/', import.meta.url)))
const PORT = 5181
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.ico': 'image/x-icon',
  '.json': 'application/json', '.wasm': 'application/wasm', '.woff': 'font/woff', '.woff2': 'font/woff2', '.map': 'application/json',
}

const server = createServer(async (req, res) => {
  try {
    let path = normalize(decodeURIComponent(new URL(req.url, 'http://x').pathname))
    let file = resolve(join(ROOT, path))
    if (file !== ROOT && !file.startsWith(ROOT + sep)) throw new Error('outside')
    if ((await stat(file).catch(() => null))?.isDirectory()) file = join(file, 'index.html')
    const body = await readFile(file)
    res.writeHead(200, { 'Content-Type': TYPES[extname(file).toLowerCase()] ?? 'application/octet-stream', 'Cache-Control': 'no-cache' })
    res.end(body)
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain' })
    res.end('Not found')
  }
})

server.on('error', (error) => {
  if (error.code === 'EADDRINUSE') console.log(`\nPort ${PORT} is busy. KRONOS may already be running: open http://localhost:${PORT}/ in your browser.\n`)
  else console.log(error.message)
  process.exit(1)
})

server.listen(PORT, '127.0.0.1', () => {
  const url = `http://localhost:${PORT}/`
  console.log(`\nKRONOS Offline, version dev\nKRONOS is running at ${url}\nKeep this window open while you use it. Close it (or press Ctrl+C) to stop.\n`)
  const open = process.platform === 'win32' ? `start "" "${url}"` : process.platform === 'darwin' ? `open "${url}"` : `xdg-open "${url}"`
  exec(open, () => {})
})
