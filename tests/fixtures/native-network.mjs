// A delayed HTTP proxy for the native integration harness. Never shipped in the Windows app.
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';

const root = process.cwd();
const directory = path.join(root, '.runtime/native-test');
const serviceRoot = path.join(directory, 'root');
const backend = JSON.parse(
  fs.readFileSync('.runtime/endpoint.json', 'utf8').replace(/^\uFEFF/, ''),
);
fs.mkdirSync(path.join(serviceRoot, '.runtime'), { recursive: true });
fs.mkdirSync(path.join(serviceRoot, 'data'), { recursive: true });
fs.copyFileSync('data/financials.json', path.join(serviceRoot, 'data/financials.json'));
const controlPath = path.join(directory, 'network.json');
fs.writeFileSync(controlPath, '{}');
let baseUrl;
function controls() {
  try {
    return JSON.parse(fs.readFileSync(controlPath, 'utf8'));
  } catch {
    return {};
  }
}
function headers(request) {
  const values = { ...request.headers, host: new URL(backend.BaseUrl).host };
  if (values.origin === baseUrl) values.origin = backend.BaseUrl;
  return values;
}
const sockets = new Set();
const server = http.createServer((request, response) => {
  const forward = () => {
    if (response.destroyed) return;
    const upstream = http.request(
      new URL(request.url, backend.BaseUrl),
      { method: request.method, headers: headers(request) },
      (incoming) => {
        response.writeHead(incoming.statusCode, incoming.headers);
        incoming.pipe(response);
      },
    );
    upstream.on('error', () => {
      if (!response.headersSent) response.writeHead(502);
      response.end();
    });
    response.on('close', () => upstream.destroy());
    request.pipe(upstream);
  };
  const delay = request.url.startsWith('/api/facts/') ? (controls().delayMs ?? 0) : 0;
  const timer = setTimeout(forward, delay);
  response.on('close', () => clearTimeout(timer));
});
server.on('connection', (socket) => {
  sockets.add(socket);
  socket.on('close', () => sockets.delete(socket));
});
server.on('upgrade', (request, socket, head) => {
  const upstream = http.request(new URL(request.url, backend.BaseUrl), {
    headers: headers(request),
  });
  upstream.on('upgrade', (response, connection, upstreamHead) => {
    socket.write(
      `HTTP/1.1 ${response.statusCode} ${response.statusMessage}\r\n${Object.entries(
        response.headers,
      )
        .map(([key, value]) => `${key}: ${value}`)
        .join('\r\n')}\r\n\r\n`,
    );
    if (head.length) connection.write(head);
    if (upstreamHead.length) socket.write(upstreamHead);
    connection.pipe(socket).pipe(connection);
    socket.on('close', () => connection.destroy());
    connection.on('error', () => socket.destroy());
  });
  upstream.on('response', (response) => {
    response.resume();
    socket.destroy();
  });
  upstream.on('error', () => socket.destroy());
  socket.on('error', () => upstream.destroy());
  upstream.end();
});
server.listen(0, '127.0.0.1', () => {
  baseUrl = `http://127.0.0.1:${server.address().port}`;
  fs.writeFileSync(
    path.join(serviceRoot, '.runtime/endpoint.json'),
    JSON.stringify({ ...backend, BaseUrl: baseUrl, ProcessId: process.pid }),
  );
  fs.writeFileSync(
    path.join(directory, 'network-ready.json'),
    JSON.stringify({ pid: process.pid }),
  );
});
setInterval(() => {
  if (controls().stop) {
    for (const socket of sockets) socket.destroy();
    server.close(() => process.exit(0));
  }
}, 100).unref();
