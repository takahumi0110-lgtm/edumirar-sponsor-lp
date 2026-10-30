/**
 * ローカル確認用の簡易静的サーバー(依存パッケージ不要・Node.js標準機能のみ)。
 * data/sponsors.xlsx をfetchで読み込む都合上、index.htmlをダブルクリックで
 * 直接開く(file://)と読み込めないブラウザがあるため、このサーバーを使う。
 *
 * 使い方: 「ローカル確認用サーバーを起動.bat」をダブルクリック
 *        → ブラウザで http://localhost:8080 を開く
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const PORT = 8080;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

const server = http.createServer(function (req, res) {
  var urlPath = decodeURIComponent(req.url.split('?')[0]);
  if (urlPath === '/') { urlPath = '/index.html'; }

  var filePath = path.join(ROOT, urlPath);
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(filePath, function (err, data) {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found: ' + urlPath);
      return;
    }
    var ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, function () {
  console.log('ローカル確認用サーバーを起動しました: http://localhost:' + PORT);
  console.log('終了するにはこのウィンドウを閉じるか Ctrl+C を押してください。');
});
