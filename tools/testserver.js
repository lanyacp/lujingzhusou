/**
 * testserver.js — 临时测试服务器: 静态服务项目目录 + 收集加载项 JS 自测结果
 * 用法: node tools/testserver.js   (端口 3889, 与 publish.xml 的 jspluginonline 一致)
 * 测试结果写入 tools/test-results/latest.json
 */
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const PORT = 3889;
const OUT_DIR = path.join(__dirname, 'test-results');

const MIME = {
    '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
    '.xml': 'text/xml; charset=utf-8', '.svg': 'image/svg+xml', '.json': 'application/json'
};

let resultCount = 0;

const server = http.createServer((req, res) => {
    const url = decodeURIComponent(req.url.split('?')[0]);

    if (req.method === 'POST' && url === '/__test/result') {
        let body = '';
        req.on('data', c => body += c);
        req.on('end', () => {
            try {
                fs.mkdirSync(OUT_DIR, { recursive: true });
                resultCount++;
                const f = path.join(OUT_DIR, 'latest.json');
                fs.writeFileSync(f, body);
                fs.writeFileSync(path.join(OUT_DIR, `run-${resultCount}.json`), body);
                console.log(`[test-result #${resultCount}] received ${(body.length / 1024).toFixed(1)}KB -> ${f}`);
                // 同时打印摘要便于观察
                try {
                    const o = JSON.parse(body);
                    console.log('  hasApp=' + o.hasApp + ' hasABP=' + o.hasABP +
                        (o.forms ? ' forms=' + JSON.stringify(o.forms) : '') +
                        (o.fatal ? ' fatal=' + o.fatal : ''));
                } catch (e) { }
            } catch (e) {
                console.log('save err: ' + e.message);
            }
            res.writeHead(200, { 'Content-Type': 'text/plain', 'Access-Control-Allow-Origin': '*' });
            res.end('ok');
        });
        return;
    }

    // 静态文件
    console.log(`[req] ${req.method} ${url}`);
    let fp = path.join(ROOT, url === '/' ? 'index.html' : url);
    if (fp.indexOf(ROOT) !== 0) { res.writeHead(403); return res.end(); }
    fs.readFile(fp, (err, data) => {
        if (err) { res.writeHead(404); return res.end('not found: ' + url); }
        res.writeHead(200, {
            'Content-Type': MIME[path.extname(fp)] || 'application/octet-stream',
            'Cache-Control': 'no-store, no-cache, must-revalidate',
            'Pragma': 'no-cache'
        });
        res.end(data);
    });
});

server.listen(PORT, '127.0.0.1', () => console.log(`test server on http://127.0.0.1:${PORT}`));
