'use strict';
/**
 * SCF 静态网站托管函数（国内可访问镜像）
 * ------------------------------------------------------------------
 * 部署包结构：index.js（本文件）+ site/（网页静态文件）。
 * 函数 URL 触发：按请求路径返回 site/ 下对应文件，'/' → index.html。
 * 无任何密钥；按量计费免费额度（每月 100 万次调用）足够，无到期/续期。
 * ------------------------------------------------------------------
 */
const fs = require('fs');
const path = require('path');

const SITE_DIR = path.join(__dirname, 'site');
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.webp': 'image/webp',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

function resolvePath(event) {
  var p = (event && event.path) ||
    (event && event.requestContext && event.requestContext.http && event.requestContext.http.path) ||
    '/';
  try { p = decodeURIComponent(p); } catch (e) {}
  p = p.split('?')[0].split('#')[0];
  if (p === '' || p === '/') return '/index.html';
  if (p.endsWith('/')) p += 'index.html';
  return p;
}

function reply(statusCode, body, contentType, isB64) {
  return {
    isBase64Encoded: !!isB64,
    statusCode: statusCode,
    headers: { 'Content-Type': contentType },
    body: body
  };
}

exports.main_handler = async (event) => {
  const p = resolvePath(event);
  const filePath = path.normalize(path.join(SITE_DIR, p));
  // 防目录穿越
  if (filePath !== SITE_DIR && !filePath.startsWith(SITE_DIR + path.sep)) {
    return reply(403, 'forbidden', 'text/plain; charset=utf-8');
  }
  let data;
  try {
    data = fs.readFileSync(filePath);
  } catch (e) {
    return reply(404, 'not found: ' + p, 'text/plain; charset=utf-8');
  }
  const ext = path.extname(filePath).toLowerCase();
  const mime = MIME[ext] || 'application/octet-stream';
  const isText = /^text\//.test(mime) || mime.indexOf('json') >= 0 || mime.indexOf('javascript') >= 0;
  const res = reply(
    200,
    isText ? data.toString('utf8') : data.toString('base64'),
    mime,
    !isText
  );
  // html/json 不缓存（ fares.json 每日更新），其余静态资源缓存 5 分钟
  res.headers['Cache-Control'] = (ext === '.html' || ext === '.json') ? 'no-cache' : 'public, max-age=300';
  return res;
};
