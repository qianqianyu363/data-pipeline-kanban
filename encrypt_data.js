#!/usr/bin/env node
/**
 * 看板数据加密脚本
 *   data.json(明文) → data.enc.json(密文信封) + 同步写入 index.html 内嵌块
 *
 * 算法：PBKDF2-SHA256 派生密钥 + AES-256-GCM
 *   —— 与浏览器 WebCrypto SubtleCrypto 完全兼容，前端零依赖解密
 *
 * 用法：
 *   node encrypt_data.js --pw "你的口令" --save-key   # 首次：设置口令并存本地密钥文件
 *   node encrypt_data.js                              # 之后：自动读取本地密钥
 *   node encrypt_data.js --verify                     # 校验：解密并比对原文
 *
 * 密钥来源优先级：--pw 参数 > 环境变量 KANBAN_PW > .kanban_key 文件
 * 注意：.kanban_key 已在 .gitignore 中，切勿提交
 */
const fs = require('fs');
const path = require('path');
const { webcrypto } = require('crypto');

const subtle = webcrypto.subtle;
const ITERATIONS = 200000;
const KEY_FILE = path.join(__dirname, '.kanban_key');
const IN_FILE = path.join(__dirname, 'data.json');
const OUT_FILE = path.join(__dirname, 'data.enc.json');
const HTML_FILE = path.join(__dirname, 'index.html');

const b64 = (u8) => Buffer.from(u8).toString('base64');
const unb64 = (s) => new Uint8Array(Buffer.from(s, 'base64'));

function resolvePassword(argv) {
  const i = argv.indexOf('--pw');
  if (i !== -1 && argv[i + 1]) return argv[i + 1];
  if (process.env.KANBAN_PW) return process.env.KANBAN_PW;
  if (fs.existsSync(KEY_FILE)) return fs.readFileSync(KEY_FILE, 'utf-8').trim();
  return null;
}

async function deriveKey(password, salt, iterations) {
  const km = await subtle.importKey(
    'raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
  return subtle.deriveKey(
    { name: 'PBKDF2', salt, iterations, hash: 'SHA-256' },
    km, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}

async function encrypt(plaintext, password) {
  const salt = webcrypto.getRandomValues(new Uint8Array(16));
  const iv = webcrypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(password, salt, ITERATIONS);
  const ct = await subtle.encrypt({ name: 'AES-GCM', iv }, key,
    new TextEncoder().encode(plaintext));
  return {
    v: 1,
    alg: 'AES-256-GCM',
    kdf: 'PBKDF2-SHA256',
    iter: ITERATIONS,
    salt: b64(salt),
    iv: b64(iv),
    ct: b64(new Uint8Array(ct)),
  };
}

async function decrypt(env, password) {
  const key = await deriveKey(password, unb64(env.salt), env.iter || ITERATIONS);
  const pt = await subtle.decrypt({ name: 'AES-GCM', iv: unb64(env.iv) },
    key, unb64(env.ct));
  return new TextDecoder().decode(pt);
}

/** 把密文信封同步进 index.html 内嵌块（替代原来的明文） */
function syncEmbed(envelope) {
  if (!fs.existsSync(HTML_FILE)) { console.log('   ⚠️  index.html 不存在，跳过内嵌同步'); return false; }
  const html = fs.readFileSync(HTML_FILE, 'utf-8');
  const pat = /(<script id="dataEmbed" type="application\/json">)([\s\S]*?)(<\/script>)/;
  if (!pat.test(html)) { console.log('   ⚠️  未找到 dataEmbed 块，跳过内嵌同步'); return false; }
  const payload = '\n' + JSON.stringify(envelope) + '\n';
  fs.writeFileSync(HTML_FILE, html.replace(pat, (_m, a, _b, c) => a + payload + c), 'utf-8');
  console.log(`   ✅ 已同步密文到 index.html 内嵌块 (${payload.length} 字符)`);
  return true;
}

(async () => {
  const argv = process.argv.slice(2);

  if (argv.includes('-h') || argv.includes('--help')) {
    console.log(fs.readFileSync(__filename, 'utf-8').split('*/')[0]);
    return;
  }

  const password = resolvePassword(argv);
  if (!password) {
    console.error('❌ 未提供口令。请用 --pw "口令" ，或设置环境变量 KANBAN_PW ，或先创建 .kanban_key 文件');
    process.exit(1);
  }
  if (password.length < 6) {
    console.error('❌ 口令过短（至少 6 位）');
    process.exit(1);
  }

  // --verify：解密现有 data.enc.json 并与 data.json 比对
  if (argv.includes('--verify')) {
    if (!fs.existsSync(OUT_FILE)) { console.error('❌ data.enc.json 不存在'); process.exit(1); }
    const env = JSON.parse(fs.readFileSync(OUT_FILE, 'utf-8'));
    const got = await decrypt(env, password);
    const want = fs.readFileSync(IN_FILE, 'utf-8');
    const same = JSON.stringify(JSON.parse(got)) === JSON.stringify(JSON.parse(want));
    console.log(same ? '✅ 校验通过：解密结果与 data.json 一致' : '❌ 校验失败：解密结果不一致');
    process.exit(same ? 0 : 1);
  }

  if (!fs.existsSync(IN_FILE)) { console.error(`❌ 找不到 ${IN_FILE}`); process.exit(1); }
  const plaintext = fs.readFileSync(IN_FILE, 'utf-8');
  JSON.parse(plaintext); // 提前校验 JSON 合法性

  console.log(`🔐 加密 ${path.basename(IN_FILE)} (${plaintext.length} 字符) ...`);
  const t0 = Date.now();
  const envelope = await encrypt(plaintext, password);
  fs.writeFileSync(OUT_FILE, JSON.stringify(envelope, null, 2), 'utf-8');
  console.log(`   ✅ 已生成 ${path.basename(OUT_FILE)} (${fs.statSync(OUT_FILE).size} 字节, ${Date.now() - t0}ms)`);

  if (argv.includes('--save-key')) {
    fs.writeFileSync(KEY_FILE, password, 'utf-8');
    console.log('   🔑 口令已保存到 .kanban_key（已 gitignore，勿提交）');
  }

  syncEmbed(envelope);
  console.log('🎉 完成。明文仅存于本地，仓库内只有密文。');
})().catch((e) => { console.error('❌ 失败:', e.message); process.exit(1); });
