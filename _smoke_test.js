// 前端冒烟测试：node 注入 DOM/ECharts/fetch stub + data.json → renderAll + 4 tab 渲染
// 用法: node _smoke_test.js [data.json] [index.html]
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const dataFile = process.argv[2] || 'data.json';
const htmlFile = process.argv[3] || 'index.html';
const dataJson = fs.readFileSync(dataFile, 'utf-8');
// 从 index.html 提取最后一个 <script> 块（主逻辑脚本，非内嵌 dataEmbed）
const html = fs.readFileSync(path.join(__dirname, htmlFile), 'utf-8');
const scripts = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
if (!scripts.length) { console.error('✗ index.html 中未找到 <script> 块'); process.exit(1); }
const mainSrc = scripts[scripts.length - 1];
if (mainSrc.length < 1000) { console.error('✗ 主脚本过短，疑似提取失败'); process.exit(1); }
console.log(`ℹ️ 主脚本 ${mainSrc.length} 字符，来自 ${htmlFile}`);

// ---------- DOM / 浏览器 stub ----------
function makeEl(id) {
  const el = {
    id, textContent: '', innerHTML: '', value: '', checked: false, disabled: false,
    className: '', href: '', src: '', alt: '', title: '', type: '', placeholder: '',
    files: [], style: {}, dataset: {},
    offsetWidth: 800, offsetHeight: 400, clientWidth: 800, clientHeight: 400,
    classList: {
      add() {}, remove() {}, toggle() { return false; }, contains() { return false; },
      replace() {}
    },
    setAttribute() {}, getAttribute() { return null; }, removeAttribute() {},
    appendChild(c) { return c; }, removeChild(c) { return c; }, insertBefore(c) { return c; },
    replaceChild(c) { return c; }, cloneNode() { return makeEl(id + ':clone'); },
    addEventListener() {}, removeEventListener() {}, dispatchEvent() { return true; },
    querySelector() { return makeEl(id + ':child'); },
    querySelectorAll() { return []; },
    getBoundingClientRect() { return { top: 0, left: 0, width: 800, height: 400, right: 800, bottom: 400 }; },
    getContext() { return { canvas: { width: 0, height: 0 }, measureText() { return { width: 10 }; } }; },
    focus() {}, click() {}, scrollIntoView() {}, remove() {},
    setInterval() { return 0; }, toDataURL() { return 'data:image/png;base64,'; },
    innerText: '', outerHTML: '', children: [], firstChild: null, nextSibling: null,
  };
  el.getElementsByTagName = () => [];
  el.getElementsByClassName = () => [];
  return el;
}
const els = {};
function getEl(id) { if (!els[id]) els[id] = makeEl(id); return els[id]; }

const ctx = {
  console, Math, JSON, Date, Number, String, Boolean, Array, Object, RegExp, Error, Promise,
  parseInt, parseFloat, isNaN, isFinite, setTimeout, setInterval, clearTimeout, clearInterval,
  URLSearchParams, Blob, TextEncoder, TextDecoder, Map, Set, Symbol, Proxy, Reflect, WeakMap, WeakSet,
  getEl,
  document: {
    getElementById: getEl,
    querySelector: (s) => getEl(String(s)),
    querySelectorAll: () => [],
    addEventListener() {}, removeEventListener() {},
    createElement: (t) => makeEl(String(t)),
    createTextNode: (t) => makeEl('text:' + t),
    body: makeEl('body'),
    documentElement: makeEl('documentElement'),
    head: makeEl('head'),
    readyState: 'complete',
    location: { href: 'http://localhost/', search: '', hash: '', reload() {} },
  },
  echarts: {
    init() { return chartStub(); },
    getInstanceByDom() { return chartStub(); },
    dispose() {},
  },
  fetch: async () => { throw new Error('fetch stub: 网络被 stub'); },
  localStorage: { getItem: () => null, setItem() {}, removeItem() {}, clear() {} },
  sessionStorage: { getItem: () => null, setItem() {}, removeItem() {}, clear() {} },
  navigator: { userAgent: 'smoke', clipboard: { writeText: async () => {} } },
  requestAnimationFrame: (cb) => 0, cancelAnimationFrame: () => {},
  matchMedia: () => ({ matches: false, addListener() {}, removeListener() {} }),
  ResizeObserver: function () { this.observe = () => {}; this.unobserve = () => {}; this.disconnect = () => {}; },
  IntersectionObserver: function () { this.observe = () => {}; this.unobserve = () => {}; this.disconnect = () => {}; },
  FileReader: function () { this.readAsDataURL = () => {}; this.readAsText = () => {}; },
  location: { href: 'http://localhost/', search: '', reload() {} },
  history: { pushState() {} },
  getComputedStyle: () => ({ getPropertyValue: () => '' }),
  XMLHttpRequest: function () {},
  DevicePixelRatio: 1,
  alert: (m) => { console.log('[alert]', m); },
  confirm: () => true,
  prompt: () => null,
};
function chartStub() {
  return {
    setOption() {}, getOption() { return { series: [] }; }, resize() {}, on() {}, off() {},
    dispatchAction() {}, dispose() {}, showLoading() {}, hideLoading() {}, clear() {},
    getWidth() { return 800; }, getHeight() { return 400; }, getDom() { return makeEl('chart'); },
    convertToPixel() { return [0, 0]; }, convertFromPixel() { return [0, 0]; },
    containsPoint() { return false; }, getZr() { return { on() {}, off() {} }; },
  };
}
ctx.window = ctx;
ctx.globalThis = ctx;
ctx.self = ctx;

// ---------- 执行主脚本 ----------
vm.createContext(ctx);
vm.runInContext(mainSrc, ctx, { filename: htmlFile });
// 注入真实 DATA（词法作用域内的 let DATA 需在同环境赋值）
vm.runInContext('DATA = ' + dataJson + ';', ctx, { filename: 'data-inject' });
// 桥接：让外部能读取词法作用域内的 DATA
vm.runInContext('__getData = () => DATA;', ctx, { filename: 'bridge' });

// ---------- 逐项渲染 ----------
const steps = ['renderAll', 'renderProgressTab', 'renderCapacityTab', 'renderBudgetTab', 'renderTaskTable'];
let pass = 0, fail = 0;
const liveData = (typeof ctx.__getData === 'function') ? ctx.__getData() : null;
for (const fn of steps) {
  try {
    const f = ctx[fn];
    if (typeof f !== 'function') { console.log(`✗ ${fn}: 函数不存在`); fail++; continue; }
    if (fn === 'renderAll') {
      f();
    } else {
      // tab 函数带真实参数独立调用（renderAll 已覆盖过默认路径）
      const tasks = (liveData && liveData.tasks) || [];
      // 用与 renderAll 一致的真实派生指标，避免空对象缺字段
      const m = (typeof ctx.calcMetrics === 'function') ? ctx.calcMetrics(tasks) : {};
      if (fn === 'renderTaskTable') f(tasks);
      else f(tasks, m);
    }
    console.log(`✓ ${fn} 通过`);
    pass++;
  } catch (e) {
    console.log(`✗ ${fn} 抛错: ${e.message}`);
    console.log(e.stack.split('\n').slice(0, 4).join('\n'));
    fail++;
  }
}
console.log(`\n结果: ${pass} 通过 / ${fail} 失败`);
process.exit(fail ? 1 : 0);
