/**
 * static-check.js — 免依赖静态校验 + taskpane.js 状态机模拟测试
 * 运行: node tools/static-check.js   (在项目根目录)
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
let failures = 0;
function check(cond, msg) {
    console.log((cond ? '  [OK]  ' : '  [FAIL] ') + msg);
    if (!cond) failures++;
}

// ---------- 1. ribbon.xml 校验 ----------
console.log('\n== ribbon.xml ==');
const ribbonXml = fs.readFileSync(path.join(ROOT, 'ribbon.xml'), 'utf8');
const ribbonJs = fs.readFileSync(path.join(ROOT, 'js/ribbon.js'), 'utf8');

// 良构粗检: 标签配平
{
    const stack = [];
    const re = /<(\/?)([\w:]+)([^>]*?)\/?>/g;
    let m, balanced = true;
    const voidOk = new Set(['?xml']); // xml 声明单独处理
    const xmlDecl = /^<\?xml[^>]*\?>/;
    let s = ribbonXml.trim().replace(xmlDecl, '');
    while ((m = re.exec(s))) {
        const closing = m[1] === '/', tag = m[2], selfClose = /\/>$/.test(m[0]);
        if (closing) {
            const top = stack.pop();
            if (top !== tag) { balanced = false; console.log('    tag mismatch: </' + tag + '> vs <' + top + '>'); break; }
        } else if (!selfClose) {
            stack.push(tag);
        }
    }
    check(balanced && stack.length === 0, 'XML 标签配平');
}

// 回调函数存在性
for (const attr of ['onLoad', 'onAction', 'getImage']) {
    const re = new RegExp(attr + '="([\\w.]+)"', 'g');
    let m, names = new Set();
    while ((m = re.exec(ribbonXml))) names.add(m[1]);
    for (const n of names) {
        check(new RegExp('function\\s+' + n.replace('.', '\\.') + '\\s*\\(').test(ribbonJs),
            '回调函数存在: ' + attr + ' → ' + n);
    }
}

// 按钮 id 覆盖: OnAction 的 switch 分支 & 图标文件
{
    const ids = [...ribbonXml.matchAll(/<button\s[^>]*\bid="(\w+)"/g)].map(m => m[1]);
    check(ids.length >= 4, '按钮数量 = ' + ids.length);
    for (const id of ids) {
        check(ribbonJs.includes("case '" + id + "'"), 'OnAction 分支覆盖: ' + id);
    }
    const imgMap = [...ribbonJs.matchAll(/'(\w+)':\s*'(images\/[\w.-]+)'/g)];
    for (const [, id, img] of imgMap) {
        check(fs.existsSync(path.join(ROOT, img)), '图标存在: ' + img + ' (' + id + ')');
    }
}

// ---------- 2. 入口/引用完整性 ----------
console.log('\n== 文件引用 ==');
check(fs.existsSync(path.join(ROOT, 'index.html')), 'index.html 存在');
for (const f of ['js/enums.js', 'js/util.js', 'js/core.js', 'js/ribbon.js']) {
    check(new RegExp(f.replace(/\//g, '\\/')).test(fs.readFileSync(path.join(ROOT, 'main.js'), 'utf8')),
        'main.js 引入 ' + f);
}
const tpHtml = fs.readFileSync(path.join(ROOT, 'ui/taskpane.html'), 'utf8');
for (const src of [...tpHtml.matchAll(/src="([^"]+)"/g)].map(m => m[1])) {
    check(fs.existsSync(path.join(ROOT, 'ui', src)), 'taskpane.html 引用存在: ' + src);
}
for (const id of ['status', 'moverName', 'step1', 'step2', 'aPos', 'bPos', 'delta', 'distNorm',
    'dur', 'trig', 'smooth', 'replaceAnim', 'keepMarkers', 'autoChk',
    'btnGen', 'btnCross', 'btnCleanSlide', 'btnDelAnim', 'btnDiag', 'result']) {
    check(tpHtml.includes('id="' + id + '"'), 'taskpane.html 元素存在: #' + id);
}

// ---------- 3. taskpane.js 状态机模拟测试 ----------
console.log('\n== taskpane.js 状态机模拟 ==');

// --- 最小 DOM mock ---
function makeEl(id) {
    return {
        id, textContent: '', innerHTML: '', className: '',
        value: id === 'dur' ? '2' : (id === 'pathForm' ? 'line' : '1'),
        checked: id === 'autoChk' || id === 'snapChk',
        style: {}, onclick: null, onchange: null,
        addEventListener() { }
    };
}
const els = {};
['status', 'moverName', 'step1', 'step2', 'aPos', 'bPos', 'cdPos', 'delta', 'distNorm', 'dur', 'trig',
    'smooth', 'replaceAnim', 'keepMarkers', 'autoChk', 'snapChk', 'btnGen', 'btnCross', 'btnCleanSlide',
    'btnDelAnim', 'btnDiag', 'btnPlaceC', 'btnPlaceD', 'btnClearCD', 'result',
    'pathForm', 'rowLine', 'rowPara', 'rowArc', 'rowWp', 'rowTrace', 'rowAnchor', 'traceHint',
    'paraPos', 'arcPos', 'wpPos', 'tracePos', 'formHint', 'smoothW',
    'btnParaD', 'btnParaP', 'btnArcD', 'btnArcM', 'btnAddW', 'btnClearW', 'btnTrace', 'btnTraceDone',
    'btnTracePoint', 'btnAnchor', 'btnAnchorClear']
    .forEach(id => els[id] = makeEl(id));

let intervalFn = null;
let timeOffset = 0;   // 模拟时间流逝(跨过冷却期)
const confirmAnswers = [];
const realDateNow = Date.now;
const sandbox = {
    document: { getElementById: id => els[id] || (els[id] = makeEl(id)) },
    window: { confirm: () => confirmAnswers.length ? confirmAnswers.shift() : true },
    setInterval: fn => { intervalFn = fn; return 1; },
    clearInterval: () => { },
    Date: { now: () => realDateNow() + timeOffset },
    console, Math, parseFloat, parseInt, String, Object, Array, RegExp, JSON, Error
};

// --- 真实 core.js 数学函数(独立 vm 加载, 供 mock 桥接与期望值计算) ---
const REAL = (() => {
    const ctx = { console, isFinite, parseFloat, Math, String, Number, RegExp, JSON };
    vm.createContext(ctx);
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'js/enums.js'), 'utf8'), ctx);
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'js/core.js'), 'utf8'), ctx);
    return ctx.ABP;
})();

// --- mock ABP: 模拟一个 WPS 文档世界 ---
const calls = { createPathEffect: null, cleanFlow: 0, preview: 0, cleanedSlides: [] };
const psStore = {};   // PluginStorage mock
const psMock = {
    getItem: k => (k in psStore ? psStore[k] : null),
    setItem: (k, v) => { psStore[k] = String(v); }
};
let presName = '测试.pptx';
function Shape(id, name, l, t, w, h, slideIdx) {
    return { Id: id, Name: name, Left: l, Top: t, Width: w, Height: h, slideIdx: slideIdx || 1,
        deleted: false,
        Delete() { this.deleted = true; }, Select() { calls.selected = this.Name; } };
}
const mover = Shape(101, '矩形1', 100, 100, 120, 80, 1);   // 幻灯片1, 中心 (160,140)
const other = Shape(102, '圆形1', 500, 300, 80, 80, 1);    // 幻灯片1, 中心 (540,340)
const shp2  = Shape(103, '星形2', 200, 200, 60, 60, 2);    // 幻灯片2 (跨页切换测试)
const bMark = Shape(900, '__pthlp_B__', 0, 0, 16, 16, 1);
const aMark = Shape(901, '__pthlp_A__', 0, 0, 14, 14, 1);
const pcMark = Shape(902, '__pthlp_PC__', -50, -50, 14, 14, 1);
const pdMark = Shape(903, '__pthlp_PD__', -50, -50, 14, 14, 1);
const pmMark = Shape(904, '__pthlp_PM__', -50, -50, 14, 14, 1);
const trMark = Shape(905, '__pthlp_TR__', -50, -50, 14, 14, 1);
pmMark.deleted = true; trMark.deleted = true;
const wMarks = [];            // 航点
const snapCross = Shape(906, '__pthlp_C_102', 533, 333, 14, 14, 1);  // other 中心处的红十字(吸附目标)
pcMark.deleted = true; pdMark.deleted = true;   // 初始未放置
let snapTargetMock = null;    // nearestSnapTarget 返回值控制
const allShapes = [mover, other, shp2, bMark, aMark, pcMark, pdMark, pmMark, trMark, snapCross];
let currentSelection = [];   // Shape 数组

const mockABP = {
    MARK: '__pthlp_',
    APP: () => ({ PluginStorage: psMock }),
    getPresentation: () => ({ Name: presName }),
    parentSlide: s => ({ idx: s.slideIdx }),
    locateShape: raw => (raw ? { slide: { idx: raw.slideIdx || 1 }, shape: raw } : null),
    getActiveSlide: () => ({ idx: 1 }),
    slideIndexOf: s => s.idx,
    slideByIndex: (p, i) => (i >= 1 && i <= 2 ? { idx: i } : null),
    slideSize: () => ({ w: 960, h: 540, ok: true }),
    shapeIdOf: s => s.Id,
    isMarker: s => s.Name.indexOf('__pthlp_') === 0,
    selectionInfo: () => ({ ok: true, type: 2, shapes: currentSelection.slice(), msg: '' }),
    findShapeById: (slide, id) => allShapes.find(s => s.Id === id && !s.deleted) || null,
    looksAlive: s => !!s && !s.deleted && !s.bad && !(s.Left === 0 && s.Top === 0 && s.Width === 0 && s.Height === 0),
    refetchShapeFresh: (pres, slideIdx, shapeId, shapeName) => {
        const s = allShapes.find(x => x.Id === shapeId && !x.deleted && !x.bad) || null;
        return s;
    },
    center: s => {
        if (s.deleted || s.bad) return null;   // 模拟坐标不可读(僵尸引用/桥接异常)
        return { x: s.Left + s.Width / 2, y: s.Top + s.Height / 2 };
    },
    ensureAMarker: (sl, c) => { aMark.Left = c.x - 7; aMark.Top = c.y - 7; return aMark; },
    ensureBMarker: (sl, c) => {
        if (bMark.deleted) { bMark.deleted = false; bMark.Left = c.x - 8; bMark.Top = c.y - 8; }
        return bMark;
    },
    syncFlowMarkers: (sl, c) => mockABP.syncMarkers(sl, c),
    syncMarkers: (sl, c, form) => {
        form = form || 'line';
        aMark.Left = c.x - 7; aMark.Top = c.y - 7;
        const pcAlive = !pcMark.deleted, pdAlive = !pdMark.deleted;
        const cdMode = pcAlive && pdAlive;
        const pack = (b) => ({
            a: aMark, b,
            pc: pcAlive ? pcMark : null, pd: pdAlive ? pdMark : null,
            pm: pmMark.deleted ? null : pmMark,
            tr: trMark.deleted ? null : trMark,
            ws: wMarks.filter(w => !w.deleted),
            cdMode, err: null
        });
        if (form === 'line' && !cdMode) {
            if (bMark.deleted) { bMark.deleted = false; bMark.Left = c.x - 8; bMark.Top = c.y - 8; }
            return pack(bMark);
        }
        if (!bMark.deleted) bMark.deleted = true;
        return pack(null);
    },
    scanAllMarkers: sl => ({
        a: aMark.deleted ? null : aMark, b: bMark.deleted ? null : bMark,
        pc: pcMark.deleted ? null : pcMark, pd: pdMark.deleted ? null : pdMark,
        pm: pmMark.deleted ? null : pmMark, tr: trMark.deleted ? null : trMark,
        ws: wMarks.filter(w => !w.deleted)
    }),
    placeNamedMarker: (sl, name, x, y, size, color) => {
        let m;
        if (name === '__pthlp_TR__') m = trMark;
        else if (name === '__pthlp_PM__') m = pmMark;
        else if (name === '__pthlp_PD__') m = pdMark;
        else if (name === '__pthlp_PC__') m = pcMark;
        else if (name.indexOf('__pthlp_W') === 0) {
            m = wMarks.find(w => w.Name === name);
            if (!m) { m = Shape(920 + wMarks.length + 1, name, 0, 0, 14, 14, 1); wMarks.push(m); allShapes.push(m); }
        } else return null;
        m.deleted = false; m.Left = x - 7; m.Top = y - 7;
        return m;
    },
    deleteMarkersByPrefix: (sl, prefix) => {
        let n = 0;
        [bMark, aMark, pcMark, pdMark, pmMark, trMark].concat(wMarks).forEach(s => {
            if (!s.deleted && s.Name.indexOf(prefix) === 0) { s.deleted = true; n++; }
        });
        return n;
    },
    updateRoutePreview: (sl, pts) => { calls.pvCalls = (calls.pvCalls || 0) + 1; calls.pvPts = pts; return pts.length; },
    // 真实数学桥接
    pathString: REAL.pathString,
    parabolaPath: REAL.parabolaPath,
    arcPath: REAL.arcPath,
    polylinePath: REAL.polylinePath,
    smoothPath: REAL.smoothPath,
    smoothPts: REAL.smoothPts,
    pathEndpoint: REAL.pathEndpoint,
    createPathEffectByPath: (pres, slideIdx, shapeId, pathStr, endRel, opts, shapeRef) => {
        calls.byPath = { pathStr, endRel, opts };
        return { ok: true, strategy: 1, msg: ['ok'], readback: { kind: 'path', dx: endRel && endRel.x, dy: endRel && endRel.y, raw: pathStr }, effect: null };
    },
    placePointMarker: (sl, kind, c) => {
        const m = kind === 'C' ? pcMark : pdMark;
        m.deleted = false;
        m.Left = c.x - 7 + (kind === 'C' ? -42 : 48);
        m.Top = c.y - 7 + (kind === 'C' ? -32 : 36);
        return m;
    },
    clearPointMarkers: () => {
        const n = (!pcMark.deleted ? 1 : 0) + (!pdMark.deleted ? 1 : 0);
        pcMark.deleted = true; pdMark.deleted = true;
        return n;
    },
    nearestSnapTarget: (sl, p) => snapTargetMock,
    snapMarkerTo: (m, t) => { m.Left = t.x - m.Width / 2; m.Top = t.y - m.Height / 2; return true; },
    updatePreviewLine: (sl, p1, p2) => { calls.preview++; calls.lastLine = { p1, p2 }; return {}; },
    cleanFlowMarkers: sl => {
        calls.cleanFlow++;
        if (sl && sl.idx) calls.cleanedSlides.push(sl.idx);
        // 与真实实现一致: 删除流程标记(A/B/C/D/P/T/航点), 保留中心十字
        let n = 0;
        [bMark, aMark, pcMark, pdMark, pmMark, trMark].concat(wMarks).forEach(s => {
            if (!s.deleted) { s.deleted = true; n++; }
        });
        return n;
    },
    cleanSlideMarkers: () => 0,
    deleteShapeAnimations: () => 2,
    toggleCenterCross: () => true,
    diagnose: () => [{ name: 'x', ok: true, detail: 'd', error: null }]
};

{
    const ctx = { console };
    vm.createContext(ctx);
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'js/enums.js'), 'utf8'), ctx);
    sandbox.ABP_ENUM = ctx.ABP_ENUM;
}
sandbox.ABP = mockABP;

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'js/taskpane.js'), 'utf8'), sandbox);

function tick(n) { for (let i = 0; i < (n || 1); i++) intervalFn(); }
function approx(a, b, eps) { return Math.abs(a - b) < (eps || 1e-6); }

// T1: 选中 mover → 自动开始, A 标记定位到中心
currentSelection = [mover];
tick(1);
check(els.moverName.textContent === '矩形1', 'T1 自动识别选中图形');
check(approx(aMark.Left + 7, 160) && approx(aMark.Top + 7, 140), 'T1 A 十字位于图形中心 (160,140)');
check(els.status.className === 'ok', 'T1 状态栏进入引导态');

// T2: 拖动 B 标记到 other 中心 (540,340) → 坐标/位移更新, 预览线刷新
bMark.Left = 540 - 8; bMark.Top = 340 - 8;
tick(1);
check(els.bPos.textContent.includes('540') && els.bPos.textContent.includes('340'), 'T2 B 坐标实时更新');
check(els.delta.textContent.includes('380') && els.delta.textContent.includes('200'), 'T2 位移显示 Δ(380,200)');
check(els.distNorm.textContent.includes('0.3958') && els.distNorm.textContent.includes('0.3704'), 'T2 归一化位移显示');
check(calls.preview >= 1 && approx(calls.lastLine.p2.x, 540), 'T2 预览线端点=B 中心');

// T3: 生成 → createPathEffectByPath 收到精确归一化位移 (380/960, 200/540)
els.btnGen.onclick();
check(calls.byPath && calls.byPath.pathStr.indexOf('M 0 0 L') === 0, 'T3 生成直线路径');
check(calls.byPath && approx(calls.byPath.endRel.x, 380 / 960, 1e-6) && approx(calls.byPath.endRel.y, 200 / 540, 1e-6),
    'T3 归一化位移精确 (dx=0.395833, dy=0.370370)');
check(calls.byPath && calls.byPath.opts.duration === 2 && calls.byPath.opts.trigger === 1, 'T3 参数传递(时长/触发)');
check(calls.cleanFlow >= 1, 'T3 生成后清理流程标记');
check(els.result.innerHTML.includes('策略1'), 'T3 成功结果展示');

// T4: 生成后冷却期不自动重启; 冷却后选中新图形自动开始
calls.byPath = null;
tick(2);
check(calls.byPath === null && els.moverName.textContent === '未选择', 'T4 冷却期不自动重启');
timeOffset += 2000;                       // 越过 1500ms 冷却
currentSelection = [other];
tick(1);
check(els.moverName.textContent === '圆形1', 'T4 冷却后选中新图形自动开始');

// T5: 流程中切换图形 → 以新图形重新开始
currentSelection = [mover];
tick(1);
check(els.moverName.textContent === '矩形1', 'T5 切换到新图形');

// T6: B 标记被删 → 自动重生于 A
bMark.deleted = true;
bMark.Left = -999; bMark.Top = -999;
tick(1);
check(!bMark.deleted && approx(bMark.Left + 8, mover.Left + 60), 'T6 B 标记误删后自动重生');

// T7: 位移≈0 → confirm 弹窗
timeOffset += 2000;
confirmAnswers.push(false);
els.btnGen.onclick();
check(calls.byPath === null || Math.abs(calls.byPath.endRel.x) > 1e-9, 'T7 零位移可取消');
confirmAnswers.push(true);
els.btnGen.onclick();
check(calls.byPath && approx(calls.byPath.endRel.x, 0), 'T7 确认后可生成零位移动画');

// T8: 工具按钮 — 删除动画 / 诊断
currentSelection = [mover];
tick(1);
els.btnDelAnim.onclick();
check(els.status.textContent.includes('2 个动画'), 'T8 删除图形动画');
els.btnDiag.onclick();
check(els.result.innerHTML.includes('API 诊断报告'), 'T8 诊断报告渲染');

// T9: 流程进行中切到另一幻灯片的图形 → 旧页(幻灯片1)标记被清理, 新页开始
timeOffset += 2000;
currentSelection = [other];
tick(1);                                   // 先在幻灯片1重启流程(圆形1)
check(els.moverName.textContent === '圆形1', 'T9 前置: 流程进行中');
currentSelection = [shp2];
calls.cleanedSlides.length = 0;
tick(1);
check(els.moverName.textContent === '星形2', 'T9 切换到其他幻灯片图形');
check(calls.cleanedSlides.indexOf(1) >= 0, 'T9 旧幻灯片流程标记已清理');

// T10: 切换演示文稿 → 流程重置并提示
presName = '另一个文档.pptx';
tick(1);
check(els.moverName.textContent === '未选择', 'T10 切换演示文稿后流程重置');
check(els.status.className === 'warn', 'T10 切换文稿给出提示');
presName = '测试.pptx';
tick(1);

// T11: 功能区「清理标记」信号 → 流程立即停止不重建
timeOffset += 2000;
currentSelection = [mover];
tick(1);
check(els.moverName.textContent === '矩形1', 'T11 前置: 流程进行中');
psStore['abp_clean_seq'] = '1';
tick(1);
check(els.moverName.textContent === '未选择', 'T11 清理信号后流程停止');
check(els.status.textContent.includes('功能区清理'), 'T11 清理联动提示');

// T12: 生成后选区保持同一图形不自动重启; 选区变化后恢复自动
currentSelection = [other];
tick(1);
bMark.Left = 400; bMark.Top = 300;
tick(1);
els.btnGen.onclick();
timeOffset += 2000;   // 越过冷却期: 此处检验的是选区签名门控而非冷却
tick(3);
check(els.moverName.textContent === '未选择', 'T12 生成后同图形不自动重启');
currentSelection = [];          // 点空白(空选) → 解除门控
tick(1);
currentSelection = [mover];     // 重新点选图形 → 自动开始
tick(1);
check(els.moverName.textContent === '矩形1', 'T12 选区变化后自动恢复');

// T13: 移动图形被删除 → 重置一次即停, 不再无限闪烁; 重新点选可恢复
mover.deleted = true;                 // 模拟用户删除了移动图形(引用变僵尸, Id扫描也找不到)
tick(1);
check(els.moverName.textContent === '未选择', 'T13 图形消失后流程重置');
check(els.status.textContent.includes('重新点选'), 'T13 重置给出引导提示');
tick(5);                              // 连续多拍: 选区没变 → 不允许自动重启(防闪)
check(els.moverName.textContent === '未选择', 'T13 无闪烁循环(选区未变不重启)');
currentSelection = [other];           // 重新点选其他图形 → 正常开始
tick(1);
check(els.moverName.textContent === '圆形1', 'T13 重新点选后恢复');

// T14: 坐标全 0 的僵尸图形 → 明确报错并终止流程, 不再把标记画到(0,0)左上角
timeOffset += 2000;
const bad = Shape(110, '坏图形', 0, 0, 60, 60, 1);
bad.bad = true;                        // 模拟坐标全 0 的僵尸引用
allShapes.push(bad);
currentSelection = [bad];
tick(1);
check(els.moverName.textContent === '未选择', 'T14 僵尸图形触发流程重置');
check(els.status.textContent.includes('坐标读取异常'), 'T14 给出明确错误提示');
check(aMark.Left > 0 && bMark.Left > 0, 'T14 标记未被错误画到左上角(0,0)');
currentSelection = [other];            // 恢复正常
tick(1);
check(els.moverName.textContent === '圆形1', 'T14 后续正常图形不受影响');

// T15: 标记点 C→D 模式 —— 放C/D, 位移 = D-C; 生成按标记点对齐
timeOffset += 2000;
mover.deleted = false;                 // 恢复 T13 模拟删除的 mover
currentSelection = [mover];
tick(1);
check(els.moverName.textContent === '矩形1', 'T15 前置: 流程进行中');
els.btnPlaceC.onclick();
check(!pcMark.deleted && approx(pcMark.Left + 7, 118), 'T15 放置C点(默认偏移左上-42,-32可见)');
els.btnPlaceD.onclick();
check(!pdMark.deleted && approx(pdMark.Left + 7, 208), 'T15 放置D点(默认偏移右下48,36)');
pcMark.Left = 150; pcMark.Top = 90;    // 用户把 C 拖到矩形左上角(参照点)
pdMark.Left = 533; pdMark.Top = 333;   // 用户把 D 拖到圆形中心红十字处
tick(1);
check(els.cdPos.textContent.includes('标记点 C→D'), 'T15 界面进入标记点模式');
check(bMark.deleted, 'T15 标记点模式下 B 标记自动移除');
els.btnGen.onclick();
check(calls.byPath && approx(calls.byPath.endRel.x, (540 - 157) / 960, 1e-6),
    'T15 生成位移 = D中心-C中心 (dx=(540-157)/960)');
check(approx(calls.byPath.endRel.y, (340 - 97) / 540, 1e-6), 'T15 dy=(340-97)/540');
check(els.result.innerHTML.includes('标记点 C→D'), 'T15 结果标注标记点模式');
els.btnClearCD.onclick();
timeOffset += 2000;                     // 越过生成冷却(虚拟时钟)
currentSelection = [];                  // 空选解除生成门控
tick(1);
currentSelection = [mover];             // 重新激活流程 → B 标记重建
tick(1);
check(pcMark.deleted && pdMark.deleted && !bMark.deleted, 'T15 清除标记点恢复中心模式');

// T16: 自动吸附 —— B 拖到红十字附近松手(位置稳定), 两拍后精确吸附
timeOffset += 2000;
calls.byPath = null;
bMark.Left = 528; bMark.Top = 331;     // 距离目标(540,341) 约12pt内但不重合
snapTargetMock = { x: 540, y: 341, dist: 5 };
tick(1);                               // 第1拍: 记录位置
tick(1);                               // 第2拍: 刚稳定 → 吸附
check(approx(bMark.Left + 8, 540) && approx(bMark.Top + 8, 341), 'T16 松手后自动吸附到红十字中心');
check(els.status.textContent.includes('吸附'), 'T16 吸附状态提示');
tick(2);                               // 已在目标上 → 不重复写
check(approx(bMark.Left + 8, 540), 'T16 已吸附后不重复吸附');
els.snapChk.checked = false;           // 关闭吸附后允许自由微调
bMark.Left = 520; bMark.Top = 320;
tick(2);
check(approx(bMark.Left, 520), 'T16 关闭吸附选项后不再吸附');
els.snapChk.checked = true;
snapTargetMock = null;

// ---------- 路径形态 T17-T20 ----------
const sizeMock = { w: 960, h: 540 };
function setForm(f) {
    els.pathForm.value = f;
    els.pathForm.onchange.call(els.pathForm);
}
function restartFlow() {
    timeOffset += 2000;
    currentSelection = [];
    tick(1);
    currentSelection = [mover];
    tick(1);
}

// T17 抛物线: 锚点=中心(160,140), D=(540,340), P=(300,180) → 路径串与真实数学一致
restartFlow();
setForm('parabola');
check(els.rowPara.style.display !== 'none' && els.rowLine.style.display === 'none', 'T17 形态行切换');
els.btnParaD.onclick();
els.btnParaP.onclick();
check(!pdMark.deleted && !pmMark.deleted, 'T17 放置 D/P 标记');
pdMark.Left = 540 - 7; pdMark.Top = 340 - 7;
pmMark.Left = 300 - 7; pmMark.Top = 180 - 7;
tick(1);
check(els.cdPos.textContent.includes('抛物线'), 'T17 界面进入抛物线形态');
els.btnGen.onclick();
const expPb = REAL.parabolaPath({ x: 160, y: 140 }, { x: 300, y: 180 }, { x: 540, y: 340 }, sizeMock);
check(calls.byPath && calls.byPath.pathStr === expPb.path, 'T17 抛物线路径串与真实数学一致');
check(calls.byPath && calls.byPath.pathStr.indexOf('M 0 0 C') === 0, 'T17 贝塞尔命令');
check(calls.byPath && approx(calls.byPath.endRel.x, 380 / 960, 1e-6) && approx(calls.byPath.endRel.y, 200 / 540, 1e-6), 'T17 抛物线终点');

// T18 圆弧: 锚点(160,140) M(300,60) D(540,340) → 三点定圆
restartFlow();
setForm('arc');
els.btnArcD.onclick();
els.btnArcM.onclick();
pdMark.Left = 540 - 7; pdMark.Top = 340 - 7;
pmMark.Left = 300 - 7; pmMark.Top = 60 - 7;
tick(1);
check(els.cdPos.textContent.includes('圆弧'), 'T18 界面进入圆弧形态');
els.btnGen.onclick();
const expArc = REAL.arcPath({ x: 160, y: 140 }, { x: 300, y: 60 }, { x: 540, y: 340 }, sizeMock);
check(calls.byPath && calls.byPath.pathStr === expArc.path, 'T18 圆弧路径串与真实数学一致');
check(calls.byPath && calls.byPath.pathStr.indexOf('M 0 0 C') === 0, 'T18 贝塞尔命令');

// T19 航点: 2 个航点 → 折线; 平滑 → Catmull-Rom
restartFlow();
setForm('waypoints');
els.btnAddW.onclick();
els.btnAddW.onclick();
check(wMarks.length === 2 && wMarks.every(w => !w.deleted), 'T19 添加2个航点');
wMarks[0].Left = 300 - 7; wMarks[0].Top = 200 - 7;
wMarks[1].Left = 480 - 7; wMarks[1].Top = 300 - 7;
tick(1);
check(els.wpPos.textContent.includes('2 个航点'), 'T19 航点计数显示');
els.btnGen.onclick();
const expLine2 = REAL.polylinePath({ x: 160, y: 140 }, [{ x: 300, y: 200 }, { x: 480, y: 300 }], sizeMock);
check(calls.byPath && calls.byPath.pathStr === expLine2.path, 'T19 折线路径串一致');
els.smoothW.checked = true;
timeOffset += 2000; currentSelection = []; tick(1); currentSelection = [mover]; tick(1);
setForm('waypoints');
els.btnClearW.onclick();                 // 清空旧航点再重建
els.btnAddW.onclick(); els.btnAddW.onclick();
const aliveW = wMarks.filter(w => !w.deleted);
check(aliveW.length === 2, 'T19 清空后重建2航点');
aliveW[0].Left = 300 - 7; aliveW[0].Top = 200 - 7;
aliveW[1].Left = 480 - 7; aliveW[1].Top = 300 - 7;
tick(1);
els.btnGen.onclick();
const expSm = REAL.smoothPath({ x: 160, y: 140 }, [{ x: 160, y: 140 }, { x: 300, y: 200 }, { x: 480, y: 300 }], sizeMock);
check(calls.byPath && calls.byPath.pathStr === expSm.path, 'T19 平滑曲线路径串一致');

// T20 描迹: 开始→拖动采样→完成→生成折线
restartFlow();
setForm('trace');
els.btnTrace.onclick();
check(!trMark.deleted && els.status.textContent.includes('描迹中'), 'T20 开始描迹');
trMark.Left = 220 - 7; trMark.Top = 160 - 7; tick(1);
trMark.Left = 320 - 7; trMark.Top = 210 - 7; tick(1);
trMark.Left = 440 - 7; trMark.Top = 290 - 7; tick(1);
check(els.tracePos.textContent.includes('4 点'), 'T20 轨迹采样 4 点(含起点): ' + els.tracePos.textContent);
els.btnTraceDone.onclick();
check(els.status.textContent.includes('描迹完成'), 'T20 完成描迹');
els.btnGen.onclick();
check(calls.byPath && calls.byPath.pathStr.indexOf('M 0 0 L') === 0, 'T20 描迹折线路径');
check(calls.byPath && calls.byPath.pathStr.split('L').length >= 4, 'T20 多段折线');
setForm('line');

// T21: 标记锁定 —— 放C后点其他图形, 流程与标记不被清除; 清除标记后恢复切换
timeOffset += 2000; currentSelection = []; tick(1); currentSelection = [mover]; tick(1);
check(els.moverName.textContent === '矩形1', 'T21 前置: 流程在矩形1');
els.btnPlaceC.onclick();
currentSelection = [other];                  // 点了另一个图形
tick(2);
check(els.moverName.textContent === '矩形1', 'T21 移动图形未切换(标记锁定)');
check(!pcMark.deleted, 'T21 C 标记保留');
check(els.status.textContent.includes('锁定'), 'T21 锁定提示');
els.btnClearCD.onclick();                    // 清除标记 → 解锁
tick(1);
check(els.moverName.textContent === '圆形1', 'T21 清除标记后恢复自动切换');

// T22: 点选他图形后放D → D 直接落其中心
timeOffset += 2000; currentSelection = []; tick(1); currentSelection = [mover]; tick(1);
els.btnPlaceC.onclick();
currentSelection = [other];                  // 先点目标图形
els.btnPlaceD.onclick();                     // 再放D → 应落在圆形中心(540,340)
check(!pdMark.deleted && approx(pdMark.Left + 7, 540) && approx(pdMark.Top + 7, 340), 'T22 D点直接落在所选图形中心');
check(els.status.textContent.includes('所选图形中心'), 'T22 提示落在所选中心');
els.btnClearCD.onclick();
currentSelection = [mover]; tick(1);
setForm('line');

// T23: 描迹形态可放C起点(锚点行显示); 完成后可继续描迹; 少点生成弹确认
timeOffset += 2000; currentSelection = []; tick(1); currentSelection = [mover]; tick(1);
setForm('trace');
check(els.rowAnchor.style.display !== 'none', 'T23 非直线形态显示锚点行');
els.btnAnchor.onclick();
check(!pcMark.deleted && approx(pcMark.Left + 7, 160 - 42), 'T23 C起点放置(默认左上偏移)');
els.btnTrace.onclick();
check(els.status.textContent.includes('描迹中'), 'T23 开始描迹(分段提示)');
trMark.Left = 300 - 7; trMark.Top = 200 - 7; tick(1);
els.btnTraceDone.onclick();
check(els.status.textContent.includes('描迹完成'), 'T23 完成(仅2点)');
check(els.status.textContent.includes('继续描迹'), 'T23 少点提示可继续');
els.btnTrace.onclick();
check(els.status.textContent.includes('继续描迹'), 'T23 恢复模式继续描迹');
trMark.Left = 400 - 7; trMark.Top = 260 - 7; tick(1);
check(els.tracePos.textContent.includes('3 点'), 'T23 恢复后累计3点');
calls.byPath = null;                            // 清旧值再测确认流程
confirmAnswers.push(false);                     // 少点生成 → 取消
els.btnGen.onclick();
check(!calls.byPath, 'T23 少点生成确认可取消');
confirmAnswers.push(true);                      // 确认 → 生成
els.btnGen.onclick();
check(calls.byPath && calls.byPath.pathStr.split('L').length >= 3, 'T23 确认后按折线生成');
setForm('line');
check(els.rowAnchor.style.display === 'none', 'T23 回到直线形态隐藏锚点行');

// T24: 「记一点」按钮逐点记录(不依赖拖拽轮询)
timeOffset += 2000; currentSelection = []; tick(1); currentSelection = [mover]; tick(1);
setForm('trace');
els.btnTrace.onclick();                       // 开始, 记起点(1点)
trMark.Left = 250 - 7; trMark.Top = 170 - 7;  // 拖到第2处(不tick, 模拟拖拽中无轮询)
els.btnTracePoint.onclick();                  // 点按钮记录
trMark.Left = 350 - 7; trMark.Top = 230 - 7;
els.btnTracePoint.onclick();
trMark.Left = 460 - 7; trMark.Top = 300 - 7;
els.btnTracePoint.onclick();
check(els.tracePos.textContent.includes('4 点'), 'T24 记一点累计4点(含起点): ' + els.tracePos.textContent);
els.btnTracePoint.onclick();                  // 未移动再点 → 不重复
check(els.tracePos.textContent.includes('4 点'), 'T24 未移动不重复记点');
els.btnTraceDone.onclick();
calls.byPath = null;
els.btnGen.onclick();
check(calls.byPath && calls.byPath.pathStr.split('L').length >= 4, 'T24 生成多段折线(4点)');
setForm('line');

console.log('\n' + (failures === 0 ? '全部通过 ✔' : failures + ' 项失败 ✘'));
process.exit(failures === 0 ? 0 : 1);
