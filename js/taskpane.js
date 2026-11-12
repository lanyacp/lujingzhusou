/**
 * taskpane.js — 任务窗格逻辑: 状态机(选中图形 → 放置标记 → 拖B点 → 生成) + 轮询
 * 依赖: ../js/enums.js, ../js/core.js (由 taskpane.html 引入)
 */
(function () {
    'use strict';

    var E = ABP_ENUM;
    var $ = function (id) { return document.getElementById(id); };

    var state = {
        active: false,      // 流程进行中(已放置 A/B 标记)
        slideIdx: 0,        // mover 所在幻灯片索引
        moverId: null,      // 移动图形 Id
        moverName: '',
        moverRef: null,     // 移动图形引用缓存(优先使用, 失效再按 Id 重扫)
        slideRef: null,     // 幻灯片引用缓存
        auto: true,         // 选中新图形自动开始
        coolUntil: 0,       // 生成后的冷却(避免自动重启误触发)
        presKey: '',        // 流程所属演示文稿标识(切换文档时重置)
        postGenSelKey: null,// 门控: 选区未变化前不自动重启(生成后/异常重置后)
        lastCleanSeq: null, // 功能区「清理标记」信号(PluginStorage 计数)
        lastA: null,
        lastB: null,
        snapWatch: {},       // 吸附监看: 每个可拖标记的稳定性 {b:{x,y,n}, pc:..., pd:...}
        form: 'line',        // 路径形态: line|parabola|arc|waypoints|trace
        trace: null,         // 描迹状态 {active, pts[], lastMoveAt}
        lastPVsig: null      // 曲线预览签名(节流)
    };

    // ---------------- 小工具 ----------------

    function fmt(v) { return (Math.round(v * 100) / 100).toFixed(1); }
    function pt2cm(v) { return v * 2.54 / 72; }
    function escapeHtml(s) {
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function setStatus(msg, cls) {
        var el = $('status');
        el.textContent = msg;
        el.className = cls || '';
    }

    function setResult(html) {
        var el = $('result');
        el.innerHTML = html;
        el.style.display = 'block';
    }

    // ---------------- 流程控制 ----------------

    function activate(rawMover) {
        // 新版桥接防御: 选区引用几何不可靠 → 一律通过扫描定位取新鲜引用
        var loc = ABP.locateShape(rawMover);
        if (!loc) {
            state.postGenSelKey = ABP.shapeIdOf(rawMover);
            setStatus('⚠ 无法定位所选图形（坐标读取异常），流程已重置；请重新点选图形', 'warn');
            deactivate(false);
            return;
        }
        var mover = loc.shape;
        var slide = loc.slide;
        // 跨幻灯片切换: 先清理旧幻灯片上的流程标记, 避免残留孤儿标记
        if (state.active) {
            var oldPres = ABP.getPresentation();
            var oldSlide = oldPres ? ABP.slideByIndex(oldPres, state.slideIdx) : null;
            if (oldSlide && ABP.slideIndexOf(oldSlide) !== ABP.slideIndexOf(slide)) {
                ABP.cleanFlowMarkers(oldSlide);
            }
        }
        ABP.cleanFlowMarkers(slide);   // 清掉本页旧流程标记(保留中心十字)
        state.active = true;
        state.slideIdx = ABP.slideIndexOf(slide);
        state.moverId = ABP.shapeIdOf(mover);
        state.moverRef = mover;        // 缓存引用: 轮询优先使用, 免去每拍全页扫描
        state.slideRef = slide;
        var pres = ABP.getPresentation();
        try { state.presKey = (pres && pres.Name) || ''; } catch (e0) { state.presKey = ''; }
        try { state.moverName = mover.Name || '(未命名)'; } catch (e) { state.moverName = '(未命名)'; }
        var c = ABP.center(mover);
        if (!c) {
            state.postGenSelKey = ABP.shapeIdOf(mover);   // 选区未变化前不反复重试
            setStatus('⚠ 图形坐标读取异常，请重新点选图形', 'warn');
            deactivate(false);
            return;
        }
        state.lastA = null;
        state.lastB = null;
        state.lastPVsig = null;
        var refs0 = ABP.syncMarkers(slide, c, state.form);
        // B 标记仅直线形态需要(曲线形态用 D/P/M/W/T)
        var needB = (state.form === 'line') && !refs0.cdMode;
        if (refs0.err || !refs0.a || (needB && !refs0.b)) {
            state.postGenSelKey = ABP.shapeIdOf(mover);
            setStatus('⚠ ' + (refs0.err || '标记创建失败') + '，流程未开始', 'warn');
            deactivate(true);
            return;
        }
        ABP.updatePreviewLine(slide, c, c);
        renderMover();
        setStatus('● 蓝/红十字重叠在图形中心，按住蓝色十字拖到目标位置 B', 'ok');
    }

    function deactivate(clean) {
        cancelTrace();
        if (clean) {
            var pres = ABP.getPresentation();
            var slide = state.slideRef || (pres ? ABP.slideByIndex(pres, state.slideIdx) : null);
            if (slide) ABP.cleanFlowMarkers(slide);
        }
        state.active = false;
        state.moverId = null;
        state.moverName = '';
        state.moverRef = null;
        state.slideRef = null;
        state.lastA = null;
        state.lastB = null;
        renderMover();
    }

    /** 解析移动图形: 缓存引用须通过僵尸检测, 否则按 Id/名称新鲜扫描 */
    function resolveMover() {
        if (state.moverRef) {
            try {
                if (ABP.looksAlive(state.moverRef) && ABP.center(state.moverRef)) return state.moverRef;
            } catch (e0) { }
            state.moverRef = null;
        }
        var pres = ABP.getPresentation();
        var slide = resolveSlide(pres);
        if (!slide) return null;
        var m = ABP.refetchShapeFresh(pres, state.slideIdx, state.moverId, state.moverName) ||
                ABP.findShapeById(slide, state.moverId);
        if (m) { state.moverRef = m; state.slideRef = slide; }
        return m;
    }

    /** 解析流程幻灯片: 优先缓存引用 */
    function resolveSlide(pres) {
        if (state.slideRef) {
            try {
                if (ABP.slideIndexOf(state.slideRef) >= 1) return state.slideRef;
            } catch (e0) { }
            state.slideRef = null;
        }
        var s = ABP.slideByIndex(pres, state.slideIdx);
        if (s) state.slideRef = s;
        return s;
    }

    function renderMover() {
        var el = $('moverName');
        el.textContent = state.active ? state.moverName : '未选择';
        el.className = state.active ? 'mover on' : 'mover';
        $('step1').className = 'card' + (state.active ? ' live' : '');
        $('step2').className = 'card' + (state.active ? ' live' : '');
    }

    /** 按钮前置检查: 流程进行中才允许操作(醒目引导) */
    function ensureFlow(msg) {
        if (!state.active) {
            setStatus('⚠ ' + (msg || '请先选中图形开始流程'), 'warn');
            setResult('<div class="err">请先开始流程：<br>① 在幻灯片中点选一个图形<br>② 功能区「路径助手」→「创建路径」<br>看到图形中心出现红/蓝十字后，再使用标记点功能</div>');
            return false;
        }
        return true;
    }

    /**
     * 吸附监看: 标记位置连续 2 拍不变(刚停止拖动) 且 12pt 内有吸附目标 → 精确吸附
     * 目标 = 中心红十字 + 所有图形中心(排除移动图形自身与插件标记)
     * @return 吸附目标 {x,y,kind} 或 null
     */
    function watchSnap(key, ref, pos, slide) {
        try {
            if (!$('snapChk').checked) return null;
        } catch (e0) { }
        var w = state.snapWatch[key];
        if (!w) {
            w = state.snapWatch[key] = { x: null, y: null, n: 0 };
        }
        if (w.x !== null && Math.abs(w.x - pos.x) < 0.05 && Math.abs(w.y - pos.y) < 0.05) {
            w.n++;
            if (w.n === 2) {   // 刚刚停止移动
                var t = ABP.nearestSnapTarget(slide, pos, state.moverId);
                if (t && (Math.abs(pos.x - t.x) > 0.05 || Math.abs(pos.y - t.y) > 0.05)) {
                    if (ABP.snapMarkerTo(ref, t)) return t;
                }
            }
        } else {
            w.x = pos.x; w.y = pos.y; w.n = 1;
        }
        return null;
    }

    // ---------------- 路径形态引擎 ----------------

    /** 轮询节流控制 */
    var tickWrap = function () {
        try { tick(); } catch (e) {
            try { setStatus('轮询异常: ' + (e.message || e), 'warn'); } catch (e2) { }
        }
    };
    var mainTimer = null;
    function setTickRate(ms) {
        if (mainTimer !== null) clearInterval(mainTimer);
        mainTimer = setInterval(tickWrap, ms);
    }

    /** 锚点: C标记(若有)否则图形中心 */
    function anchorOf(c, refs) {
        return (refs.pc ? ABP.center(refs.pc) : c) || c;
    }

    /**
     * 按当前形态构建路线
     * @return {ok, need, from, to, pts, pathStr, endRel, drag[], desc}
     */
    function buildRoute(c, refs, size) {
        var f = state.form;
        var S = anchorOf(c, refs);
        if (f === 'line') {
            var cdMode = !!(refs.pc && refs.pd);
            var from, to, drag;
            if (cdMode) {
                from = ABP.center(refs.pc);
                to = ABP.center(refs.pd);
                drag = [['pc', refs.pc], ['pd', refs.pd]];
            } else {
                from = c;
                to = refs.b ? ABP.center(refs.b) : null;
                drag = refs.b ? [['b', refs.b]] : [];
            }
            if (!to) return { ok: false, need: '拖动蓝色十字到目标位置', drag: drag };
            return { ok: true, from: from, to: to, pts: [from, to], drag: drag, desc: cdMode ? '标记点 C→D' : '中心→B' };
        }
        if (f === 'parabola' || f === 'arc') {
            var Ep = refs.pd ? ABP.center(refs.pd) : null;
            var K = refs.pm ? ABP.center(refs.pm) : null;
            var drag2 = [];
            if (refs.pd) drag2.push(['pd', refs.pd]);
            if (refs.pm) drag2.push(['pm', refs.pm]);
            if (!Ep || !K) {
                var needP = (f === 'parabola') ? '「放顶点P」' : '「放弧点M」';
                return { ok: false, need: (Ep ? '还需要 ' + needP : '还需要 「放终点D」' + (K ? '' : ' 和 ' + needP)), drag: drag2 };
            }
            var b = (f === 'parabola') ? ABP.parabolaPath(S, K, Ep, size) : ABP.arcPath(S, K, Ep, size);
            return { ok: true, from: S, to: Ep, pts: b.pts, pathStr: b.path, endRel: ABP.pathEndpoint(b.path), drag: drag2, desc: (f === 'parabola' ? '抛物线' : '圆弧') };
        }
        if (f === 'waypoints') {
            var ws = refs.ws || [];
            var drag3 = [];
            var centers = [];
            for (var wi = 0; wi < ws.length; wi++) {
                var wc = ABP.center(ws[wi]);
                if (wc) { centers.push(wc); drag3.push(['w' + wi, ws[wi]]); }
            }
            if (!centers.length) return { ok: false, need: '点「+航点」添加航点并拖到位（至少 1 个）', drag: [] };
            var allPts = [S].concat(centers);
            var smooth = false;
            try { smooth = $('smoothW').checked; } catch (e0) { }
            var b2 = smooth ? ABP.smoothPath(S, allPts, size) : ABP.polylinePath(S, centers, size);
            return { ok: true, from: S, to: allPts[allPts.length - 1], pts: b2.pts, pathStr: b2.path, endRel: ABP.pathEndpoint(b2.path), drag: drag3, desc: smooth ? '平滑航点' : '折线航点' };
        }
        if (f === 'trace') {
            var tr = state.trace;
            if (!tr || !tr.pts || tr.pts.length < 2) {
                return { ok: false, need: '点「开始描迹」，拖动洋红标记描出路线', drag: [] };
            }
            var b3 = ABP.polylinePath(tr.pts[0], tr.pts.slice(1), size);
            return { ok: true, from: tr.pts[0], to: tr.pts[tr.pts.length - 1], pts: tr.pts.slice(), pathStr: b3.path, endRel: ABP.pathEndpoint(b3.path), drag: [], desc: '描迹(' + tr.pts.length + '点)' };
        }
        return { ok: false, need: '未知形态', drag: [] };
    }

    /** 形态切换: 清理他形态标记 + 行显隐 + 提示 */
    function switchForm(f) {
        if (state.form === f) return;
        cancelTrace(false);
        state.form = f;
        state.lastPVsig = null;
        state.lastA = null;
        state.lastB = null;
        var pres = ABP.getPresentation();
        var slide = resolveSlide(pres);
        if (slide) {
            if (f !== 'parabola' && f !== 'arc') ABP.deleteMarkersByPrefix(slide, '__pthlp_PM');
            ABP.deleteMarkersByPrefix(slide, '__pthlp_PV');
            ABP.deleteMarkersByPrefix(slide, '__pthlp_LINE');
        }
        var rows = { line: 'rowLine', parabola: 'rowPara', arc: 'rowArc', waypoints: 'rowWp', trace: 'rowTrace' };
        Object.keys(rows).forEach(function (k) {
            try { $(rows[k]).style.display = (k === f) ? '' : 'none'; } catch (e0) { }
        });
        var hints = {
            line: '拖动蓝色十字到目标位置。「放C/D」可切换为标记点对齐模式。',
            parabola: '抛物线：锚点(C或中心)为起点，「放终点D」「放顶点P」并拖到位，曲线精确经过 P。适合平抛/斜抛。',
            arc: '圆弧：锚点为起点，「放终点D」「放弧点M」，沿 S→M→D 的圆弧运动。',
            waypoints: '航点：「+航点」逐个放置并拖到位，图形按顺序经过；勾选「平滑」为圆滑曲线。',
            trace: '描迹：先手绘曲线做参照 → 点「开始描迹」→ 拖动出现的【洋红大十字】沿曲线走一遍 → 停住自动完成。'
        };
        try { $('formHint').textContent = hints[f] || ''; } catch (e1) { }
        if (f === 'trace') setStatus('● 描迹形态：点「开始描迹」按钮，会出现洋红大十字', 'ok');
    }

    /** 描迹引擎 */
    function startTrace() {
        // 流程未开始时: 若当前正好选中一个图形, 自动开始流程(减少一步操作断点)
        if (!state.active) {
            var sel0 = ABP.selectionInfo();
            var cands = [];
            for (var i0 = 0; i0 < sel0.shapes.length; i0++) {
                if (!ABP.isMarker(sel0.shapes[i0])) cands.push(sel0.shapes[i0]);
            }
            if (cands.length === 1) {
                activate(cands[0]);
            } else {
                ensureFlow('请先选中一个图形（或点「创建路径」）');
                return;
            }
        }
        var pres = ABP.getPresentation();
        var slide = resolveSlide(pres);
        var mover = resolveMover();
        var c = mover ? ABP.center(mover) : null;
        if (!slide || !c) return setStatus('⚠ 无法定位图形坐标', 'warn');
        var refs = ABP.scanAllMarkers(slide);
        var anchor = (refs.pc ? ABP.center(refs.pc) : c) || c;
        // 洋红标记比 C 点大一号并置顶, 避免与 C 重叠时看不见
        var t = ABP.placeNamedMarker(slide, '__pthlp_TR__', anchor.x, anchor.y, 9, E.RGB_MAGENTA);
        if (!t) return setStatus('⚠ 描迹标记创建失败', 'warn');
        state.trace = { active: true, pts: [{ x: anchor.x, y: anchor.y }], lastMoveAt: Date.now() };
        setTickRate(130);
        setStatus('● 描迹中：拖动【洋红大十字】沿参照曲线慢慢移动，停住约1.6秒自动完成', 'ok');
    }
    function finishTrace(auto) {
        if (!state.trace || !state.trace.active) return;
        state.trace.active = false;
        state.trace.pts = ABP.smoothPts(state.trace.pts);
        setTickRate(350);
        setStatus('✚ 描迹完成（' + state.trace.pts.length + ' 点），点「生成路径动画」', 'ok');
    }
    function cancelTrace() {
        if (state.trace && state.trace.active) {
            state.trace.active = false;
            setTickRate(350);
        }
        state.trace = null;
        var pres = ABP.getPresentation();
        var slide = resolveSlide(pres);
        if (slide) ABP.deleteMarkersByPrefix(slide, '__pthlp_TR__');
    }

    /** 当前选中的"其他图形"(非标记/非移动图形): 用于 D 点/航点直接落位 */
    function selectedOtherShape() {
        try {
            var sel = ABP.selectionInfo();
            for (var i = 0; i < sel.shapes.length; i++) {
                var s = sel.shapes[i];
                if (ABP.isMarker(s)) continue;
                var loc = ABP.locateShape(s);
                if (!loc) continue;
                if (ABP.shapeIdOf(loc.shape) === state.moverId) return null;   // 选中移动图形→用默认偏移
                return loc.shape;
            }
        } catch (e) { }
        return null;
    }

    /** 放置形态标记(D/P/M); D 在选中其他图形时直接放其中心 */
    function placeFormMarker(name, ox, oy, color) {
        if (!ensureFlow('请先开始流程')) return;
        var pres = ABP.getPresentation();
        var slide = resolveSlide(pres);
        var mover = resolveMover();
        var c = mover ? ABP.center(mover) : null;
        if (!slide || !c) return setStatus('⚠ 无法定位图形坐标', 'warn');
        var refs = ABP.scanAllMarkers(slide);
        var anchor = (refs.pc ? ABP.center(refs.pc) : c) || c;
        var tx = anchor.x + ox, ty = anchor.y + oy, atShape = false;
        if (name === '__pthlp_PD__') {
            var other = selectedOtherShape();
            var oc = other ? ABP.center(other) : null;
            if (oc) { tx = oc.x; ty = oc.y; atShape = true; }
        }
        var m = ABP.placeNamedMarker(slide, name, tx, ty, 7, color);
        var tag = name.indexOf('__pthlp_PM') === 0 ? (state.form === 'parabola' ? '顶点P(紫)' : '弧点M(紫)') : '终点D(橙)';
        setStatus(m ? ('✚ 已放置 ' + tag + (atShape ? ' 在所选图形中心（精确对齐），可再拖动微调' : '，拖到位置（拖近图形松手自动吸附中心）')) : '⚠ 标记创建失败', m ? 'ok' : 'warn');
    }

    /** 判断候选图形是否就是当前 mover(新版桥接下候选的 Id/名称可能失真, 需兜底定位比对) */
    function sameCandidate(cand) {
        try { if (ABP.shapeIdOf(cand) === state.moverId) return true; } catch (e0) { }
        try { if (state.moverName && cand.Name === state.moverName) return true; } catch (e1) { }
        try {
            var loc = ABP.locateShape(cand);
            return !!(loc && ABP.shapeIdOf(loc.shape) === state.moverId);
        } catch (e2) { return false; }
    }

    // ---------------- 轮询主循环 ----------------

    /** 读取功能区「清理标记」信号计数(跨页面经 PluginStorage 传递) */
    function readCleanSeq() {
        try {
            var app = ABP.APP();
            if (app && app.PluginStorage) {
                return parseInt(app.PluginStorage.getItem('abp_clean_seq') || '0', 10) || 0;
            }
        } catch (e) { }
        return 0;
    }

    function tick() {
        var pres = ABP.getPresentation();
        if (!pres) {
            setStatus('⚠ 未打开演示文稿，请先新建或打开一个文件', 'warn');
            if (state.active) deactivate(false);
            renderMover();
            return;
        }

        // 功能区「清理标记」联动: 流程立即停止, 避免标记被下一拍重建
        var cleanSeq = readCleanSeq();
        if (state.lastCleanSeq === null) state.lastCleanSeq = cleanSeq;
        else if (cleanSeq !== state.lastCleanSeq) {
            state.lastCleanSeq = cleanSeq;
            if (state.active) { state.active = false; state.moverId = null; renderMover(); }
            setStatus('已通过功能区清理辅助标记，流程已重置', 'ok');
            return;
        }

        // 切换演示文稿: 旧文档里的标记无法从这里清理, 仅重置状态
        if (state.active && state.presKey) {
            var pk = '';
            try { pk = pres.Name || ''; } catch (e0) { }
            if (pk && pk !== state.presKey) {
                setStatus('⚠ 已切换到其他演示文稿，流程已重置（旧文档残留标记可用「清理标记」移除）', 'warn');
                deactivate(false);
                renderMover();
                return;
            }
        }

        var sel = ABP.selectionInfo();
        var candidates = [];
        for (var i = 0; i < sel.shapes.length; i++) {
            if (!ABP.isMarker(sel.shapes[i])) candidates.push(sel.shapes[i]);
        }

        if (state.active) {
            var mover = resolveMover();
            var slide = resolveSlide(pres);
            if (!mover || !slide) {
                // 防闪: 重置后选区未变化前不自动重启, 避免activate/重置死循环
                state.postGenSelKey = state.moverId;
                setStatus('⚠ 移动图形已不存在（或无法读取），流程已重置；重新点选图形可再次开始', 'warn');
                deactivate(true);
                return;
            }
            var c = ABP.center(mover);
            var refs = (c && slide) ? ABP.syncMarkers(slide, c, state.form) : null;
            if (!c || !refs || refs.err) {
                state.postGenSelKey = state.moverId;
                setStatus('⚠ 无法读取图形坐标或创建标记，流程已重置；重新点选图形可再次开始', 'warn');
                deactivate(true);
                return;
            }
            // 中途切换: 另选了唯一的其他图形 → 以新图形重新开始(顺带清理旧页标记)
            // 但已放置 C/D/P/M/航点/描迹标记时锁定移动图形(点其他图形常是为了定位 D/做参照)
            var armed = !!(refs.pc || refs.pd || refs.pm || (refs.ws && refs.ws.length) || refs.tr);
            var wantSwitch = state.auto && Date.now() > state.coolUntil &&
                candidates.length === 1 && !sameCandidate(candidates[0]);
            if (wantSwitch && !armed) {
                activate(candidates[0]);
                return;
            }
            var size = ABP.slideSize(pres);

            // 描迹进行中: 记录轨迹 + 停住自动完成
            if (state.form === 'trace' && state.trace && state.trace.active) {
                var tp2 = refs.tr ? ABP.center(refs.tr) : null;
                if (tp2) {
                    var lastP = state.trace.pts[state.trace.pts.length - 1];
                    if (Math.sqrt(Math.pow(tp2.x - lastP.x, 2) + Math.pow(tp2.y - lastP.y, 2)) > 2.5) {
                        state.trace.pts.push(tp2);
                        state.trace.lastMoveAt = Date.now();
                    } else if (Date.now() - state.trace.lastMoveAt > 1600 && state.trace.pts.length >= 3) {
                        finishTrace(true);
                        return;
                    }
                }
            }

            var route = buildRoute(c, refs, size);

            // 吸附: 对可拖标记做"刚停止移动"检测 (描迹中不吸附)
            var snapped = null;
            for (var di = 0; di < (route.drag || []).length; di++) {
                var dref = route.drag[di][1];
                if (!dref) continue;
                var dpos = ABP.center(dref);
                if (!dpos) continue;
                var st = watchSnap(route.drag[di][0], dref, dpos, slide);
                if (st) snapped = st;
            }
            if (snapped) {
                route = buildRoute(c, refs, size);
                state.lastB = null;
                state.lastPVsig = null;
            }

            $('aPos').textContent = '(' + fmt(c.x) + ', ' + fmt(c.y) + ') pt';
            if (!route.ok) {
                $('bPos').textContent = '—';
                $('cdPos').textContent = '⚠ ' + route.need;
                $('delta').textContent = '—';
                $('distNorm').textContent = '—';
                setStatus(wantSwitch && armed ? '🔒 已放置标记：移动图形已锁定（点其他图形不会切换）。换图形请先「清除标记点」'
                    : '● ' + route.need, 'ok');
                return renderMover();
            }

            var dx = route.to.x - route.from.x, dy = route.to.y - route.from.y;
            var len = 0;
            for (var li = 1; li < route.pts.length; li++) {
                len += Math.sqrt(Math.pow(route.pts[li].x - route.pts[li - 1].x, 2) +
                    Math.pow(route.pts[li].y - route.pts[li - 1].y, 2));
            }
            $('bPos').textContent = '(' + fmt(route.to.x) + ', ' + fmt(route.to.y) + ') pt';
            $('cdPos').textContent = route.desc + (state.form === 'line' ? '' :
                (' · 锚点=' + (refs.pc ? 'C点' : '中心')));
            $('delta').textContent = '直线Δ(' + fmt(dx) + ',' + fmt(dy) + ')pt · 路线长' + fmt(len) + 'pt(' + fmt(pt2cm(len)) + 'cm)';
            $('distNorm').textContent = '终点 Δx=' + (dx / size.w).toFixed(4) + '  Δy=' + (dy / size.h).toFixed(4);

            // 预览: 直线=单虚线; 曲线=折线段(签名节流)
            if (state.form === 'line') {
                var la = state.lastA, lb = state.lastB;
                if (!la || !lb ||
                    Math.abs(la.x - route.from.x) > 0.3 || Math.abs(la.y - route.from.y) > 0.3 ||
                    Math.abs(lb.x - route.to.x) > 0.3 || Math.abs(lb.y - route.to.y) > 0.3) {
                    ABP.updatePreviewLine(slide, route.from, route.to);
                    state.lastA = route.from;
                    state.lastB = route.to;
                }
            } else {
                var sig = '';
                for (var si = 0; si < route.pts.length; si++) {
                    sig += Math.round(route.pts[si].x) + ',' + Math.round(route.pts[si].y) + ';';
                }
                if (sig !== state.lastPVsig) {
                    ABP.deleteMarkersByPrefix(slide, '__pthlp_LINE');
                    ABP.updateRoutePreview(slide, route.pts);
                    state.lastPVsig = sig;
                }
            }

            // 形态信息行
            try {
                if (state.form === 'parabola' && refs.pm) {
                    var pp = ABP.center(refs.pm);
                    if (pp) $('paraPos').textContent = 'P(' + fmt(pp.x) + ',' + fmt(pp.y) + ')';
                }
                if (state.form === 'arc' && refs.pm) {
                    var mp = ABP.center(refs.pm);
                    if (mp) $('arcPos').textContent = 'M(' + fmt(mp.x) + ',' + fmt(mp.y) + ')';
                }
                if (state.form === 'waypoints') {
                    $('wpPos').textContent = ((refs.ws || []).length) + ' 个航点';
                }
                if (state.form === 'trace') {
                    $('tracePos').textContent = (state.trace && state.trace.active) ?
                        ('描迹中 ' + state.trace.pts.length + ' 点') :
                        (((state.trace && state.trace.pts) || []).length + ' 点已记录');
                }
            } catch (eInfo) { }

            setStatus(snapped ? ('✦ 已自动吸附到' + (snapped.kind || '') + '中心（微调时可取消勾选吸附）') :
                (wantSwitch && armed ? '🔒 已放置标记：移动图形已锁定（点其他图形不会切换）。换图形请先「清除标记点」' :
                    '● ' + route.desc + ' 就绪，点「生成路径动画」'), 'ok');
        } else if (state.auto && Date.now() > state.coolUntil) {
            if (candidates.length === 1) {
                var candId = ABP.shapeIdOf(candidates[0]);
                if (state.postGenSelKey !== null && candId === state.postGenSelKey) {
                    // 选区仍是生成时选中的图形(生成后会自动选中它), 不自动重启;
                    // 一旦选区变化(选中其他图形/空选)即解除门控, 之后重新点选可正常开始
                } else {
                    state.postGenSelKey = null;
                    activate(candidates[0]);
                }
            } else if (state.postGenSelKey !== null) {
                state.postGenSelKey = null;   // 空选/多选 → 解除门控
            }
        }
        renderMover();
    }

    // ---------------- 生成路径动画 ----------------

    function generate() {
        if (!state.active) {
            setResult('<div class="err">请先在幻灯片中选中一个图形（步骤①）</div>');
            return;
        }
        var pres = ABP.getPresentation();
        if (!pres) { setResult('<div class="err">演示文稿已关闭</div>'); return; }
        var mover = resolveMover();
        var slide = resolveSlide(pres);
        if (!mover || !slide) { setResult('<div class="err">找不到移动图形，请重新选择</div>'); deactivate(true); return; }

        var c = ABP.center(mover);
        if (!c) { setResult('<div class="err">无法读取图形坐标，请重新选择图形</div>'); return; }
        var refs = ABP.syncMarkers(slide, c, state.form);
        if (!refs || refs.err) {
            setResult('<div class="err">' + ((refs && refs.err) || '标记不可用') + '</div>');
            return;
        }
        var size = ABP.slideSize(pres);
        var route = buildRoute(c, refs, size);
        if (!route.ok) {
            setResult('<div class="err">' + escapeHtml(route.need || '路线未就绪') + '</div>');
            return;
        }

        // 路线总长(零位移确认)
        var len = 0;
        for (var li = 1; li < route.pts.length; li++) {
            len += Math.sqrt(Math.pow(route.pts[li].x - route.pts[li - 1].x, 2) +
                Math.pow(route.pts[li].y - route.pts[li - 1].y, 2));
        }
        if (len < 0.5) {
            var go = false;
            try {
                go = window.confirm('路线长度几乎为零。\n确定生成"原地不动"的路径动画吗？');
            } catch (e) { go = true; }
            if (!go) return;
        }

        // 路径串: 直线用 pathString, 曲线用构建器产物
        var pathStr, endRel;
        if (state.form === 'line') {
            var dx = (route.to.x - route.from.x) / size.w;
            var dy = (route.to.y - route.from.y) / size.h;
            pathStr = ABP.pathString(dx, dy);
            endRel = { x: dx, y: dy };
        } else {
            pathStr = route.pathStr;
            endRel = route.endRel;
        }
        if (!pathStr || !endRel) {
            setResult('<div class="err">路径构建失败，请检查标记是否齐全</div>');
            return;
        }

        var opts = {
            duration: parseFloat($('dur').value) || 2,
            trigger: parseInt($('trig').value, 10) || E.msoAnimTriggerOnPageClick,
            smooth: $('smooth').checked
        };

        var replaced = 0;
        if ($('replaceAnim').checked) replaced = ABP.deleteShapeAnimations(slide, mover);

        // 图形定位: 优先活引用(新版 WPS 的 Id 会漂移), Id 扫描兜底
        var r = ABP.createPathEffectByPath(pres, state.slideIdx, state.moverId, pathStr, endRel, opts, mover);

        if (!$('keepMarkers').checked) ABP.cleanFlowMarkers(slide);
        state.coolUntil = Date.now() + 1500;
        state.postGenSelKey = ABP.shapeIdOf(mover);   // 选区未变化前不自动重启流程
        state.active = false;
        state.moverId = null;
        if (state.trace && state.trace.active) finishTrace(false);
        renderMover();

        if (r.ok) {
            var rbTxt = (r.readback && r.readback.raw) ? escapeHtml(r.readback.raw) : '';
            setResult('<div class="ok">✓ 已生成路径动画（' + escapeHtml(route.desc) + '，策略' + r.strategy + '）' +
                (replaced ? '，已替换 ' + replaced + ' 个旧动画' : '') + '<br>' +
                '终点: Δx=' + endRel.x.toFixed(4) + ', Δy=' + endRel.y.toFixed(4) + '（幻灯片百分比）<br>' +
                (rbTxt ? '回读: ' + rbTxt + '<br>' : '') +
                '提示: 选中图形即可看到路径 —— 绿点=起点, 红点=终点</div>');
            try { mover.Select(); } catch (e) { }
        } else {
            setResult('<div class="err">✗ 生成失败:<br>' +
                r.msg.map(escapeHtml).join('<br>') + '<br>请点「API 诊断」排查</div>');
        }
    }

    // ---------------- 工具按钮 ----------------

    function bindTools() {
        // 逐按钮容错绑定: 单个元素缺失不影响其余按钮
        function bind(id, fn) {
            try {
                var el = $(id);
                if (el) el.onclick = fn;
            } catch (e0) { }
        }
        bind('btnGen', generate);

        // 路径形态
        try {
            var pf = $('pathForm');
            if (pf) pf.onchange = function () { switchForm(this.value); };
        } catch (eF) { }
        bind('btnParaD', function () { placeFormMarker('__pthlp_PD__', 48, 36, E.RGB_ORANGE); });
        bind('btnParaP', function () { placeFormMarker('__pthlp_PM__', 24, -70, E.RGB_PURPLE); });
        bind('btnArcD', function () { placeFormMarker('__pthlp_PD__', 48, 36, E.RGB_ORANGE); });
        bind('btnArcM', function () { placeFormMarker('__pthlp_PM__', 24, -50, E.RGB_PURPLE); });
        bind('btnAddW', function () {
            if (!ensureFlow('请先开始流程')) return;
            var pres = ABP.getPresentation();
            var slide = resolveSlide(pres);
            var mover = resolveMover();
            var c = mover ? ABP.center(mover) : null;
            if (!slide || !c) return setStatus('⚠ 无法定位图形坐标', 'warn');
            var refs = ABP.scanAllMarkers(slide);
            var anchor = (refs.pc ? ABP.center(refs.pc) : c) || c;
            var n = (refs.ws || []).length + 1;
            // 选中其他图形时, 航点直接落其中心(快速规划"经过哪些图形")
            var other = selectedOtherShape();
            var oc = other ? ABP.center(other) : null;
            var wx = oc ? oc.x : anchor.x + 30 + 22 * (n - 1);
            var wy = oc ? oc.y : anchor.y + 28;
            var m = ABP.placeNamedMarker(slide, '__pthlp_W' + n + '__', wx, wy, 7, E.RGB_CYAN);
            setStatus(m ? ('✚ 已放置航点 W' + n + (oc ? ' 在所选图形中心' : '(青)，拖到位') + '；可继续「+航点」') : '⚠ 航点创建失败', m ? 'ok' : 'warn');
        });
        bind('btnClearW', function () {
            var pres = ABP.getPresentation();
            var slide = resolveSlide(pres);
            if (!slide) return setStatus('⚠ 未打开演示文稿', 'warn');
            var n = ABP.deleteMarkersByPrefix(slide, '__pthlp_W');
            setStatus(n ? '已清空 ' + n + ' 个航点' : '没有航点', 'ok');
        });
        bind('btnTrace', startTrace);
        bind('btnTraceDone', function () { finishTrace(false); });

        // 标记点 C/D
        bind('btnPlaceC', function () {
            if (!ensureFlow('请先在流程中（选中图形并点创建路径）')) return;
            var pres = ABP.getPresentation();
            var slide = resolveSlide(pres);
            var mover = resolveMover();
            var c = mover ? ABP.center(mover) : null;
            if (!slide || !c) return setStatus('⚠ 无法定位图形坐标', 'warn');
            var m = ABP.placePointMarker(slide, 'C', c);
            setStatus(m ? '✚ C 点(绿十字)已放在图形左上方，拖动它到参照位置（如顶点）' : '⚠ C 点放置失败', m ? 'ok' : 'warn');
        });
        bind('btnPlaceD', function () {
            if (!ensureFlow('请先在流程中（选中图形并点创建路径）')) return;
            var pres = ABP.getPresentation();
            var slide = resolveSlide(pres);
            var mover = resolveMover();
            var c = mover ? ABP.center(mover) : null;
            if (!slide || !c) return setStatus('⚠ 无法定位图形坐标', 'warn');
            // 选中其他图形时, D 直接落其中心(先点目标图形再点放D点即可)
            var other = selectedOtherShape();
            var oc = other ? ABP.center(other) : null;
            var m = oc ? ABP.placeNamedMarker(slide, '__pthlp_PD__', oc.x, oc.y, 7, E.RGB_ORANGE)
                       : ABP.placePointMarker(slide, 'D', c);
            setStatus(m ? ('✚ D 点(橙十字)已放' + (oc ? '在所选图形中心（精确对齐），可再拖动微调' : '在图形右下方，拖动它到目标位置（拖近图形/红十字自动吸附）')) : '⚠ D 点放置失败', m ? 'ok' : 'warn');
        });
        bind('btnClearCD', function () {
            var pres = ABP.getPresentation();
            var slide = resolveSlide(pres);
            if (!slide) return setStatus('⚠ 未打开演示文稿', 'warn');
            var n = ABP.clearPointMarkers(slide);
            setStatus(n ? '已清除标记点，恢复中心→B 模式' : '没有已放置的标记点', 'ok');
        });

        bind('btnCross', function () {
            var pres = ABP.getPresentation();
            if (!pres) return setStatus('⚠ 未打开演示文稿', 'warn');
            var sel = ABP.selectionInfo();
            var shapes = [];
            for (var i = 0; i < sel.shapes.length; i++) {
                if (!ABP.isMarker(sel.shapes[i])) shapes.push(sel.shapes[i]);
            }
            if (!shapes.length) return setStatus('请先选中一个或多个图形', 'warn');
            var slide = ABP.parentSlide(shapes[0]) || ABP.getActiveSlide();
            if (!slide) return setStatus('⚠ 无法定位幻灯片', 'warn');
            var on = 0, off = 0, bad = 0;
            for (var j = 0; j < shapes.length; j++) {
                var r = ABP.toggleCenterCross(slide, shapes[j]);
                if (r === true) on++;
                else if (r === false) off++;
                else bad++;
            }
            if (bad) setStatus('⚠ ' + bad + ' 个图形坐标读取失败，未加中心十字', 'warn');
            else setStatus(off ? ('已移除 ' + off + ' 个中心十字') :
                ('已为 ' + on + ' 个图形添加中心十字（静态标记，图形移动后请重新点击）'), 'ok');
        });

        bind('btnCleanSlide', function () {
            var slide = ABP.getActiveSlide();
            if (!slide) return setStatus('⚠ 未打开演示文稿', 'warn');
            var n = ABP.cleanSlideMarkers(slide);
            if (state.active) { state.active = false; state.moverId = null; renderMover(); }
            setStatus('已清理本页 ' + n + ' 个辅助标记', 'ok');
        });

        bind('btnDelAnim', function () {
            var sel = ABP.selectionInfo();
            var shapes = [];
            for (var i = 0; i < sel.shapes.length; i++) {
                if (!ABP.isMarker(sel.shapes[i])) shapes.push(sel.shapes[i]);
            }
            if (shapes.length !== 1) return setStatus('请先选中一个图形', 'warn');
            var slide = ABP.parentSlide(shapes[0]) || ABP.getActiveSlide();
            if (!slide) return setStatus('⚠ 无法定位幻灯片', 'warn');
            var n = ABP.deleteShapeAnimations(slide, shapes[0]);
            setStatus(n >= 0 ? ('已删除该图形的 ' + n + ' 个动画') : '删除失败（无法访问动画时间线）', n >= 0 ? 'ok' : 'warn');
        });

        bind('btnDiag', function () {
            var steps = ABP.diagnose();
            var html = steps.map(function (s) {
                return '<div class="' + (s.ok ? 'okline' : 'errline') + '">' + (s.ok ? '✓' : '✗') +
                    ' ' + escapeHtml(s.name) +
                    (s.detail ? ' — ' + escapeHtml(s.detail) : '') +
                    (s.error ? ' — <i>' + escapeHtml(s.error) + '</i>' : '') + '</div>';
            }).join('');
            setResult('<div class="diag"><b>API 诊断报告</b><br>' + html + '</div>');
        });

        bind('autoChk', function () { state.auto = this.checked; });
    }

    // ---------------- 启动 ----------------

    // 展示桥接自测报告(selftest 写入 PluginStorage, 诊断期功能)
    try {
        var psRep = ABP.APP().PluginStorage.getItem('abp_st_report');
        if (psRep) {
            setResult('<div class="diag"><b>桥接自测报告（自动生成）</b><br>' +
                escapeHtml(String(psRep)).replace(/,"/g, ',<br>"') + '</div>');
        }
    } catch (e0) { }

    bindTools();
    renderMover();
    tick();
    setTickRate(350);   // 描迹时加速到 130ms
})();
