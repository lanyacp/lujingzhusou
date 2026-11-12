/**
 * core.js — 「路径助手」核心逻辑
 *
 * 坐标体系（单位: 磅 pt）:
 *   图形中心 A = (Left + Width/2, Top + Height/2)
 *   路径动画位移(归一化) dx = (B.x - A.x)/幻灯片宽, dy = (B.y - A.y)/幻灯片高
 *   MotionEffect 的 FromX/ToX/Path 坐标: 原点=图形中心, 1.0=一个幻灯片宽/高
 *
 * 辅助形状命名前缀 __pthlp_（A 十字 / B 标记 / 预览线 / 中心十字 C_xxx）
 */
var ABP = (function () {
    'use strict';

    var E = ABP_ENUM;
    var MARK = '__pthlp_';
    var NAME_A = '__pthlp_A__';      // A 点十字(红)
    var NAME_B = '__pthlp_B__';      // B 点标记(蓝, 可拖拽)
    var NAME_LINE = '__pthlp_LINE__';// 预览虚线
    var NAME_PC = '__pthlp_PC__';    // 标记点C(绿, 移动图形上的参照点/锚点)
    var NAME_PD = '__pthlp_PD__';    // 标记点D(橙, 终点)
    var NAME_PM = '__pthlp_PM__';    // 顶点P/弧点M(紫)
    var NAME_TR = '__pthlp_TR__';    // 描迹标记(洋红)
    var NAME_WP = '__pthlp_W';       // 航点前缀: __pthlp_W3__
    var NAME_PV = '__pthlp_PV';      // 曲线预览线段前缀

    // ---------------- 上下文 ----------------

    function APP() {
        try {
            if (window.Application) return window.Application;
        } catch (e) { }
        try {
            if (typeof wps !== 'undefined' && wps && wps.WppApplication) return wps.WppApplication();
        } catch (e) { }
        return null;
    }

    function getPresentation() {
        var app = APP();
        if (!app) return null;
        try {
            var p = app.ActivePresentation;
            return p || null;
        } catch (e) { return null; }
    }

    /** 沿 Parent 链找到图形所在的 Slide (逐级容错: 单级属性异常不中断整链) */
    function parentSlide(shp) {
        var o = shp;
        var guard = 0;
        while (o && guard++ < 8) {
            var par = null;
            try { par = o.Parent; } catch (e0) { return null; }
            if (!par) return null;
            var isSlide = false;
            try {
                isSlide = (typeof par.SlideIndex === 'number') && (typeof par.Shapes !== 'undefined');
            } catch (e1) { isSlide = false; }
            if (isSlide) return par;
            o = par;
        }
        return null;
    }

    /** 判断对象是否"像幻灯片"(兼容 SlideIndex 为 number/string 的桥接实现) */
    function isSlideLike(o) {
        try {
            return !!o && (typeof o.SlideIndex === 'number' || typeof o.SlideIndex === 'string') &&
                typeof o.Shapes !== 'undefined';
        } catch (e) { return false; }
    }

    /**
     * 当前(活动)幻灯片: 多重降级
     * 新版桥接实测: 选区图形的 Parent 链不可用, 需优先 View.Slide / ShapeRange.Parent / SlideRange
     */
    function getActiveSlide() {
        var app = APP();
        if (!app) return null;
        try { var v = app.ActiveWindow.View.Slide; if (isSlideLike(v)) return v; } catch (e0) { }
        try {
            var sel = app.ActiveWindow.Selection;
            if (sel.Type === E.ppSelectionSlides) {
                var s0 = sel.SlideRange.Item(1);
                if (isSlideLike(s0)) return s0;
            }
            if (sel.Type === E.ppSelectionShapes && sel.ShapeRange.Count > 0) {
                try { var p = sel.ShapeRange.Parent; if (isSlideLike(p)) return p; } catch (e1) { }
                var s1 = parentSlide(sel.ShapeRange.Item(1));
                if (isSlideLike(s1)) return s1;
            }
        } catch (e2) { }
        try {
            var p2 = getPresentation();
            if (p2 && p2.SlideIndex) return p2.Slides.Item(p2.SlideIndex);
        } catch (e3) { }
        return null;
    }

    function slideIndexOf(slide) {
        try { return slide.SlideIndex; } catch (e) { return -1; }
    }

    function slideByIndex(pres, idx) {
        if (!pres || !idx) return null;
        try { return pres.Slides.Item(idx); } catch (e) { return null; }
    }

    function slideSize(pres) {
        try {
            return { w: pres.PageSetup.SlideWidth, h: pres.PageSetup.SlideHeight, ok: true };
        } catch (e) {
            return { w: 960, h: 540, ok: false };
        }
    }

    // ---------------- 图形 / 选择 ----------------

    function shapeIdOf(shp) {
        try {
            var id = shp.Id;
            if (typeof id === 'number') return id;
        } catch (e) { }
        try { return 'n' + (shp.Name || '').replace(/[^\w]/g, ''); } catch (e) { return 'n?'; }
    }

    /**
     * 子形状提升为顶层图形（动画目标必须是顶层图形）
     * 新版桥接实测: 顶层图形的 ParentGroup 会返回垃圾包装(Name=default, 坐标1,1),
     * 必须验证父对象确实是组合图形(Type=6)才上溯
     */
    function resolveTop(shp) {
        var guard = 0;
        while (shp && guard++ < 10) {
            var p = null;
            try { p = shp.ParentGroup; } catch (e0) { break; }
            var isGroup = false;
            try { isGroup = !!p && p.Type === E.msoShapeTypeGroup; } catch (e1) { }
            if (isGroup) shp = p; else break;
        }
        return shp;
    }

    /** 判断是否本插件辅助标记: 名称前缀 / 备注文本(第二通道) / 结构特征(≤20pt十字) */
    function isMarker(shp) {
        try { if ((shp.Name || '').indexOf(MARK) === 0) return true; } catch (e0) { }
        try { if ((shp.AlternativeText || '').indexOf(MARK) === 0) return true; } catch (e1) { }
        try {
            if (shp.Type === E.msoShapeTypeAutoShape &&
                shp.AutoShapeType === E.msoShapeCross &&
                shp.Width <= 20 && shp.Height <= 20) return true;
        } catch (e2) { }
        return false;
    }

    function indexOfShape(arr, shp) {
        var id = shapeIdOf(shp);
        for (var i = 0; i < arr.length; i++) {
            if (shapeIdOf(arr[i]) === id) return i;
        }
        return -1;
    }

    /**
     * 在幻灯片 Shapes 集合中按 名称(优先)/Id 查找新鲜引用
     * 新版桥接: 选区图形的几何(Left/Top=0)与 Id 不可靠, 名称相对可靠
     */
    function findBy(slide, id, nm) {
        if (!slide) return null;
        try {
            var ss = slide.Shapes;
            var n = ss.Count;
            var byId = null, byName = null;
            for (var i = 1; i <= n; i++) {
                var c = null;
                try { c = ss.Item(i); } catch (e0) { continue; }
                var cn = '';
                try { cn = c.Name || ''; } catch (e1) { }
                if (nm && cn && cn === nm && looksAlive(c)) byName = c;
                if (id !== null && id !== undefined && shapeIdOf(c) === id && looksAlive(c)) byId = c;
                if (byName && byId) break;
            }
            return byName || byId || null;
        } catch (e) { }
        return null;
    }

    /** 宽高唯一匹配兜底: 半僵尸引用的 W/H 是真实值, 若页面上尺寸唯一即可定位 */
    function findByGeom(slide, w, h) {
        if (!slide || !isFinite(w) || !isFinite(h) || w <= 0 || h <= 0) return null;
        try {
            var ss = slide.Shapes;
            var n = ss.Count;
            var hit = null, hits = 0;
            for (var i = 1; i <= n; i++) {
                var c = null;
                try { c = ss.Item(i); } catch (e0) { continue; }
                try {
                    if (Math.abs(c.Width - w) < 0.5 && Math.abs(c.Height - h) < 0.5 && looksAlive(c)) {
                        hit = c; hits++;
                        if (hits > 1) return null;   // 尺寸不唯一, 放弃
                    }
                } catch (e1) { }
            }
            return hits === 1 ? hit : null;
        } catch (e2) { }
        return null;
    }

    /** 用幻灯片扫描的新鲜引用替换选区给的引用(几何/API调用都必须用新鲜引用) */
    function refreshRef(raw, slide) {
        var nm = '';
        try { nm = raw.Name || ''; } catch (e0) { }
        var id = shapeIdOf(raw);
        var s = findBy(slide, id, nm);
        if (s) return s;
        // 身份匹配失败 → 宽高唯一匹配兜底(半僵尸的 W/H 是真的)
        try { s = findByGeom(slide, raw.Width, raw.Height); } catch (e1) { }
        return s || raw;
    }

    /**
     * 图形终极定位(新版桥接专用策略):
     * 选区引用只提供身份线索(名称/Id/宽高), 几何与 API 调用一律用扫描引用。
     * 依次: 活动幻灯片(多重降级)按名/Id/宽高找 → 全文稿逐页扫描 → 健康原引用。
     * @return {slide, shape} 或 null
     */
    function locateShape(raw) {
        if (!raw) return null;
        var nm = '';
        try { nm = raw.Name || ''; } catch (e0) { }
        var id = shapeIdOf(raw);
        var w = NaN, h = NaN;
        try { w = raw.Width; h = raw.Height; } catch (e1) { }
        // 1) 活动幻灯片
        var slide = getActiveSlide();
        if (slide) {
            var s = findBy(slide, id, nm) || findByGeom(slide, w, h);
            if (s) return { slide: slide, shape: s };
        }
        // 2) 全文稿逐页扫描(页面数通常很少)
        var pres = getPresentation();
        if (pres) {
            try {
                var n = pres.Slides.Count;
                for (var i = 1; i <= n; i++) {
                    var sl = pres.Slides.Item(i);
                    var s2 = findBy(sl, id, nm) || findByGeom(sl, w, h);
                    if (s2) return { slide: sl, shape: s2 };
                }
            } catch (e2) { }
        }
        // 3) 原引用本身健康(老版本路径)
        if (looksAlive(raw)) {
            var sl3 = parentSlide(raw) || getActiveSlide();
            if (sl3) return { slide: sl3, shape: raw };
        }
        return null;
    }

    /** 读取当前选区, 返回 {ok,type,shapes[](去重后的顶层图形,已换成新鲜引用),msg} */
    function selectionInfo() {
        var res = { ok: false, type: -1, shapes: [], msg: '' };
        var app = APP();
        if (!app) { res.msg = '无法获取 WPS 演示 Application'; return res; }
        try {
            var sel = app.ActiveWindow.Selection;
            res.type = sel.Type;
            res.ok = true;
            if (sel.Type === E.ppSelectionShapes) {
                var range = sel.ShapeRange;
                var raws = [];
                for (var i = 1; i <= range.Count; i++) {
                    var top = resolveTop(range.Item(i));
                    if (indexOfShape(raws, top) < 0) raws.push(top);
                }
                // 新版桥接防御: 选区引用只取身份/尺寸, 几何与 API 调用用幻灯片扫描的新鲜引用
                var slide = getActiveSlide();
                for (var k = 0; k < raws.length; k++) {
                    res.shapes.push(refreshRef(raws[k], slide));
                }
            }
        } catch (e) {
            res.ok = false;
            res.msg = '读取选区失败: ' + (e.message || e);
        }
        return res;
    }

    function findShapeById(slide, id) {
        if (!slide) return null;
        var stats = { scanned: 0, unreadable: 0 };
        try {
            var ss = slide.Shapes;
            var n = ss.Count;
            for (var i = 1; i <= n; i++) {
                var s = null;
                try {
                    s = ss.Item(i);
                    stats.scanned++;
                } catch (e0) {
                    stats.unreadable++;
                    continue;   // 单个图形读取失败不中断扫描
                }
                if (shapeIdOf(s) === id) return s;
            }
        } catch (e) { }
        findShapeById.lastStats = stats;
        return null;
    }

    /** 图形中心; 读取失败返回 null (调用方必须判空, 严禁默认(0,0)) */
    function center(shp) {
        try {
            var x = shp.Left + shp.Width / 2;
            var y = shp.Top + shp.Height / 2;
            if (isFinite(x) && isFinite(y)) return { x: x, y: y };
        } catch (e) { }
        return null;
    }

    /**
     * 僵尸引用检测: 新版桥接里失效引用不抛错, 而是所有坐标返回 0。
     * 正常图形不可能 Left/Top/Width/Height 全为 0 (0×0 大小无意义), 全 0 即判失效。
     */
    function looksAlive(shp) {
        try {
            var l = shp.Left, t = shp.Top, w = shp.Width, h = shp.Height;
            if (typeof l !== 'number' || typeof w !== 'number') return false;
            if (!isFinite(l) || !isFinite(t) || !isFinite(w) || !isFinite(h)) return false;
            if (l === 0 && t === 0 && w === 0 && h === 0) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    /** 按 Id + 名称在幻灯片上取新鲜引用(都校验 looksAlive) */
    function refetchShapeFresh(pres, slideIdx, shapeId, shapeName) {
        var slide = null;
        try { slide = pres.Slides.Item(slideIdx); } catch (e0) { return null; }
        return findBy(slide, shapeId, shapeName) ||
               (function () { var s = findShapeById(slide, shapeId); return (s && looksAlive(s)) ? s : null; })();
    }

    // ---------------- 辅助标记 ----------------

    function findMarker(slide, name) {
        if (!slide) return null;
        try {
            var ss = slide.Shapes;
            var n = ss.Count;
            for (var i = 1; i <= n; i++) {
                var s = null;
                try { s = ss.Item(i); } catch (e0) { continue; }
                try { if (s.Name === name) return s; } catch (e1) { }
            }
        } catch (e) { }
        return null;
    }

    /** 给标记形状打标识; 返回名称通道是否成功(失败则无法按名寻址, 调用方应放弃该形状) */
    function tagMarker(shp, name) {
        var named = false;
        try { shp.Name = name; } catch (e0) { }
        try { named = (shp.Name === name); } catch (e1) { }
        try { shp.AlternativeText = MARK + name; } catch (e2) { }   // 第二通道
        return named;
    }

    function addCross(slide, name, cx, cy, size, rgb) {
        var shp = slide.Shapes.AddShape(E.msoShapeCross, cx - size / 2, cy - size / 2, size, size);
        if (!tagMarker(shp, name)) {
            // 命名失败将导致下一拍找不到标记而无限重建 → 立即删除并报错
            try { shp.Delete(); } catch (e0) { }
            return null;
        }
        try { shp.Line.Visible = E.msoFalse; } catch (e) { }
        try {
            shp.Fill.Visible = E.msoTrue;
            shp.Fill.ForeColor.RGB = rgb;
            shp.Fill.Transparency = 0.1;
        } catch (e) { }
        try { shp.ZOrder(E.msoBringToFront); } catch (e) { }
        return shp;
    }

    /** A 十字: 不存在则创建; 存在则移到 c (图形移动后跟随); 失败返回 null */
    function ensureAMarker(slide, c) {
        var a = findMarker(slide, NAME_A);
        if (!a) return addCross(slide, NAME_A, c.x, c.y, 7, E.RGB_RED);
        try {
            a.Left = c.x - a.Width / 2;
            a.Top = c.y - a.Height / 2;
        } catch (e) { }
        return a;
    }

    /** B 标记: 不存在(被误删)则在 c 处重生; 失败返回 null */
    function ensureBMarker(slide, c) {
        var b = findMarker(slide, NAME_B);
        if (!b) return addCross(slide, NAME_B, c.x, c.y, 8, E.RGB_BLUE);
        return b;
    }

    /**
     * 单遍扫描同步全部流程标记 (A/B/C/D/P/W/T)
     * - A 始终对准 c (图形移动后跟随)
     * - form='line' 且 C、D 同时存在 → 标记点模式: 移除 B
     * - form='line' 且非标记点模式 → B 缺失则重生; 其他形态 → 移除 B
     * @return {a,b,pc,pd,pm,tr,ws,cdMode,err}
     */
    function syncMarkers(slide, c, form) {
        form = form || 'line';
        if (!slide) return { a: null, b: null, pc: null, pd: null, pm: null, tr: null, ws: [], cdMode: false, err: '无幻灯片' };
        var m = scanAllMarkers(slide);
        var a = m.a;
        if (!a) a = addCross(slide, NAME_A, c.x, c.y, 7, E.RGB_RED);
        else {
            try {
                a.Left = c.x - a.Width / 2;
                a.Top = c.y - a.Height / 2;
            } catch (e1) { }
        }
        var cdMode = !!(m.pc && m.pd);
        var b = m.b;
        if (form === 'line' && !cdMode) {
            if (!b) b = addCross(slide, NAME_B, c.x, c.y, 8, E.RGB_BLUE);
        } else if (b) {
            try { b.Delete(); } catch (e2) { }
            b = null;
        }
        if (!a || (form === 'line' && !cdMode && !b)) {
            return { a: a, b: b, pc: m.pc, pd: m.pd, pm: m.pm, tr: m.tr, ws: m.ws, cdMode: cdMode, err: '标记创建失败（图形命名不被支持）' };
        }
        return { a: a, b: b, pc: m.pc, pd: m.pd, pm: m.pm, tr: m.tr, ws: m.ws, cdMode: cdMode, err: null };
    }

    /** 兼容旧名 */
    function syncFlowMarkers(slide, c) { return syncMarkers(slide, c); }

    /** 放置标记点: C(绿)偏左上、D(橙)偏右下, 避免被中心处的 A/B 标记遮挡 */
    function placePointMarker(slide, kind, c) {
        if (!slide || !c) return null;
        var name = (kind === 'C') ? NAME_PC : NAME_PD;
        var color = (kind === 'C') ? E.RGB_GREEN : E.RGB_ORANGE;
        var x = c.x, y = c.y;
        if (kind === 'C') { x -= 42; y -= 32; }
        else { x += 48; y += 36; }
        var exist = findMarker(slide, name);
        if (exist) return exist;
        return addCross(slide, name, x, y, 7, color);
    }

    /** 清除标记点 C/D */
    function clearPointMarkers(slide) {
        var n = 0;
        if (!slide) return 0;
        try {
            var ss = slide.Shapes;
            var cnt = ss.Count;
            var victims = [];
            for (var i = 1; i <= cnt; i++) {
                var s = null;
                try { s = ss.Item(i); } catch (e0) { continue; }
                var nm = '';
                try { nm = s.Name || ''; } catch (e1) { continue; }
                if (nm === NAME_PC || nm === NAME_PD) victims.push(i);
            }
            for (var j = victims.length - 1; j >= 0; j--) {
                try { ss.Item(victims[j]).Delete(); n++; } catch (e2) { }
            }
        } catch (e) { }
        return n;
    }

    /**
     * 找最近的吸附目标, 12pt 半径内; 返回 {x,y,dist,kind} 或 null
     * 目标 = 中心十字(__pthlp_C_*) + 全部顶层图形中心
     * 排除: 本插件其他标记(避免标记互吸) + 被移动图形自身(其中心已有A红十字, C点常需放顶点)
     */
    function nearestSnapTarget(slide, p, excludeId) {
        if (!slide || !p) return null;
        var best = null, bestD = 12;
        try {
            var ss = slide.Shapes;
            var n = ss.Count;
            for (var i = 1; i <= n; i++) {
                var s = null;
                try { s = ss.Item(i); } catch (e0) { continue; }
                var nm = '';
                try { nm = s.Name || ''; } catch (e1) { continue; }
                var isCross = nm.indexOf(MARK + 'C_') === 0;
                if (!isCross) {
                    if (isMarker(s)) continue;
                    if (excludeId !== null && excludeId !== undefined && shapeIdOf(s) === excludeId) continue;
                }
                var t = center(s);
                if (!t) continue;
                var d = Math.sqrt((t.x - p.x) * (t.x - p.x) + (t.y - p.y) * (t.y - p.y));
                if (d < bestD) {
                    bestD = d;
                    best = { x: t.x, y: t.y, dist: d, kind: isCross ? '十字' : '图形' };
                }
            }
        } catch (e) { }
        return best;
    }

    /** 把标记精确移动到吸附点 */
    function snapMarkerTo(m, target) {
        try {
            m.Left = target.x - m.Width / 2;
            m.Top = target.y - m.Height / 2;
            return true;
        } catch (e) { return false; }
    }

    // ================= 曲线路径数学引擎 =================
    // 约定: 输入绝对坐标(pt), 输出 {path: 相对锚点的path字符串, pts: 采样绝对点列(预览/测长用)}

    function relPt(anchor, p, size) {
        return { x: (p.x - anchor.x) / size.w, y: (p.y - anchor.y) / size.h };
    }

    /** 点列折线 (pts 不含锚点) */
    function polylinePath(anchor, pts, size) {
        var s = 'M 0 0';
        for (var i = 0; i < pts.length; i++) {
            var r = relPt(anchor, pts[i], size);
            s += ' L ' + fmtNum(r.x) + ' ' + fmtNum(r.y);
        }
        return { path: s + ' E', pts: [anchor].concat(pts) };
    }

    /**
     * 抛物线: 起点S 终点E, 曲线精确经过顶点P (t=0.5 处过 P)
     * 二次贝塞尔控制点 Q=(4P-S-E)/2, 精确转三次贝塞尔
     */
    function parabolaPath(S, P, E, size) {
        var Q = { x: (4 * P.x - S.x - E.x) / 2, y: (4 * P.y - S.y - E.y) / 2 };
        var c1 = { x: S.x + (Q.x - S.x) * 2 / 3, y: S.y + (Q.y - S.y) * 2 / 3 };
        var c2 = { x: E.x + (Q.x - E.x) * 2 / 3, y: E.y + (Q.y - E.y) * 2 / 3 };
        var r1 = relPt(S, c1, size), r2 = relPt(S, c2, size), r4 = relPt(S, E, size);
        var path = 'M 0 0 C ' + fmtNum(r1.x) + ' ' + fmtNum(r1.y) + ' ' +
            fmtNum(r2.x) + ' ' + fmtNum(r2.y) + ' ' + fmtNum(r4.x) + ' ' + fmtNum(r4.y) + ' E';
        // 采样12点用于预览
        var pts = [];
        for (var i = 0; i <= 12; i++) {
            var t = i / 12, u = 1 - t;
            pts.push({
                x: u * u * u * S.x + 3 * u * u * t * c1.x + 3 * u * t * t * c2.x + t * t * t * E.x,
                y: u * u * u * S.y + 3 * u * u * t * c1.y + 3 * u * t * t * c2.y + t * t * t * E.y
            });
        }
        return { path: path, pts: pts };
    }

    /**
     * 圆弧: 三点定圆(S起, M弧上, E终), 沿 S→M→E 方向的圆弧, 贝塞尔分段(每段≤90°)
     * 三点共线时退化为直线
     */
    function arcPath(S, M, E, size) {
        var d = 2 * (S.x * (M.y - E.y) + M.x * (E.y - S.y) + E.x * (S.y - M.y));
        if (Math.abs(d) < 1e-6) {
            return polylinePath(S, [E], size);
        }
        var ss = S.x * S.x + S.y * S.y, mm = M.x * M.x + M.y * M.y, ee = E.x * E.x + E.y * E.y;
        var ux = (ss * (M.y - E.y) + mm * (E.y - S.y) + ee * (S.y - M.y)) / d;
        var uy = (ss * (E.x - M.x) + mm * (S.x - E.x) + ee * (M.x - S.x)) / d;
        var r = Math.sqrt((S.x - ux) * (S.x - ux) + (S.y - uy) * (S.y - uy));
        var a0 = Math.atan2(S.y - uy, S.x - ux);
        var a1 = Math.atan2(E.y - uy, E.x - ux);
        var am = Math.atan2(M.y - uy, M.x - ux);
        function norm(a) { while (a < 0) a += 2 * Math.PI; while (a >= 2 * Math.PI) a -= 2 * Math.PI; return a; }
        var d1 = norm(a1 - a0), dm = norm(am - a0);
        var total = (dm < d1) ? d1 : d1 - 2 * Math.PI;   // 保证经过 M
        var n = Math.max(1, Math.ceil(Math.abs(total) / (Math.PI / 2)));
        var seg = total / n;
        var k = (4 / 3) * Math.tan(seg / 4);             // 有符号: 方向内建 (四分之一圆时=0.5523)
        var path = 'M 0 0';
        var pts = [S];
        for (var i = 1; i <= n; i++) {
            var t0 = a0 + seg * (i - 1), t1 = a0 + seg * i;
            var p1 = { x: ux + r * Math.cos(t0), y: uy + r * Math.sin(t0) };
            var p4 = { x: ux + r * Math.cos(t1), y: uy + r * Math.sin(t1) };
            var c1 = { x: p1.x - k * r * Math.sin(t0), y: p1.y + k * r * Math.cos(t0) };
            var c2 = { x: p4.x + k * r * Math.sin(t1), y: p4.y - k * r * Math.cos(t1) };
            var r1 = relPt(S, c1, size), r2 = relPt(S, c2, size), r4 = relPt(S, p4, size);
            path += ' C ' + fmtNum(r1.x) + ' ' + fmtNum(r1.y) + ' ' +
                fmtNum(r2.x) + ' ' + fmtNum(r2.y) + ' ' + fmtNum(r4.x) + ' ' + fmtNum(r4.y);
            for (var j = 1; j <= 6; j++) {               // 每段采样6点
                var tt = j / 6, uu = 1 - tt;
                pts.push({
                    x: uu * uu * uu * p1.x + 3 * uu * uu * tt * c1.x + 3 * uu * tt * tt * c2.x + tt * tt * tt * p4.x,
                    y: uu * uu * uu * p1.y + 3 * uu * uu * tt * c1.y + 3 * uu * tt * tt * c2.y + tt * tt * tt * p4.y
                });
            }
        }
        return { path: path + ' E', pts: pts };
    }

    /** Catmull-Rom 平滑曲线过全部点(航点/描迹), 转三次贝塞尔链 */
    function smoothPath(anchor, pts, size) {
        if (!pts || pts.length < 2) return { path: 'M 0 0 E', pts: pts || [] };
        if (pts.length === 2) return polylinePath(anchor, [pts[1]], size);
        var S = anchor;
        var path = 'M 0 0';
        for (var i = 0; i < pts.length - 1; i++) {
            var p0 = pts[Math.max(0, i - 1)], p1 = pts[i], p2 = pts[i + 1], p3 = pts[Math.min(pts.length - 1, i + 2)];
            var c1 = { x: p1.x + (p2.x - p0.x) / 6, y: p1.y + (p2.y - p0.y) / 6 };
            var c2 = { x: p2.x - (p3.x - p1.x) / 6, y: p2.y - (p3.y - p1.y) / 6 };
            var r1 = relPt(S, c1, size), r2 = relPt(S, c2, size), r4 = relPt(S, p2, size);
            path += ' C ' + fmtNum(r1.x) + ' ' + fmtNum(r1.y) + ' ' +
                fmtNum(r2.x) + ' ' + fmtNum(r2.y) + ' ' + fmtNum(r4.x) + ' ' + fmtNum(r4.y);
        }
        return { path: path + ' E', pts: pts.slice() };
    }

    /** 三点滑动平均轻量平滑(描迹降噪, 保留首尾) */
    function smoothPts(pts) {
        if (!pts || pts.length < 5) return pts;
        var out = [pts[0]];
        for (var i = 1; i < pts.length - 1; i++) {
            out.push({
                x: (pts[i - 1].x + pts[i].x + pts[i + 1].x) / 3,
                y: (pts[i - 1].y + pts[i].y + pts[i + 1].y) / 3
            });
        }
        out.push(pts[pts.length - 1]);
        return out;
    }

    /** 解析 path 字符串终点(兼容 L 与 C 结尾) */
    function pathEndpoint(str) {
        if (!str || typeof str !== 'string') return null;
        var m = str.match(/C\s*(-?[\d.eE+]+)\s+(-?[\d.eE+]+)\s+(-?[\d.eE+]+)\s+(-?[\d.eE+]+)\s+(-?[\d.eE+]+)\s+(-?[\d.eE+]+)\s*[ZE]?\s*$/);
        if (m) return { x: parseFloat(m[5]), y: parseFloat(m[6]) };
        m = str.match(/L\s*(-?[\d.eE+]+)\s+(-?[\d.eE+]+)\s*(?:[ZE]\s*)?$/);
        if (m) return { x: parseFloat(m[1]), y: parseFloat(m[2]) };
        return null;
    }

    /** 曲线预览: 用多段短线画采样点列 */
    function updateRoutePreview(slide, pts) {
        if (!slide || !pts || pts.length < 2) return 0;
        // 清理旧预览段
        try {
            var ss = slide.Shapes, cnt = ss.Count, victims = [];
            for (var i = 1; i <= cnt; i++) {
                var s = null;
                try { s = ss.Item(i); } catch (e0) { continue; }
                var nm = '';
                try { nm = s.Name || ''; } catch (e1) { continue; }
                if (nm.indexOf(NAME_PV) === 0) victims.push(i);
            }
            for (var j = victims.length - 1; j >= 0; j--) {
                try { ss.Item(victims[j]).Delete(); } catch (e2) { }
            }
        } catch (e3) { }
        var drawn = 0;
        for (var k = 1; k < pts.length && drawn < 48; k++) {
            var a = pts[k - 1], b = pts[k];
            if (Math.abs(b.x - a.x) < 0.5 && Math.abs(b.y - a.y) < 0.5) continue;
            var ln = null;
            try { ln = slide.Shapes.AddLine(a.x, a.y, b.x, b.y); } catch (e4) { continue; }
            try { ln.Name = NAME_PV + drawn; } catch (e5) { }
            try { ln.Line.DashStyle = E.msoLineDash; } catch (e6) { }
            try { ln.Line.ForeColor.RGB = E.RGB_RED; ln.Line.Weight = 1; } catch (e7) { }
            drawn++;
        }
        return drawn;
    }

    /** 通用标记放置(不存在才创建) */
    function placeNamedMarker(slide, name, x, y, size, rgb) {
        if (!slide) return null;
        var exist = findMarker(slide, name);
        if (exist) return exist;
        return addCross(slide, name, x, y, size, rgb);
    }

    /** 按前缀删除标记(航点/预览段等) */
    function deleteMarkersByPrefix(slide, prefix) {
        var n = 0;
        if (!slide) return 0;
        try {
            var ss = slide.Shapes, cnt = ss.Count, victims = [];
            for (var i = 1; i <= cnt; i++) {
                var s = null;
                try { s = ss.Item(i); } catch (e0) { continue; }
                var nm = '';
                try { nm = s.Name || ''; } catch (e1) { continue; }
                if (nm.indexOf(prefix) === 0) victims.push(i);
            }
            for (var j = victims.length - 1; j >= 0; j--) {
                try { ss.Item(victims[j]).Delete(); n++; } catch (e2) { }
            }
        } catch (e3) { }
        return n;
    }

    /** 单遍扫描全部流程标记(含形态标记) */
    function scanAllMarkers(slide) {
        var out = { a: null, b: null, pc: null, pd: null, pm: null, tr: null, ws: [] };
        if (!slide) return out;
        try {
            var ss = slide.Shapes, n = ss.Count;
            for (var i = 1; i <= n; i++) {
                var s = null;
                try { s = ss.Item(i); } catch (e0) { continue; }
                var nm = '';
                try { nm = s.Name || ''; } catch (e1) { continue; }
                if (nm === NAME_A) out.a = s;
                else if (nm === NAME_B) out.b = s;
                else if (nm === NAME_PC) out.pc = s;
                else if (nm === NAME_PD) out.pd = s;
                else if (nm === NAME_PM) out.pm = s;
                else if (nm === NAME_TR) out.tr = s;
                else if (nm.indexOf(NAME_WP) === 0) {
                    var idx = parseInt(nm.slice(NAME_WP.length), 10);
                    if (isFinite(idx)) out.ws.push({ i: idx, s: s });
                }
            }
            out.ws.sort(function (p, q) { return p.i - q.i; });
            out.ws = out.ws.map(function (w) { return w.s; });
        } catch (e2) { }
        return out;
    }

    /** A→B 红色虚线预览 (重建式更新, 由调用方节流) */
    function updatePreviewLine(slide, p1, p2) {
        if (!slide) return null;
        try {
            var old = findMarker(slide, NAME_LINE);
            if (old) old.Delete();
        } catch (e) { }
        if (Math.abs(p2.x - p1.x) < 0.01 && Math.abs(p2.y - p1.y) < 0.01) return null;
        var ln;
        try { ln = slide.Shapes.AddLine(p1.x, p1.y, p2.x, p2.y); } catch (e) { return null; }
        if (!tagMarker(ln, NAME_LINE)) {   // 命名失败会导致预览线无限累积 → 放弃预览
            try { ln.Delete(); } catch (e0) { }
            return null;
        }
        try { ln.Line.DashStyle = E.msoLineDash; } catch (e) { }
        try { ln.Line.ForeColor.RGB = E.RGB_RED; } catch (e) { }
        try { ln.Line.Weight = 1.25; } catch (e) { }
        try { ln.ZOrder(E.msoBringToFront); } catch (e) { }
        return ln;
    }

    /** 清理判定: 流程标记按名/前缀删; 中心十字(C_)仅全清时删; 无名小十字视为残留 */
    function markerVictim(s, all) {
        var nm = '';
        try { nm = s.Name || ''; } catch (e) { return false; }
        if (nm === NAME_A || nm === NAME_B || nm === NAME_LINE || nm === NAME_PC ||
            nm === NAME_PD || nm === NAME_PM || nm === NAME_TR) return true;
        if (nm.indexOf(NAME_WP) === 0 || nm.indexOf(NAME_PV) === 0) return true;
        if (nm.indexOf(MARK + 'C_') === 0) return all;
        if (all) return isMarker(s);
        try {
            if (s.Type === E.msoShapeTypeAutoShape && s.AutoShapeType === E.msoShapeCross &&
                s.Width <= 20 && s.Height <= 20) return true;
        } catch (e1) { }
        return false;
    }

    function deleteMarkersOn(slide, all) {
        var n = 0;
        if (!slide) return 0;
        try {
            var ss = slide.Shapes;
            var count = ss.Count;
            var victims = [];
            for (var i = 1; i <= count; i++) {
                var s = null;
                try { s = ss.Item(i); } catch (e0) { continue; }
                if (markerVictim(s, all)) victims.push(i);
            }
            for (var j = victims.length - 1; j >= 0; j--) {
                try { ss.Item(victims[j]).Delete(); n++; } catch (e1) { }
            }
        } catch (e) { }
        return n;
    }

    function cleanFlowMarkers(slide) { return deleteMarkersOn(slide, false); }

    function cleanSlideMarkers(slide) { return deleteMarkersOn(slide, true); }

    function cleanAllMarkers(pres) {
        var n = 0;
        if (!pres) return 0;
        try {
            for (var i = 1; i <= pres.Slides.Count; i++) n += deleteMarkersOn(pres.Slides.Item(i), true);
        } catch (e) { }
        return n;
    }

    /** 给图形中心放/移除红十字 (静态, 不随图形移动); true=已加 false=已移除 null=失败 */
    function toggleCenterCross(slide, shape) {
        if (!slide) return null;
        var name = MARK + 'C_' + shapeIdOf(shape);
        var exist = findMarker(slide, name);
        if (exist) {
            try { exist.Delete(); } catch (e) { }
            return false;
        }
        var c = center(shape);
        if (!c) return null;
        return !!addCross(slide, name, c.x, c.y, 7, E.RGB_RED);
    }

    // ---------------- 路径动画 (三级策略链) ----------------

    function fmtNum(v) {
        var s = v.toFixed(6);
        if (parseFloat(s) === 0) s = '0.000000';  // 避免 "-0.000000"
        return s;
    }

    function pathString(dx, dy) {
        return 'M 0 0 L ' + fmtNum(dx) + ' ' + fmtNum(dy) + ' E';
    }

    function findMotionBehavior(effect) {
        try {
            var bs = effect.Behaviors;
            for (var i = 1; i <= bs.Count; i++) {
                var b = bs.Item(i);
                if (b.Type === E.msoAnimTypeMotion) return b;
            }
        } catch (e) { }
        return null;
    }

    /** 解析 path 字符串里的最后一个 "L x y" (兼容科学计数法) */
    function parsePath(str) {
        if (!str || typeof str !== 'string') return null;
        var last = null;
        var re = /L\s*(-?[\d.]+(?:[eE][-+]?\d+)?)\s+(-?[\d.]+(?:[eE][-+]?\d+)?)/g;
        var m;
        while ((m = re.exec(str))) {
            last = { x: parseFloat(m[1]), y: parseFloat(m[2]) };
        }
        return last;
    }

    /** 回读 MotionEffect: 优先 Path(终点取 L/C 尾点), 其次 ToX/ToY */
    function readbackMotion(me) {
        var out = { kind: null, dx: null, dy: null, raw: '' };
        try {
            var p = me.Path;
            if (typeof p === 'string' && p) {
                out.raw = p;
                var pt = pathEndpoint(p) || parsePath(p);
                if (pt) {
                    out.kind = 'path';
                    out.dx = pt.x;
                    out.dy = pt.y;
                    return out;
                }
            }
        } catch (e) { }
        try {
            var tx = me.ToX, ty = me.ToY;
            if (typeof tx === 'number' && typeof ty === 'number') {
                out.kind = 'to';
                out.dx = tx;
                out.dy = ty;
                out.raw = 'ToX=' + tx + ' ToY=' + ty;
            }
        } catch (e) { }
        return out;
    }

    function closeEnough(a, b) { return Math.abs(a - b) < 1e-4; }

    /** PathRight 预设动画的默认位移 (用于识别"写入被静默忽略") */
    function looksDefaultPreset(rb) {
        return !!rb && rb.kind !== null &&
            closeEnough(rb.dx, 0.25) && closeEnough(rb.dy, 0);
    }

    /**
     * 多形式尝试 AddEffect 并捕获真实错误信息
     * @return {eff, err} err 为各形式失败原因(拼接)
     */
    function addEffectSafe(seq, shape, effectId, trigger) {
        var errs = [];
        // 官方 TS 签名: AddEffect(Shape, effectId, Level, trigger, Index); 文档示例也出现过 3 参(Shape, effectId, trigger)
        var forms = [
            function () { return seq.AddEffect(shape, effectId, E.msoAnimateLevelNone, trigger); },
            function () { return seq.AddEffect(shape, effectId, trigger); },
            function () { return seq.AddEffect(shape, effectId); },
            function () { return seq.AddEffect(shape, effectId, E.msoAnimateLevelNone, trigger, -1); }
        ];
        // 环境若提供枚举常量对象, 追加一次用枚举对象调用
        var enumVal = null;
        try {
            var app = APP();
            var en = app && app.Enum;
            if (en) {
                var key = (effectId === E.msoAnimEffectPathRight) ? 'msoAnimEffectPathRight' : 'msoAnimEffectCustom';
                if (typeof en[key] !== 'undefined' && en[key] !== effectId) enumVal = en[key];
            }
        } catch (e0) { }
        if (enumVal !== null) {
            var ev = enumVal;
            forms.push(function () { return seq.AddEffect(shape, ev); });
        }
        for (var i = 0; i < forms.length; i++) {
            try {
                var eff = forms[i]();
                if (eff) return { eff: eff, err: null };
            } catch (e) {
                errs.push('形式' + (i + 1) + ':' + (e.message || String(e)));
            }
        }
        return {
            eff: null,
            err: errs.length ? errs.join(' | ') : '所有形式均返回空(未抛错)——图形/时间线引用可能已失效'
        };
    }

    /**
     * WPS 实测要点(2026-11, COM 层与 JS API 同源):
     * 1. Sequence/Effect 等代理在时间线变更后会失效 —— 每次操作前必须重新获取
     *    (slide.TimeLine.MainSequence 每次重新走链)
     * 2. MotionEffect.Path 可读可写, 且是唯一可靠的路径写入方式
     * 3. MotionEffect.FromX/ToX 写入会被静默忽略(回读恒为默认值), 不可用
     */
    function freshSeq(slide) {
        return slide.TimeLine.MainSequence;
    }

    /** 取时间线上最新的 Effect (AddEffect 追加到末尾) */
    function lastEffect(slide) {
        var s = freshSeq(slide);
        var n = 0;
        try { n = s.Count; } catch (e) { return null; }
        if (n < 1) return null;
        return s.Item(n);
    }

    /** 重新走链读取最新 Effect 的 Motion 行为 Path 字符串 */
    function readLastPath(slide) {
        try {
            var eff = lastEffect(slide);
            if (!eff) return null;
            var b = findMotionBehavior(eff);
            if (!b) return null;
            return readbackMotion(b.MotionEffect);
        } catch (e) {
            return null;
        }
    }

    /** 删除时间线上最新 Effect (失败策略的清理; 注意用全新引用) */
    function deleteLastEffect(slide) {
        try {
            var eff = lastEffect(slide);
            if (eff) eff.Delete();
        } catch (e) { }
    }

    /**
     * 创建精确直线路径动画
     * 定位策略: 优先使用调用方传入的活引用 shapeRef(实测新版 Id 会漂移),
     * 失败再按 Id 全新走链扫描。TimeLine.MainSequence 始终每次重新获取。
     * @param pres     演示文稿
     * @param slideIdx 幻灯片索引
     * @param shapeId  移动图形 Id(备用定位)
     * @param dx       归一化位移 (B.x-A.x)/幻灯片宽
     * @param dy       归一化位移 (B.y-A.y)/幻灯片高
     * @param opts     {duration, trigger, smooth}
     * @param shapeRef 移动图形活引用(首选)
     * @return {ok, strategy, msg[], readback, effect}
     */
    function createPathEffect(pres, slideIdx, shapeId, dx, dy, opts, shapeRef) {
        return createPathEffectByPath(pres, slideIdx, shapeId, pathString(dx, dy), { x: dx, y: dy }, opts, shapeRef);
    }

    /**
     * 通用路径动画创建: 接受任意 path 字符串(直线/贝塞尔/折线)
     * @param pathStr  完整路径串, 如 "M 0 0 C ... E"
     * @param endRel   预期终点(归一化), 用于回读校验; null 则只做整串比对
     */
    function createPathEffectByPath(pres, slideIdx, shapeId, pathStr, endRel, opts, shapeRef) {
        opts = opts || {};
        var report = { ok: false, strategy: 0, msg: [], readback: null, effect: null };
        var trigger = opts.trigger || E.msoAnimTriggerOnPageClick;

        // ---- 定位 slide / shape / seq ----
        var slide = null, shape = null, seq = null;
        try {
            slide = pres.Slides.Item(slideIdx);
        } catch (e2) {
            report.msg.push('无法定位幻灯片(索引 ' + slideIdx + '): ' + (e2.message || e2));
            return report;
        }
        // 引用校验: 僵尸引用(坐标全 0)不用; 新鲜 Id/名称扫描优先
        if (shapeRef && looksAlive(shapeRef)) {
            shape = shapeRef;
            report.msg.push('图形定位: 活引用');
        } else if (shapeRef) {
            report.msg.push('传入引用已失效(坐标全 0), 改用扫描');
        }
        if (!shape) {
            shape = refetchShapeFresh(pres, slideIdx, shapeId, null);
            if (!shape) {
                try {
                    var st = findShapeById.lastStats || {};
                    report.msg.push('无法定位移动图形(Id=' + shapeId + ', 扫描 ' + (st.scanned || 0) +
                        ' 个图形, ' + (st.unreadable || 0) + ' 个读取失败)');
                } catch (e3) {
                    report.msg.push('无法定位移动图形(Id=' + shapeId + ')');
                }
                return report;
            }
            report.msg.push('图形定位: 新鲜扫描');
        }
        try {
            seq = slide.TimeLine.MainSequence;
        } catch (e4) {
            report.msg.push('无法访问动画时间线: ' + (e4.message || e4));
            return report;
        }

        function freshSeqLocal() {
            try { return slide.TimeLine.MainSequence; } catch (e) { return null; }
        }
        /** 校验: 优先整串一致(空格归一), 否则终点(L/C尾点)比对 */
        function verify() {
            var s2 = freshSeqLocal();
            if (!s2) return null;
            var n = 0;
            try { n = s2.Count; } catch (e0) { return null; }
            if (n < 1) return null;
            var eff = null;
            try { eff = s2.Item(n); } catch (e1) { return null; }
            var b = findMotionBehavior(eff);
            if (!b) return null;
            var rb = readbackMotion(b.MotionEffect);
            report.readback = rb;
            if (!rb.kind) return null;
            // 整串比对
            try {
                var norm = function (x) { return String(x).replace(/\s+/g, ' ').replace(/\s+E\s*$/, ' E').trim(); };
                if (norm(rb.raw) === norm(pathStr)) return rb;
            } catch (e2) { }
            // 终点比对
            var ep = pathEndpoint(rb.raw);
            if (ep && endRel && closeEnough(ep.x, endRel.x) && closeEnough(ep.y, endRel.y)) {
                rb.endpointOk = true;
                return rb;
            }
            return null;
        }

        function applyTiming() {
            try {
                var s3 = freshSeqLocal();
                var eff = s3 ? s3.Item(s3.Count) : null;
                if (!eff) return;
                var t = eff.Timing;
                try { t.Duration = opts.duration; } catch (e) {
                    try { t.Speed = 1 / (opts.duration || 2); } catch (e2) { }
                }
                var sm = opts.smooth ? E.msoTrue : E.msoFalse;
                try { t.SmoothStart = sm; } catch (e) { }
                try { t.SmoothEnd = sm; } catch (e) { }
                try { t.TriggerType = trigger; } catch (e) { }
            } catch (e) { }
        }

        // ---- 策略1: 预设路径动画打底 + MotionEffect.Path 直写 (COM 实测可用, 首选) ----
        try {
            var r1 = addEffectSafe(seq, shape, E.msoAnimEffectPathRight, trigger);
            if (!r1.eff) throw new Error('AddEffect 全部形式失败 → ' + r1.err);
            var b1 = findMotionBehavior(r1.eff);
            if (!b1) throw new Error('预设路径动画不含 Motion 行为');
            b1.MotionEffect.Path = pathStr;
            var rb1 = verify();          // 重新走链回读
            if (rb1) {
                applyTiming();
                report.ok = true; report.strategy = 1;
                report.msg.push('策略1成功: 预设路径 + Path 直写, 回读一致');
                return report;
            }
            var raw = (report.readback && report.readback.raw) || '(空)';
            if (looksDefaultPreset(report.readback) && !endRel) {
                throw new Error('Path 写入被忽略, 回读仍为预设默认值: ' + raw);
            }
            throw new Error('Path 写入后回读不符: ' + raw);
        } catch (e) {
            report.msg.push('策略1(预设+Path直写)未通过: ' + (e.message || e));
            try { deleteLastEffect(slide); } catch (e0) { }
        }

        // ---- 策略2: 自定义动画 + Behaviors.Add(motion) + 对新增行为写 Path ----
        try {
            var r2 = addEffectSafe(freshSeqLocal() || seq, shape, E.msoAnimEffectCustom, trigger);
            if (!r2.eff) throw new Error('AddEffect(msoAnimEffectCustom) 全部形式失败 → ' + r2.err);
            var beh = null;
            try { beh = r2.eff.Behaviors.Add(E.msoAnimTypeMotion); } catch (e1) { }
            if (!beh) throw new Error('Behaviors.Add 未返回对象');
            beh.MotionEffect.Path = pathStr;
            var rb2 = verify();
            if (rb2) {
                applyTiming();
                report.ok = true; report.strategy = 2;
                report.msg.push('策略2成功: 自定义动画 + 新增 Motion 行为 + Path 直写, 回读一致');
                return report;
            }
            throw new Error('新增行为 Path 写入后回读不符: ' + ((report.readback && report.readback.raw) || '(空)'));
        } catch (e) {
            report.msg.push('策略2(Custom+Behaviors.Add+Path)未通过: ' + (e.message || e));
            try { deleteLastEffect(slide); } catch (e2) { }
        }

        // ---- 策略3: 全新鲜引用重试 (AddEffect 静默返回 null 时, 换全新扫描的图形+幻灯片) ----
        try {
            var slide3 = pres.Slides.Item(slideIdx);
            var shape3 = refetchShapeFresh(pres, slideIdx, shapeId, null);
            if (!shape3) throw new Error('新鲜扫描无法定位图形');
            var seq3 = slide3.TimeLine.MainSequence;
            var r3 = addEffectSafe(seq3, shape3, E.msoAnimEffectPathRight, trigger);
            if (!r3.eff) throw new Error('AddEffect 仍失败 → ' + r3.err);
            var b3 = findMotionBehavior(r3.eff);
            if (!b3) throw new Error('无 Motion 行为');
            b3.MotionEffect.Path = pathStr;
            var s4 = slide3.TimeLine.MainSequence;
            var n4 = s4.Count;
            var eff4 = n4 > 0 ? s4.Item(n4) : null;
            var b4 = eff4 ? findMotionBehavior(eff4) : null;
            var rb3 = b4 ? readbackMotion(b4.MotionEffect) : null;
            report.readback = rb3;
            var ok3 = false;
            if (rb3 && rb3.kind) {
                var ep3 = pathEndpoint(rb3.raw);
                ok3 = !!ep3 && endRel && closeEnough(ep3.x, endRel.x) && closeEnough(ep3.y, endRel.y);
            }
            if (ok3) {
                applyTiming();
                report.ok = true; report.strategy = 3;
                report.msg.push('策略3成功: 全新鲜引用 + Path 直写, 回读一致');
                return report;
            }
            throw new Error('回读不符: ' + ((rb3 && rb3.raw) || '(空)'));
        } catch (e) {
            report.msg.push('策略3(全新鲜引用)未通过: ' + (e.message || e));
            try { deleteLastEffect(slide); } catch (e2) { }
        }

        report.msg.push('所有策略均失败, 请点「API 诊断」并把结果发给开发者');
        return report;
    }

    /** 删除某图形在主时间线上的全部动画, 返回删除数(-1=失败); Id 不稳定时按名称兜底 */
    function deleteShapeAnimations(slide, shape) {
        var n = 0;
        var targetId = shapeIdOf(shape);
        var targetName = '';
        try { targetName = shape.Name || ''; } catch (e0) { }
        var guard = 200;
        function isTarget(sh) {
            try {
                if (shapeIdOf(sh) === targetId) return true;
                if (targetName && targetName.length > 1 && sh.Name === targetName) return true;
            } catch (e1) { }
            return false;
        }
        try {
            // 每轮重新获取 Sequence 引用(变更后旧引用会失效)
            var i = freshSeq(slide).Count;
            while (i >= 1 && guard-- > 0) {
                var eff = null, es = null, hit = false;
                try {
                    eff = freshSeq(slide).Item(i);
                    es = eff.Shape;
                    if (es && isTarget(es)) hit = true;
                } catch (e) { }
                if (hit) {
                    try { eff.Delete(); n++; } catch (e2) { }
                    // 删除后重新计数, 防止索引漂移
                    i = freshSeq(slide).Count;
                } else {
                    i--;
                }
            }
        } catch (e) {
            return n > 0 ? n : -1;
        }
        return n;
    }

    // ---------------- API 诊断 ----------------

    /**
     * 诊断报告: 逐项探测能力 (会创建临时图形+临时动画, 结束后清理)
     * @return [{name, ok, detail, error}]
     */
    function diagnose() {
        var steps = [];
        function step(name, fn) {
            try {
                var r = fn();
                steps.push({ name: name, ok: r !== false && r !== null && r !== '', detail: (typeof r === 'string') ? r : '', error: null });
            } catch (e) {
                steps.push({ name: name, ok: false, detail: '', error: e.message || String(e) });
            }
        }

        var pres = getPresentation();
        step('打开的演示文稿', function () { return pres ? pres.Name : false; });
        if (!pres) return steps;

        step('幻灯片尺寸(磅)', function () {
            var s = slideSize(pres);
            return s.ok ? (s.w + ' × ' + s.h) : false;
        });

        var slide = getActiveSlide();
        step('当前幻灯片', function () {
            return slide ? ('第 ' + slideIndexOf(slide) + ' 页') : false;
        });
        if (!slide) return steps;

        step('动画时间线', function () {
            return 'MainSequence 现有 ' + slide.TimeLine.MainSequence.Count + ' 个动画';
        });

        step('选择变化事件 API', function () {
            var app = APP();
            return (app && app.ApiEvent && typeof app.ApiEvent.AddApiEventListener === 'function')
                ? 'ApiEvent.AddApiEventListener 可用' : false;
        });

        var tmp = null, eff = null;
        step('创建临时图形', function () {
            tmp = slide.Shapes.AddShape(E.msoShapeRectangle, -60, -60, 8, 8);
            try { tmp.Name = MARK + 'DIAG__'; } catch (e) { }
            return true;
        });

        if (tmp) {
            step('AddEffect(预设路径)', function () {
                var r = addEffectSafe(freshSeq(slide), tmp, E.msoAnimEffectPathRight, E.msoAnimTriggerWithPrevious);
                eff = r.eff;
                return r.eff ? '已创建' : ('失败 → ' + r.err);
            });
            if (eff) {
                step('Motion 行为', function () {
                    var b = findMotionBehavior(eff);
                    return b ? ('Behaviors 含 Motion (共 ' + eff.Behaviors.Count + ' 个行为)') : false;
                });
                step('写入并回读 Path(全新引用)', function () {
                    var me = findMotionBehavior(eff);
                    if (!me) return false;
                    me.MotionEffect.Path = pathString(0.123456, -0.065432);
                    var rb = readLastPath(slide);   // 重新走链回读
                    if (rb && rb.kind === 'path' && closeEnough(rb.dx, 0.123456) && closeEnough(rb.dy, -0.065432)) {
                        return '写入成功, 回读一致: ' + rb.raw;
                    }
                    return '写入后回读: ' + ((rb && rb.raw) || '(空)') + ' (可能不被支持)';
                });
                step('FromX/ToX 写入(信息项, WPS 通常忽略)', function () {
                    var b = findMotionBehavior(eff);
                    if (!b) return false;
                    var me = b.MotionEffect;
                    try {
                        me.FromX = 0; me.ToX = 0.111111; me.ToY = 0.222222;
                    } catch (e0) { }
                    var rb = readbackMotion(me);
                    return 'ToX 读回 ' + (function () { try { return me.ToX; } catch (e1) { return '?'; } })() +
                        ' (' + ((rb && rb.raw) || '空') + ')';
                });
                step('Timing 设置', function () {
                    var e2 = lastEffect(slide);   // 全新引用
                    e2.Timing.Duration = 1.5;
                    e2.Timing.TriggerType = E.msoAnimTriggerWithPrevious;
                    var d = e2.Timing.Duration;
                    return (typeof d === 'number' && Math.abs(d - 1.5) < 0.01) ? 'Duration/TriggerType 写入成功' : ('Duration 读回 ' + d);
                });
            }
            step('Effect.Delete(全新引用)', function () {
                var e3 = lastEffect(slide);
                e3.Delete();
                return '已删除, 剩余 ' + freshSeq(slide).Count + ' 个动画';
            });
            try { tmp.Delete(); } catch (e) { }
        }
        return steps;
    }

    // ---------------- 导出 ----------------

    return {
        APP: APP,
        getPresentation: getPresentation,
        parentSlide: parentSlide,
        getActiveSlide: getActiveSlide,
        slideIndexOf: slideIndexOf,
        slideByIndex: slideByIndex,
        slideSize: slideSize,
        shapeIdOf: shapeIdOf,
        looksAlive: looksAlive,
        refetchShapeFresh: refetchShapeFresh,
        locateShape: locateShape,
        findBy: findBy,
        findByGeom: findByGeom,
        resolveTop: resolveTop,
        isMarker: isMarker,
        selectionInfo: selectionInfo,
        findShapeById: findShapeById,
        findMarker: findMarker,
        center: center,
        ensureAMarker: ensureAMarker,
        ensureBMarker: ensureBMarker,
        syncFlowMarkers: syncFlowMarkers,
        syncMarkers: syncMarkers,
        placePointMarker: placePointMarker,
        clearPointMarkers: clearPointMarkers,
        nearestSnapTarget: nearestSnapTarget,
        snapMarkerTo: snapMarkerTo,
        polylinePath: polylinePath,
        parabolaPath: parabolaPath,
        arcPath: arcPath,
        smoothPath: smoothPath,
        smoothPts: smoothPts,
        pathEndpoint: pathEndpoint,
        updateRoutePreview: updateRoutePreview,
        placeNamedMarker: placeNamedMarker,
        deleteMarkersByPrefix: deleteMarkersByPrefix,
        scanAllMarkers: scanAllMarkers,
        createPathEffectByPath: createPathEffectByPath,
        NAME_PM: NAME_PM,
        NAME_TR: NAME_TR,
        NAME_WP: NAME_WP,
        NAME_PV: NAME_PV,
        updatePreviewLine: updatePreviewLine,
        cleanFlowMarkers: cleanFlowMarkers,
        cleanSlideMarkers: cleanSlideMarkers,
        cleanAllMarkers: cleanAllMarkers,
        toggleCenterCross: toggleCenterCross,
        pathString: pathString,
        createPathEffect: createPathEffect,
        deleteShapeAnimations: deleteShapeAnimations,
        diagnose: diagnose,
        MARK: MARK
    };
})();
