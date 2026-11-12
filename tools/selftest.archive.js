/**
 * selftest.js — 【诊断】新版 WPS 桥接自测 (定位插件问题时注入; 稳定后可移除)
 * 每次启动后等待用户打开演示文稿, 在其第1页用画布外临时图形做一轮探测:
 *   选区引用 vs 扫描引用的几何、AddEffect 各种形式、Path 直写、Timing、保存
 * 报告通道: %TEMP%\abp_bridge_report.json (wps.FileSystem) → 本地3889(POST) → PluginStorage(任务窗格展示)
 */
(function () {
    'use strict';

    var REPORT_PATH = 'C:\\Users\\Administrator\\AppData\\Local\\Temp\\abp_bridge_report.json';
    var GUARD_MIN_MS = 10 * 60 * 1000;   // 两次运行最小间隔

    function errText(e) {
        var s = (e && e.message) ? e.message : String(e);
        try { if (e && e.number !== undefined) s += ' [num=0x' + (e.number >>> 0).toString(16) + ']'; } catch (e0) { }
        return s;
    }
    function now() { return Date.now(); }

    function getPs() {
        try {
            var app = window.Application;
            return app && app.PluginStorage ? app.PluginStorage : null;
        } catch (e) { return null; }
    }

    function shouldRun() {
        var ps = getPs();
        if (!ps) return true;
        try {
            var last = parseInt(ps.getItem('abp_st_last') || '0', 10) || 0;
            if (now() - last < GUARD_MIN_MS) return false;
            ps.setItem('abp_st_last', String(now()));
            return true;
        } catch (e) { return true; }
    }

    function emit(out) {
        var json = JSON.stringify(out);
        // 1) FileSystem 写文件
        var fsDone = false;
        try {
            if (typeof wps !== 'undefined' && wps.FileSystem) {
                var names = ['WriteAllText', 'WriteTextFile', 'WriteFile'];
                for (var i = 0; i < names.length; i++) {
                    try {
                        if (typeof wps.FileSystem[names[i]] === 'function') {
                            wps.FileSystem[names[i]](REPORT_PATH, json);
                            fsDone = true; break;
                        }
                    } catch (e0) { }
                }
            }
        } catch (e1) { }
        out.fsChannel = fsDone;
        // 2) POST 本地服务器(可能没开, 失败无所谓)
        try {
            var x = new XMLHttpRequest();
            x.open('POST', 'http://127.0.0.1:3889/__test/result', true);
            x.timeout = 2500;
            x.setRequestHeader('Content-Type', 'text/plain');
            x.send(json);
        } catch (e2) { }
        // 3) PluginStorage → 任务窗格展示
        try {
            var ps = getPs();
            if (ps) ps.setItem('abp_st_report', json);
        } catch (e3) { }
    }

    function runProbe(pres) {
        var out = {
            t: new Date().toISOString(), ver: 'selftest-v2',
            pres: '', slideCount: 0, size: '',
            tempShape: null, scanGeom: null,
            addEffectForms: null, motionBeh: null,
            pathWrite: null, readback: null,
            timing: null, cleanup: null, extra: {}
        };
        var slide = null, shp = null, eff = null;
        try {
            try { out.pres = pres.Name; } catch (e0) { }
            try { out.slideCount = pres.Slides.Count; } catch (e1) { }
            try { out.size = pres.PageSetup.SlideWidth + 'x' + pres.PageSetup.SlideHeight; } catch (e2) { }

            slide = pres.Slides.Item(1);
            shp = slide.Shapes.AddShape(1, -120, -120, 60, 40);   // 画布外临时矩形
            try { shp.Name = '__abp_probe__'; } catch (e3) { }
            out.tempShape = { created: true, named: (function () { try { return shp.Name === '__abp_probe__'; } catch (e4) { return false; } })() };
            out.scanGeom = (function () {
                try {
                    var n = slide.Shapes.Count;
                    for (var i = 1; i <= n; i++) {
                        var c = slide.Shapes.Item(i);
                        if (c.Name === '__abp_probe__') return { L: c.Left, T: c.Top, W: c.Width, H: c.Height };
                    }
                    return 'not-found-by-name';
                } catch (e5) { return 'err:' + errText(e5); }
            })();

            // AddEffect 形式矩阵 (全部用扫描来的新鲜图形引用)
            var seq = slide.TimeLine.MainSequence;
            var app = window.Application;
            var forms = {
                f4_num: function () { return seq.AddEffect(shp, 149, 0, 1); },
                f3_num: function () { return seq.AddEffect(shp, 149, 1); },
                f2_num: function () { return seq.AddEffect(shp, 149); },
                f5_num: function () { return seq.AddEffect(shp, 149, 0, 1, -1); },
                f2_enum: function () { try { return seq.AddEffect(shp, app.Enum.msoAnimEffectPathRight); } catch (e) { throw e; } },
                f4_enum: function () { try { return seq.AddEffect(shp, app.Enum.msoAnimEffectPathRight, 0, 1); } catch (e) { throw e; } },
                f4_custom: function () { return seq.AddEffect(shp, 0, 0, 1); },
                f2_custom: function () { return seq.AddEffect(shp, 0); }
            };
            var res = {};
            Object.keys(forms).forEach(function (k) {
                try {
                    var r = forms[k]();
                    res[k] = r ? 'OK' : 'null';
                } catch (e6) { res[k] = 'ERR:' + errText(e6); }
                try {   // 清理本形式残留
                    var s2 = slide.TimeLine.MainSequence;
                    var n2 = s2.Count;
                    if (n2 > 0) s2.Item(n2).Delete();
                } catch (e7) { }
            });
            out.addEffectForms = res;

            // 正式做一次: Path 直写 + 新鲜引用回读
            var eff2 = null;
            try {
                eff2 = slide.TimeLine.MainSequence.AddEffect(shp, 149, 0, 1);
                if (!eff2) throw new Error('AddEffect 返回 null');
                var bs = eff2.Behaviors.Count, b = null;
                for (var i2 = 1; i2 <= bs; i2++) {
                    if (eff2.Behaviors.Item(i2).Type === 1) b = eff2.Behaviors.Item(i2);
                }
                if (!b) throw new Error('无 Motion 行为(behaviors=' + bs + ')');
                b.MotionEffect.Path = 'M 0 0 L 0.250000 0.125000 E';
                out.pathWrite = 'written';
                var s3 = slide.TimeLine.MainSequence;
                var e3b = s3.Count > 0 ? s3.Item(s3.Count) : null;
                var b3 = null;
                if (e3b) {
                    var bs3 = e3b.Behaviors.Count;
                    for (var i3 = 1; i3 <= bs3; i3++) {
                        if (e3b.Behaviors.Item(i3).Type === 1) b3 = e3b.Behaviors.Item(i3);
                    }
                }
                out.readback = b3 ? String(b3.MotionEffect.Path) : 'no-eff';
                var tm = e3b ? e3b.Timing : null;
                if (tm) {
                    tm.Duration = 1.5;
                    out.timing = 'Duration=' + (function () { try { return tm.Duration; } catch (e8) { return '?'; } })();
                }
            } catch (e9) {
                out.pathWrite = 'ERR:' + errText(e9);
            }
            eff = eff2;

            // ---- 选区模拟 + 生产路径验证(v3) ----
            var app2 = window.Application;
            var t2 = null;
            try {
                t2 = slide.Shapes.AddShape(1, 260, 160, 120, 80);
                try { t2.Name = '__abp_probe2__'; } catch (eN0) { }
                var selInfo = {};
                try { t2.Select(); selInfo.select = 'ok'; } catch (eSel) { selInfo.select = 'ERR:' + errText(eSel); }
                try {
                    var sel = app2.ActiveWindow.Selection;
                    selInfo.type = sel.Type;
                    selInfo.count = sel.ShapeRange.Count;
                    var raw = sel.ShapeRange.Item(1);
                    try { selInfo.rawName = raw.Name; } catch (eN) { selInfo.rawName = 'ERR'; }
                    try { selInfo.rawId = raw.Id; } catch (eI) { selInfo.rawId = 'ERR'; }
                    try { selInfo.rawGeom = raw.Left + ',' + raw.Top + ',' + raw.Width + ',' + raw.Height; } catch (eG) { selInfo.rawGeom = 'ERR'; }
                    // 各幻灯片解析通道
                    try { var vs = app2.ActiveWindow.View.Slide; selInfo.viewSlide = vs ? ('idx=' + vs.SlideIndex) : 'null'; } catch (eV) { selInfo.viewSlide = 'ERR'; }
                    try { var pp = sel.ShapeRange.Parent; selInfo.srParent = pp ? ('hasShapes=' + (typeof pp.Shapes !== 'undefined')) : 'null'; } catch (eP) { selInfo.srParent = 'ERR'; }
                    try { selInfo.srSlideRange = sel.SlideRange.Item(1) ? 'ok' : 'null'; } catch (eS2) { selInfo.srSlideRange = 'ERR'; }
                    // 生产函数
                    try {
                        var si = window.ABP.selectionInfo();
                        var s0 = si && si.shapes && si.shapes[0];
                        selInfo.prodSel = s0 ? (s0.Name + '@' + s0.Left + ',' + s0.Top) : 'no-shapes';
                    } catch (eP2) { selInfo.prodSel = 'ERR:' + errText(eP2); }
                    try {
                        var loc = window.ABP.locateShape(raw);
                        selInfo.prodLocate = loc ? (loc.shape.Name + '@' + loc.shape.Left + ',' + loc.shape.Top) : 'null';
                    } catch (eL2) { selInfo.prodLocate = 'ERR:' + errText(eL2); }
                } catch (eR2) { selInfo.readErr = errText(eR2); }
                out.selectSim = selInfo;
            } catch (eS3) {
                out.selectSim = { fatal: errText(eS3) };
            }
            // 清理 t2
            try {
                var cnt2 = slide.TimeLine.MainSequence.Count;
                for (var d2 = cnt2; d2 >= 1; d2--) slide.TimeLine.MainSequence.Item(d2).Delete();
                if (t2) t2.Delete();
            } catch (eC2) { }
        } catch (eA) {
            out.extra.fatal = errText(eA);
        }
        // 清理
        try {
            if (slide) {
                var cnt = slide.TimeLine.MainSequence.Count;
                for (var d = cnt; d >= 1; d--) slide.TimeLine.MainSequence.Item(d).Delete();
            }
            if (shp) shp.Delete();
            out.cleanup = 'ok';
        } catch (eB) { out.cleanup = 'err:' + errText(eB); }
        return out;
    }

    function waitForPresThenRun() {
        var waited = 0;
        var t = setInterval(function () {
            waited += 2000;
            var done = false;
            try {
                var app = window.Application;
                var pres = app && app.ActivePresentation;
                if (pres) {
                    var cnt = 0;
                    try { cnt = pres.Slides.Count; } catch (e0) { }
                    if (cnt > 0) {
                        clearInterval(t);
                        emit(runProbe(pres));
                        done = true;
                    }
                }
            } catch (e1) { }
            if (!done && waited > 180000) {   // 3分钟没等到文档就放弃
                clearInterval(t);
                emit({ t: new Date().toISOString(), ver: 'selftest-v2', giveup: 'no-presentation-in-3min' });
            }
        }, 2000);
    }

    if (shouldRun()) waitForPresThenRun();
})();
