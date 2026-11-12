/**
 * WPS 演示「路径动画 API」可行性探针（独立脚本）
 *
 * 用法（无需安装加载项）:
 *   1. 打开 WPS 演示 → 顶部「工具」→「开发工具」→「JS宏」(WPS 宏编辑器)
 *   2. 新建模块, 粘贴本文件全部内容
 *   3. 运行 RunProbe (光标放在函数内按 F5 / 点运行)
 *   4. 弹出的报告即当前 WPS 版本对路径动画 API 的支持情况
 *
 * 说明: 本脚本复刻了插件 core.js 的三级策略链,
 *       在当前幻灯片创建临时图形+临时动画, 结束后自动清理, 不影响文档内容。
 */

function RunProbe() {
    var log = [];
    function L(s) { log.push(s); }
    function OK(s) { L('[✓] ' + s); }
    function FAIL(s) { L('[✗] ' + s); }
    function tryCatch(name, fn) {
        try {
            var r = fn();
            OK(name + (r ? ' — ' + r : ''));
            return true;
        } catch (e) {
            FAIL(name + ' — ' + (e.message || e));
            return false;
        }
    }

    var app = Application;
    var pres = app.ActivePresentation;
    if (!pres) { alert('请先打开或新建一个演示文稿'); return; }
    var slide = pres.Slides.Item(1);
    var slideW = pres.PageSetup.SlideWidth;
    var slideH = pres.PageSetup.SlideHeight;
    OK('演示文稿: ' + pres.Name + '  幻灯片 ' + slideW + '×' + slideH + ' pt');

    // 枚举数值 (与插件 js/enums.js 一致)
    var msoShapeRectangle = 1;
    var msoAnimEffectPathRight = 149;
    var msoAnimEffectCustom = 0;
    var msoAnimTypeMotion = 1;
    var msoAnimTriggerWithPrevious = 2;

    function findMotionBehavior(eff) {
        var bs = eff.Behaviors;
        for (var i = 1; i <= bs.Count; i++) {
            var b = bs.Item(i);
            if (b.Type === msoAnimTypeMotion) return b;
        }
        return null;
    }
    function parsePathLastL(str) {
        var m = null, last = null, re = /L\s*(-?[\d.]+)\s+(-?[\d.]+)/g;
        while ((m = re.exec(str))) last = { x: parseFloat(m[1]), y: parseFloat(m[2]) };
        return last;
    }

    var seq = slide.TimeLine.MainSequence;
    L('当前动画数: ' + seq.Count);

    // ---- 临时图形 ----
    var shp = slide.Shapes.AddShape(msoShapeRectangle, 40, 40, 80, 60);
    shp.Name = '__probe_tmp__';
    OK('临时图形已创建 (Id=' + shp.Id + ')');

    var target = { dx: 0.187500, dy: -0.083333 };   // 期望写入的位移
    var pathStr = 'M 0 0 L ' + target.dx.toFixed(6) + ' ' + target.dy.toFixed(6) + ' E';

    var eff = null;

    // 注意(实测结论): 时间线每次变更后, Sequence/Effect 引用会失效,
    // 因此每一步都通过 slide.TimeLine.MainSequence 重新获取

    function freshLast() {
        var s = slide.TimeLine.MainSequence;
        return s.Count > 0 ? s.Item(s.Count) : null;
    }
    function freshLastPath() {
        try {
            var e = freshLast();
            if (!e) return null;
            var b = findMotionBehavior(e);
            if (!b) return null;
            return String(b.MotionEffect.Path);
        } catch (e) { return null; }
    }

    // ---- 策略1: 预设路径 + Path 直写 ----
    var s1 = false;
    try {
        eff = slide.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectPathRight, 0, msoAnimTriggerWithPrevious);
        var b1 = findMotionBehavior(eff);
        if (!b1) throw new Error('无 Motion 行为');
        var p0 = b1.MotionEffect.Path;
        b1.MotionEffect.Path = pathStr;
        var p1 = freshLastPath();          // 全新引用回读
        var pt = p1 ? parsePathLastL(p1) : null;
        if (pt && Math.abs(pt.x - target.dx) < 1e-4 && Math.abs(pt.y - target.dy) < 1e-4) {
            OK('策略1(预设+Path直写): 写入并回读一致 → ' + p1 + ' (原始预设: ' + p0 + ')');
            s1 = true;
        } else {
            FAIL('策略1(Path直写): 回读为 «' + p1 + '» (期望 «' + pathStr + '»)');
        }
    } catch (e) {
        FAIL('策略1(Path直写): ' + (e.message || e));
    }
    if (!s1) {
        try { freshLast().Delete(); } catch (e0) { }
        eff = null;
    }

    // ---- 策略2: 自定义动画 + Behaviors.Add(motion) + Path 直写 ----
    if (!s1) {
        try {
            eff = slide.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectCustom, 0, msoAnimTriggerWithPrevious);
            var beh = eff.Behaviors.Add(msoAnimTypeMotion);
            beh.MotionEffect.Path = pathStr;
            var p2 = freshLastPath();
            var pt2 = p2 ? parsePathLastL(p2) : null;
            if (pt2 && Math.abs(pt2.x - target.dx) < 1e-4 && Math.abs(pt2.y - target.dy) < 1e-4) {
                OK('策略2(Custom+Behaviors.Add+Path): 写入并回读一致 → ' + p2);
            } else {
                FAIL('策略2(Custom+Add+Path): 回读为 «' + p2 + '»');
            }
        } catch (e) {
            FAIL('策略2(Custom+Add+Path): ' + (e.message || e));
        }
    }

    // ---- FromX/ToX (信息项: 实测 WPS 会忽略这些写入) ----
    tryCatch('FromX/ToX 写入探测(预期被忽略)', function () {
        var e = freshLast();
        var b = findMotionBehavior(e);
        b.MotionEffect.ToX = 0.5;
        return '写 ToX=0.5 后读回 ' + b.MotionEffect.ToX;
    });

    // ---- Timing ----
    tryCatch('Timing.Duration 写入(全新引用)', function () {
        var e = freshLast();
        e.Timing.Duration = 1.5;
        return '读回 ' + freshLast().Timing.Duration;
    });
    tryCatch('Timing.SmoothStart 写入', function () {
        freshLast().Timing.SmoothStart = 0;
        return 'ok';
    });

    // ---- Selection/图形属性探测 ----
    tryCatch('Shape.Left/Top/Width/Height 读取', function () {
        var v = shp.Left + ',' + shp.Top + ',' + shp.Width + ',' + shp.Height;
        return v;
    });

    // ---- 清理 ----
    try { freshLast().Delete(); OK('临时动画已清理'); } catch (e) { FAIL('临时动画清理失败'); }
    try { shp.Delete(); OK('临时图形已清理'); } catch (e) { FAIL('临时图形清理失败'); }

    L('');
    L(s1 ? '结论: 支持 MotionEffect.Path 直写, 插件将以策略1(最精确)运行' :
        '结论: 策略1不可用时插件尝试策略2(Custom+新增行为+Path)');
    alert(log.join('\n'));
}

RunProbe();
