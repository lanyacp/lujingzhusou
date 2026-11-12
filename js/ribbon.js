/**
 * ribbon.js — ribbon.xml 回调入口
 * 函数名与 ribbon.xml 中属性一一对应: OnAddinLoad / OnAction / GetImage
 */

/** 加载项第一个执行的函数 */
function OnAddinLoad(ribbonUI) {
    try {
        if (typeof (window.Application.ribbonUI) !== 'object') {
            window.Application.ribbonUI = ribbonUI;
        }
    } catch (e) { }
    // 选择变化事件: 仅计数供任务窗格加速响应; 不可用时由任务窗格轮询兜底
    try {
        window.Application.ApiEvent.AddApiEventListener('WindowSelectionChange', function () {
            try {
                var ps = window.Application.PluginStorage;
                ps.setItem('abp_sel_seq', String((parseInt(ps.getItem('abp_sel_seq') || '0', 10) || 0) + 1));
            } catch (e) { }
        });
    } catch (e) { }
    return true;
}

/** 所有按钮的 onAction */
function OnAction(control) {
    var id = control.Id;
    try {
        switch (id) {
            case 'btnCreatePath': ABPToggleTaskpane(); break;
            case 'btnCenterCross': ABPRibbonCenterCross(); break;
            case 'btnClean': ABPRibbonClean(); break;
            case 'btnHelp': ABPRibbonHelp(); break;
        }
    } catch (e) {
        ABPNotify('操作失败: ' + (e.message || e));
    }
    return true;
}

/** getImage 回调 */
function GetImage(control) {
    var map = {
        btnCreatePath: 'images/icon-create.svg',
        btnCenterCross: 'images/icon-cross.svg',
        btnClean: 'images/icon-clean.svg',
        btnHelp: 'images/icon-help.svg'
    };
    return map[control.Id] || 'images/icon-create.svg';
}

/** 打开/关闭任务窗格 (复用同一实例) */
function ABPToggleTaskpane() {
    var app = window.Application;
    var id = null;
    try { id = app.PluginStorage.getItem('abp_taskpane_id'); } catch (e) { }
    if (id) {
        var tp = null;
        try { tp = app.GetTaskPane(id); } catch (e) { }
        if (tp) {
            tp.Visible = !tp.Visible;
            return;
        }
    }
    var tp2 = app.CreateTaskPane(GetUrlPath() + '/ui/taskpane.html', '路径助手');
    app.PluginStorage.setItem('abp_taskpane_id', tp2.ID);
    tp2.Visible = true;
}

/** 功能区「中心十字」: 对当前选中图形逐个 放置/移除 红十字 */
function ABPRibbonCenterCross() {
    var pres = ABP.getPresentation();
    if (!pres) { ABPNotify('请先打开一个演示文稿'); return; }
    var sel = ABP.selectionInfo();
    var shapes = [];
    for (var i = 0; i < sel.shapes.length; i++) {
        if (!ABP.isMarker(sel.shapes[i])) shapes.push(sel.shapes[i]);
    }
    if (!shapes.length) {
        ABPNotify('请先选中一个或多个图形\n（将以其中心放置/移除红十字标记）');
        return;
    }
    var slide = ABP.parentSlide(shapes[0]) || ABP.getActiveSlide();
    if (!slide) { ABPNotify('无法定位图形所在幻灯片'); return; }
    var on = 0, off = 0, bad = 0;
    for (var j = 0; j < shapes.length; j++) {
        var r = ABP.toggleCenterCross(slide, shapes[j]);
        if (r === true) on++;
        else if (r === false) off++;
        else bad++;
    }
    if (bad > 0) {
        ABPNotify(bad + ' 个图形坐标读取失败，未添加中心十字');
    } else if (off > 0) {
        ABPNotify('已移除 ' + off + ' 个中心十字标记');
    } else {
        ABPNotify('已为 ' + on + ' 个图形在中心放置红十字。\n提示: 十字是静态的, 图形移动后请再次点击以刷新。');
    }
}

/** 功能区「清理标记」: 清理整个演示文稿的所有辅助标记 */
function ABPRibbonClean() {
    var pres = ABP.getPresentation();
    if (!pres) { ABPNotify('请先打开一个演示文稿'); return; }
    var n = ABP.cleanAllMarkers(pres);
    // 通知任务窗格停止流程, 避免标记被下一拍轮询重建 (跨页面经 PluginStorage)
    try {
        var ps = window.Application.PluginStorage;
        ps.setItem('abp_clean_seq', String((parseInt(ps.getItem('abp_clean_seq') || '0', 10) || 0) + 1));
    } catch (e) { }
    ABPNotify(n ? ('已清理 ' + n + ' 个辅助标记') : '没有发现辅助标记');
}

/** 功能区「使用帮助」 */
function ABPRibbonHelp() {
    var dpr = 1;
    try { dpr = window.devicePixelRatio || 1; } catch (e) { }
    try {
        wps.ShowDialog(GetUrlPath() + '/ui/help.html', '路径助手 · 使用帮助', 440 * dpr, 580 * dpr, false);
    } catch (e) {
        ABPNotify('三步用法:\n1. 选中要移动的图形\n2. 拖动蓝色十字到目标位置 B\n3. 点任务窗格里的「生成路径动画」');
    }
}
