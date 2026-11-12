/**
 * util.js — 通用工具（按官方 wpsjs 模板风格）
 */

/** 取当前加载项的 URL 根路径（用于 taskpane / dialog 等页面地址） */
function GetUrlPath() {
    return location.href.substring(0, location.href.lastIndexOf('/'));
}

/** 统一弹窗通知（优先 wps.alert） */
function ABPNotify(msg) {
    try {
        if (typeof wps !== 'undefined' && wps && wps.alert) {
            wps.alert(msg);
            return;
        }
    } catch (e) { }
    try {
        alert(msg);
    } catch (e2) { }
}
