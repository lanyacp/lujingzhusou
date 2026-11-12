/**
 * enums.js — 枚举常量（统一使用数值，避免依赖 wps.Enum 的命名/挂载差异）
 * 数值以 WPS 官方 wps-jsapi 类型定义及 Microsoft VBA 对应枚举为准
 */
var ABP_ENUM = {
    // MsoAutoShapeType
    msoShapeRectangle: 1,
    msoShapeCross: 11,

    // MsoShapeType
    msoShapeTypeAutoShape: 1,
    msoShapeTypeGroup: 6,

    // MsoTriState
    msoTrue: -1,
    msoFalse: 0,

    // PpSelectionType
    ppSelectionNone: 0,
    ppSelectionSlides: 1,
    ppSelectionShapes: 2,
    ppSelectionText: 3,

    // MsoAnimEffect（本插件用到的）
    msoAnimEffectCustom: 0,
    msoAnimEffectPathRight: 149,

    // MsoAnimType
    msoAnimTypeMotion: 1,

    // MsoAnimTriggerType
    msoAnimTriggerOnPageClick: 1,   // 单击时
    msoAnimTriggerWithPrevious: 2,   // 与上一动画同时
    msoAnimTriggerAfterPrevious: 3,  // 上一动画之后

    // MsoAnimateByLevel
    msoAnimateLevelNone: 0,

    // MsoLineDashStyle
    msoLineDash: 4,

    // MsoZOrderCmd
    msoBringToFront: 0,
    msoSendToBack: 1,

    // Office 颜色（BGR 编码: R + G*256 + B*65536）
    RGB_RED: 255,          // RGB(255,0,0)
    RGB_BLUE: 16711680,    // RGB(0,0,255)
    RGB_GREEN: 65280,      // RGB(0,255,0) 标记点C
    RGB_ORANGE: 42495,     // RGB(255,165,0) 标记点D
    RGB_PURPLE: 8388864,   // RGB(128,0,128) 顶点P/弧点M
    RGB_CYAN: 16776960,    // RGB(0,255,255) 航点W
    RGB_MAGENTA: 16711935  // RGB(255,0,255) 描迹T
};
