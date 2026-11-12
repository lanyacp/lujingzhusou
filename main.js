/**
 * main.js — 按顺序引入加载项各脚本
 * 注意: 使用 document.write 保证在文档解析阶段同步加载
 */
document.write('<script type="text/javascript" src="./js/enums.js"><\/script>');
document.write('<script type="text/javascript" src="./js/util.js"><\/script>');
document.write('<script type="text/javascript" src="./js/core.js"><\/script>');
document.write('<script type="text/javascript" src="./js/ribbon.js"><\/script>');
