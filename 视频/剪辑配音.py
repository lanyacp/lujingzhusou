# -*- coding: utf-8 -*-
"""路径助手录屏：按功能点分段剪辑 + 中文配音 + 烧录字幕。

输入：桌面录屏「录制 11月12日 09.06.mp4」（43.7s / 1280x720 / 30fps / 静音）
输出：路径助手-非中心点路径动画-解说版.mp4

流程：每个镜头单独 TTS → 按配音时长裁/冻帧对齐画面 → 拼接 → 生成 ASS 字幕 → 合成。
运行: python 剪辑配音.py
"""
import asyncio
import re
import subprocess
import sys
from pathlib import Path

import edge_tts
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
SRC = Path(r'C:\Users\Administrator\Desktop\录制 11月12日 09.06.mp4')
WORK = ROOT / '_work'
WORK.mkdir(exist_ok=True)
OUT = ROOT / '路径助手-非中心点路径动画-解说版.mp4'
VOICE = 'zh-CN-YunxiNeural'
RATE = '+8%'
W, H, FPS = 1280, 720, 30
FONT_B = 'C:/Windows/Fonts/msyhbd.ttc'
FONT_N = 'C:/Windows/Fonts/msyh.ttc'

# 画面时间来自逐帧核对：0-7 展示窗格 / 7-17 放 C 点到顶点 /
# 17-27 放 D 点与红虚线预览 / 27-34 参数与生成按钮 / 34-43.7 生成成功与回读
SEGMENTS = [
    {'kind': 'card', 'card': 'intro',
     'text': '路径助手，怎么加一条不是中心点的路径动画，看一遍就会。'},
    {'kind': 'video', 'start': 0, 'end': 7,
     'text': '默认的路径动画按图形中心走。但三角形全等这类课，要的是顶点对顶点。'},
    {'kind': 'video', 'start': 7, 'end': 17,
     'text': '点「放C点」，把参照点拖到三角形顶点。普通形状没有可识别的顶点节点，得放大视图手动摆准。'},
    {'kind': 'video', 'start': 17, 'end': 27,
     'text': '再「放D点」放到目标位置。红色虚线是 C 到 D 的预览，位移按两点坐标差算，不用肉眼对。'},
    {'kind': 'video', 'start': 27, 'end': 34,
     'text': '时长、触发方式在窗格里设好，点「生成路径动画」。'},
    {'kind': 'video', 'start': 34, 'end': 43.7,
     'text': '生成后插件回读路径串，位移是坐标算出来的，不是眼睛估的。选中图形，绿点起点、红点终点。'},
    {'kind': 'card', 'card': 'outro',
     'text': '一个限制：普通形状的 C 点要手动摆准；带顶点的绘制图形，节点吸附能帮上忙。'},
]


def run(args, cwd=None):
    r = subprocess.run(args, cwd=cwd or WORK, capture_output=True,
                       text=True, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        raise RuntimeError('命令失败: ' + ' '.join(str(a) for a in args[:6]) +
                           '\n' + (r.stderr or '')[-1200:])
    return r.stdout


def duration(path):
    out = run(['ffprobe', '-v', 'error', '-show_entries', 'format=duration',
               '-of', 'default=nw=1:nk=1', str(path)])
    return float(out.strip())


async def tts(text, mp3):
    for attempt in range(3):
        try:
            c = edge_tts.Communicate(text, VOICE, rate=RATE)
            await c.save(str(mp3))
            if mp3.exists() and mp3.stat().st_size > 1000:
                return
        except Exception as e:
            if attempt == 2:
                raise
            print('  TTS 重试一次:', type(e).__name__)
            await asyncio.sleep(1.5)
    raise RuntimeError('TTS 未产出文件: ' + text[:20])


def grad_bg(w, h, c1, c2):
    im = Image.new('RGB', (w, h))
    d = ImageDraw.Draw(im)
    for y in range(h):
        t = y / max(1, h - 1)
        d.line([(0, y), (w, y)],
               fill=tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3)))
    return im


def make_card(kind, path):
    if kind == 'intro':
        im = grad_bg(W, H, (6, 78, 59), (16, 185, 129))
        d = ImageDraw.Draw(im)
        f1 = ImageFont.truetype(FONT_B, 78)
        f2 = ImageFont.truetype(FONT_B, 46)
        f3 = ImageFont.truetype(FONT_N, 28)
        f4 = ImageFont.truetype(FONT_N, 24)
        d.text((86, 180), '路径助手', font=f1, fill='#FFFFFF')
        d.rounded_rectangle((86, 292, 150, 300), radius=4, fill='#FDE68A')
        d.text((86, 330), '不是中心点的路径动画', font=f2, fill='#FDE68A')
        for i, line in enumerate(['① 放 C 点到顶点', '② 放 D 点对齐目标', '③ 生成并回读路径串']):
            d.text((88 + i * 370, 470), line, font=f3, fill='#ECFDF5')
        d.text((86, 630), 'WPS 演示 · 精准路径动画　/　v1.3.2', font=f4,
               fill='#C7F0E0')
    else:
        im = grad_bg(W, H, (17, 24, 39), (31, 58, 68))
        d = ImageDraw.Draw(im)
        f1 = ImageFont.truetype(FONT_B, 60)
        f2 = ImageFont.truetype(FONT_N, 36)
        f3 = ImageFont.truetype(FONT_N, 26)
        d.text((86, 150), '一个限制，先说清楚', font=f1, fill='#FDE68A')
        d.rounded_rectangle((86, 250, 150, 258), radius=4, fill='#059669')
        d.text((86, 306), '普通形状没有可识别的顶点节点，', font=f2, fill='#FFFFFF')
        d.text((86, 366), 'C 点需要放大视图手动摆准。', font=f2, fill='#FFFFFF')
        d.text((86, 458), '换成带顶点的绘制图形，', font=f2, fill='#A7F3D0')
        d.text((86, 518), '节点的自动吸附能帮上忙。', font=f2, fill='#A7F3D0')
        d.text((86, 632), '路径助手 v1.3.2　·　标记点 C→D 模式', font=f3, fill='#94A3B8')
    im.save(path)


def clip_video(seg, target, dur):
    """裁出画面区间，并按配音时长对齐：不足冻帧补足，过长裁掉尾部。"""
    src_len = seg['end'] - seg['start']
    raw = WORK / '_raw.mp4'
    run(['ffmpeg', '-hide_banner', '-loglevel', 'error', '-y',
         '-ss', str(seg['start']), '-t', f'{src_len:.3f}', '-i', str(SRC),
         '-vf', f'scale={W}:{H},fps={FPS},format=yuv420p',
         '-an', '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '20',
         '-r', str(FPS), str(raw)])
    actual = duration(raw)
    if dur > actual + 0.05:
        vf = (f'tpad=stop_mode=clone:stop_duration={dur - actual:.3f},'
              f'format=yuv420p')
    else:
        vf = 'format=yuv420p'
    run(['ffmpeg', '-hide_banner', '-loglevel', 'error', '-y', '-i', str(raw),
         '-vf', vf, '-t', f'{dur:.3f}', '-an',
         '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '20',
         '-r', str(FPS), '-pix_fmt', 'yuv420p', str(target)])


def clip_card(seg, target, dur):
    png = WORK / f'_card_{seg["card"]}.png'
    make_card(seg['card'], png)
    run(['ffmpeg', '-hide_banner', '-loglevel', 'error', '-y',
         '-loop', '1', '-framerate', str(FPS), '-i', str(png),
         '-t', f'{dur:.3f}', '-vf', f'scale={W}:{H},format=yuv420p',
         '-an', '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '20',
         '-r', str(FPS), '-pix_fmt', 'yuv420p', str(target)])


def split_caption(text, maxlen=19):
    """按标点切句，过长的再按字数切，保证单行字幕不溢出。"""
    raw = re.split(r'([，。；：！？])', text)
    chunks, buf = [], ''
    for part in raw:
        if not part:
            continue
        buf += part
        if part in '，。；：！？':
            chunks.append(buf.strip('，。；：！？ '))
            buf = ''
    if buf:
        chunks.append(buf.strip('，。；：！？ '))
    out = []
    for c in chunks:
        while len(c) > maxlen:
            out.append(c[:maxlen])
            c = c[maxlen:]
        if c:
            out.append(c)
    return out


def ts(sec):
    sec = max(0.0, sec)
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = sec % 60
    return f'{h}:{m:02d}:{s:05.2f}'


def ts_srt(sec):
    sec = max(0.0, sec)
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = sec % 60
    return f'{h:02d}:{m:02d}:{s:06.3f}'.replace('.', ',')


ASS_HEAD = '''[Script Info]
ScriptType: v4.00+
PlayResX: 1280
PlayResY: 720
WrapStyle: 2
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,Microsoft YaHei,40,&H00FFFFFF,&H000000FF,&H00000000,&H80000000,-1,0,0,0,100,100,0,0,3,4,1,2,40,40,48,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
'''


def main():
    if not SRC.exists():
        raise RuntimeError('找不到录屏: ' + str(SRC))
    vlist, alist, events = [], [], []
    offset = 0.0
    for i, seg in enumerate(SEGMENTS):
        mp3 = WORK / f'voice_{i:02d}.mp3'
        wav = WORK / f'voice_{i:02d}.wav'
        segv = WORK / f'seg_{i:02d}.mp4'
        print(f'[{i + 1}/{len(SEGMENTS)}] 配音: {seg["text"][:24]}…', flush=True)
        asyncio.run(tts(seg['text'], mp3))
        dur = duration(mp3)
        run(['ffmpeg', '-hide_banner', '-loglevel', 'error', '-y',
             '-i', str(mp3), '-ar', '48000', '-ac', '2',
             '-c:a', 'pcm_s16le', str(wav)])
        alist.append(wav.name)
        print(f'        时长 {dur:.2f}s，处理画面…', flush=True)
        if seg['kind'] == 'card':
            clip_card(seg, segv, dur)
        else:
            clip_video(seg, segv, dur)
        vlist.append(segv.name)
        if seg['kind'] == 'video':
            lines = split_caption(seg['text'])
            total = sum(len(x) for x in lines) or 1
            t = offset
            for line in lines:
                d = dur * len(line) / total
                events.append((t, t + d, line))
                t += d
        offset += dur
        print(f'        累计 {offset:.2f}s', flush=True)

    (WORK / 'vlist.txt').write_text(
        ''.join(f"file '{n}'\n" for n in vlist), encoding='utf-8')
    (WORK / 'alist.txt').write_text(
        ''.join(f"file '{n}'\n" for n in alist), encoding='utf-8')
    ass = ASS_HEAD + ''.join(
        f'Dialogue: 0,{ts(a)},{ts(b)},Default,,0,0,0,,{line}\n'
        for a, b, line in events)
    (WORK / 'sub.ass').write_text(ass, encoding='utf-8')
    srt = ''.join(f'{k}\n{ts_srt(a)} --> {ts_srt(b)}\n{line}\n\n'
                  for k, (a, b, line) in enumerate(events, 1))
    (ROOT / '字幕.srt').write_text(srt, encoding='utf-8')

    print('拼接画面与配音…', flush=True)
    run(['ffmpeg', '-hide_banner', '-loglevel', 'error', '-y',
         '-f', 'concat', '-safe', '0', '-i', 'vlist.txt',
         '-c', 'copy', 'concat_v.mp4'])
    run(['ffmpeg', '-hide_banner', '-loglevel', 'error', '-y',
         '-f', 'concat', '-safe', '0', '-i', 'alist.txt',
         '-c', 'copy', 'voice.wav'])

    print('烧录字幕并合成成品…', flush=True)
    run(['ffmpeg', '-hide_banner', '-loglevel', 'error', '-y',
         '-i', 'concat_v.mp4', '-i', 'voice.wav',
         '-vf', 'ass=sub.ass', '-map', '0:v:0', '-map', '1:a:0',
         '-c:v', 'libx264', '-preset', 'medium', '-crf', '20',
         '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '192k',
         '-shortest', '-movflags', '+faststart', str(OUT)])
    print(f'\n完成: {OUT}')
    print(f'总时长 {duration(OUT):.2f}s，字幕 {len(events)} 条，'
          f'大小 {OUT.stat().st_size / 1024 / 1024:.1f} MB')


if __name__ == '__main__':
    main()
