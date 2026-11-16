#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把自定义 Ribbon (customUI14.xml) 注入 PowerPoint 加载项 .ppam。

用法: python tools/inject-ppt-ribbon.py
读取 tools/ppt-src/customUI14.xml, 写入 PowerPoint版/PathHelper.ppam。
幂等: 重复运行会先移除旧的 customUI part 再注入。
"""
import os
import re
import shutil
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_XML = os.path.join(ROOT, "tools", "ppt-src", "customUI14.xml")
PPAM = os.path.join(ROOT, "PowerPoint版", "PathHelper.ppam")

CUSTOM_PART = "customUI/customUI14.xml"
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
CT_TYPE = "application/vnd.ms-office.customUI14+xml"
# 历史上误用过的错误关系类型, 注入时一并清除
LEGACY_REL_TYPES = [
    "http://schemas.microsoft.com/office/2009/07/customui",
    REL_TYPE,
]


def inject(path):
    with open(SRC_XML, "r", encoding="utf-8") as f:
        xml = f.read()

    with zipfile.ZipFile(path, "r") as z:
        entries = [(i, z.read(i.filename)) for i in z.infolist()]

    out = []
    for info, data in entries:
        name = info.filename
        if name == CUSTOM_PART:
            continue  # 旧的 customUI, 跳过
        if name == "_rels/.rels":
            text = data.decode("utf-8")
            # 移除旧的 customUI 关系(含历史上误用的类型)
            for legacy in LEGACY_REL_TYPES:
                text = re.sub(
                    r'<Relationship[^>]*Type="%s"[^>]*/>' % re.escape(legacy),
                    "", text,
                )
            rid = "rIdRibbon"
            text = text.replace(
                "</Relationships>",
                '<Relationship Id="%s" Type="%s" Target="%s"/></Relationships>'
                % (rid, REL_TYPE, CUSTOM_PART),
            )
            data = text.encode("utf-8")
        elif name == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = re.sub(
                r'<Override[^>]*PartName="/%s"[^>]*/>' % re.escape(CUSTOM_PART),
                "", text,
            )
            text = text.replace(
                "</Types>",
                '<Override PartName="/%s" ContentType="%s"/></Types>'
                % (CUSTOM_PART, CT_TYPE),
            )
            data = text.encode("utf-8")
        out.append((info, data))

    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            z.writestr(zi, data)
        z.writestr(CUSTOM_PART, xml.encode("utf-8"))
    shutil.move(tmp, path)
    print("Ribbon 已注入:", path)


if __name__ == "__main__":
    inject(PPAM)
