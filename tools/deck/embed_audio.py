# -*- coding: utf-8 -*-
"""把逐页配音嵌入 PPTX（OOXML 层面直接改包）。

python-pptx 没有音频 API，所以这里在 zip 层面操作：
  1. ppt/media/ 放入 mp3 与海报图
  2. slideN.xml.rels 增加 audio / image 两条关系
  3. slideN.xml 的 spTree 里插入 <p:pic>（nvPr 内是 <a:audioFile>）
  4. [Content_Types].xml 补 mp3 的 Default

--autoplay 会额外写入 <p:timing>，让音频随页面自动播放；
该结构 PowerPoint 校验较严，默认不写，产出的是「点喇叭图标播放」的稳妥版本。

运行：python embed_audio.py [--autoplay] [--mix]
"""

import argparse
import copy
import json
import os
import shutil
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
AUDIO = os.path.join(OUT, "audio")

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
RT_AUDIO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

POSTER = "ppt/media/audio-poster.png"

PIC_TMPL = """<p:pic xmlns:p="%(p)s" xmlns:a="%(a)s" xmlns:r="%(r)s">
  <p:nvPicPr>
    <p:cNvPr id="%(id)s" name="%(name)s">
      <a:hlinkClick xmlns:r="%(r)s" r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr>
      <a:audioFile r:link="%(rid_audio)s"/>
      <p:extLst>
        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
          <p14:media xmlns:p14="%(p14)s" r:embed="%(rid_img)s"/>
        </p:ext>
      </p:extLst>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="%(rid_img)s"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="%(x)s" y="%(y)s"/><a:ext cx="%(cx)s" cy="%(cy)s"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""

# 幻灯片尺寸 13.333 x 7.5 英寸 -> EMU
SLIDE_W, SLIDE_H = 12192000, 6858000
ICON = 320000                       # 约 0.33 英寸
ICON_X, ICON_Y = SLIDE_W - 640000, SLIDE_H - 1900000


def timing_xml(rid_audio):
    """音频随页面自动播放的时间轴（mainSeq + click 分组内 delay=0 播放命令）。"""
    return """<p:timing xmlns:p="%(p)s" xmlns:a="%(a)s" xmlns:r="%(r)s">
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <p:par>
                  <p:cTn id="3" fill="hold">
                    <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="4" fill="hold">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:par>
                              <p:cTn id="5" presetID="1" presetClass="aud" presetSubtype="0" fill="hold" nodeType="withEffect">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                <p:childTnLst>
                                  <p:cmd type="play" cmd="play" r:id="%(rid)s"/>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                  <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                  <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
                </p:par>
              </p:childTnLst>
            </p:cTn>
            <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
            <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
</p:timing>""" % dict(rid=rid_audio, **NS)


def next_rid(rels_root):
    used = {rel.get("Id") for rel in rels_root}
    i = 1
    while ("rId%d" % i) in used:
        i += 1
    return "rId%d" % i


def add_rel(rels_root, rid, rtype, target):
    el = etree.SubElement(rels_root, "{%s}Relationship" % NS["rel"])
    el.set("Id", rid)
    el.set("Type", rtype)
    el.set("Target", target)
    return rid


def ensure_default_ct(ct_root, ext, ctype):
    for d in ct_root.findall("{%s}Default" % NS["ct"]):
        if d.get("Extension", "").lower() == ext:
            return
    el = etree.Element("{%s}Default" % NS["ct"])
    el.set("Extension", ext)
    el.set("ContentType", ctype)
    ct_root.insert(0, el)


def embed(src_pptx, dst_pptx, manifest, autoplay=False):
    media = {}          # zip 内新增部件名 -> 磁盘文件
    slides = len(manifest["slides"])

    with zipfile.ZipFile(src_pptx) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    # 海报图（喇叭图标）
    poster = os.path.join(OUT, "audio", "_poster.png")
    make_poster(poster)
    media[POSTER] = poster

    for entry in manifest["slides"]:
        idx = entry["index"]
        media_name = "ppt/media/narration%02d.mp3" % idx
        media[media_name] = os.path.join(AUDIO, entry["audio"])

    for entry in manifest["slides"]:
        idx = entry["index"]
        slide_part = "ppt/slides/slide%d.xml" % idx
        rels_part = "ppt/slides/_rels/slide%d.xml.rels" % idx
        if slide_part not in data:
            continue

        rels = etree.fromstring(data[rels_part])
        rid_audio = next_rid(rels)
        add_rel(rels, rid_audio, RT_AUDIO,
                "../media/narration%02d.mp3" % idx)
        rid_img = next_rid(rels)
        add_rel(rels, rid_img, RT_IMAGE, "../media/audio-poster.png")
        data[rels_part] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True)

        sld = etree.fromstring(data[slide_part])
        sp_tree = sld.find(".//{%s}cSld/{%s}spTree" % (NS["p"], NS["p"]))
        if sp_tree is None:
            continue

        max_id = 1
        for cnv in sp_tree.iter("{%s}cNvPr" % NS["p"]):
            try:
                max_id = max(max_id, int(cnv.get("id")))
            except (TypeError, ValueError):
                pass

        pic = etree.fromstring(PIC_TMPL % dict(
            id=max_id + 1, name=entry["audio"], rid_audio=rid_audio,
            rid_img=rid_img, x=ICON_X, y=ICON_Y, cx=ICON, cy=ICON,
            p14=P14, **NS))
        sp_tree.append(pic)

        if autoplay:
            old = sld.find("{%s}timing" % NS["p"])
            if old is not None:
                sld.remove(old)
            sld.append(etree.fromstring(timing_xml(rid_audio)))

        data[slide_part] = etree.tostring(
            sld, xml_declaration=True, encoding="UTF-8", standalone=True)

    ct = etree.fromstring(data["[Content_Types].xml"])
    ensure_default_ct(ct, "mp3", "audio/mpeg")
    ensure_default_ct(ct, "png", "image/png")
    data["[Content_Types].xml"] = etree.tostring(
        ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(dst_pptx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:                      # 原部件保持原顺序
            zout.writestr(name, data[name])
        for part, path in media.items():
            zout.writestr(part, open(path, "rb").read())

    return dst_pptx, slides


def make_poster(path):
    """生成喇叭图标作为媒体海报图。"""
    from PIL import Image, ImageDraw
    n = 256
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, n - 1, n - 1], radius=56,
                        fill=(0x14, 0x1C, 0x38, 255),
                        outline=(0x22, 0x30, 0x5C, 255), width=6)
    d.polygon([(96, 104), (128, 104), (166, 74), (166, 182), (128, 152),
               (96, 152)], fill=(0x22, 0xD3, 0xEE, 255))
    for i, r in enumerate((22, 38)):
        d.arc([166 - r, 128 - r, 166 + r, 128 + r], start=-58, end=58,
              fill=(0x9A, 0xA6, 0xC8, 255), width=7)
    img.save(path)


def validate(path):
    """结构自检：所有部件可解析、引用的 r:id 都存在、媒体非空。"""
    problems = []
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        if bad:
            problems.append("zip 损坏: %s" % bad)
        ct = etree.fromstring(z.read("[Content_Types].xml"))
        exts = {d.get("Extension", "").lower()
                for d in ct.findall("{%s}Default" % NS["ct"])}
        if "mp3" not in exts:
            problems.append("[Content_Types].xml 缺少 mp3 声明")

        names = z.namelist()
        n_audio = 0
        for name in names:
            if not name.endswith(".xml") and not name.endswith(".rels"):
                if name.startswith("ppt/media/"):
                    if z.getinfo(name).file_size == 0:
                        problems.append("空媒体: %s" % name)
                continue
            try:
                root = etree.fromstring(z.read(name))
            except etree.XMLSyntaxError as e:
                problems.append("XML 解析失败 %s: %s" % (name, e))
                continue

            if not name.startswith("ppt/slides/_rels/"):
                continue
            slide = name.replace("_rels/", "").replace(".rels", "")
            sld = etree.fromstring(z.read(slide))
            have = {r.get("Id") for r in root}
            for el in sld.iter():
                for attr in ("{%s}id" % NS["r"], "{%s}embed" % NS["r"],
                             "{%s}link" % NS["r"]):
                    v = el.get(attr)
                    if v and v not in have:
                        problems.append("%s 引用了不存在的 %s" % (slide, v))
            if any(el.tag == "{%s}audioFile" % NS["a"] for el in sld.iter()):
                n_audio += 1
        if n_audio == 0:
            problems.append("没有任何页面嵌入音频")
    return problems, n_audio


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--autoplay", action="store_true")
    ap.add_argument("--src", default=os.path.join(OUT, "dxxjudges-intro.pptx"))
    args = ap.parse_args()

    with open(os.path.join(AUDIO, "manifest.json"), encoding="utf-8") as f:
        manifest = json.load(f)

    name = "dxxjudges-intro-autoplay.pptx" if args.autoplay \
        else "dxxjudges-intro-audio.pptx"
    dst = os.path.join(OUT, name)
    embed(args.src, dst, manifest, autoplay=args.autoplay)

    issues, n = validate(dst)
    print("已生成 %s  (%.1f KB, %d 页带音频, autoplay=%s)"
          % (os.path.basename(dst), os.path.getsize(dst) / 1024, n, args.autoplay))
    if issues:
        for i in issues:
            print("  ! " + i)
        raise SystemExit(1)
    print("结构自检通过")