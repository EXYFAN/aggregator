# -*- coding: utf-8 -*-

# Programmatically generate the .docx for the novel plan and output Base64 (chunked)
# Usage: python tools/generate_docx_xiuxian.py

import base64
import io
import math
import os
from typing import List

from docx import Document
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


TITLE = "校准修仙：我有太初器灵"
SUBTITLE = "150万字计划书（书名＋梗概＋卷纲）"
FILENAME_SUGGESTION = "校准修仙-总纲.docx"

SYNOPSIS = (
    "边城少年苏临携父亲遗物裂玉入青霄宗。外门试炼夜，裂玉中“太初器灵”苏醒，可校准选择、模拟对局、预警风险，却不代替出手。"
    "苏临与阵法天才慕青禾、体修猛将林折、炼丹师白砚结成四人小队，战场上以“协同校准”滚雪球：擂台越级、副本破围、任务暴富，拿下太初遗迹名额。"
    "血幽门与上宗黑手暗中争夺主控碎片，镇魔司内部守序派与极端派分裂，父亲“战死”背后牵连主控转移计划。"
    "小队从外门到内门、从联赛到联军战，既有个人高光也有群像热血：阵盘封控、肉身突击、药道翻盘、剑修斩首。"
    "苏临在规则压制中完成跃迁，逐步接近真相——上界极端派意图以“秩序重织”凌驾万民。他将以校准之道，与伙伴并肩，重启公正秩序。"
)

CAST = [
    ("苏临（队长/主攻）", "冷静果断，守底线敢冒险；强项最优解决策、临场反打。口头禅：“先把代价算清。”"),
    ("慕青禾（阵法/控场）", "理性细致、微洁癖；擅阵盘与地形改造。口头禅：“把变量锁死，胜率自来。”"),
    ("林折（体修/前排）", "直率重义、抗压怪；破阵开路、以伤换胜。口头禅：“我扛三息，你们结束。”"),
    ("白砚（炼丹/后勤）", "稳健谨慎、记忆超群；药剂、资源滚雪球、情报官。口头禅：“钱要花在胜率上。”"),
    ("陆惊澜（剑修/斩首，后入队）", "骄而坦诚；由对手转盟友，外放斩首手。口头禅：“剑到即分胜负。”"),
]

VOLUMES = [
    (
        "第一卷 外门逆袭",
        [
            "舞台：青霄宗外门、落星城、黑风山脉",
            "目标：外门考核夺太初遗迹候补名额",
            "爆点：器灵苏醒；擂台越级第一；副本反杀血幽门小队",
            "角色高光：苏临最优解指挥成型；林折以伤换胜；青禾三阵连控；白砚用药翻盘",
            "卷末：父亲线索初现，炼气大成",
        ],
    ),
    (
        "第二卷 古井初开",
        [
            "舞台：太初古井外围/第一层",
            "目标：积分换钥，开二层临时权限",
            "爆点：身法被“校准”至小成；短时模拟破围；暴富灵石",
            "高光：团队首次“协同校准”；陆惊澜以剑破局但暂时对立",
            "卷末：筑基成功，拿二层通行",
        ],
    ),
    (
        "第三卷 宗门风云",
        [
            "舞台：青霄宗内门、落星城拍卖行",
            "目标：高阶资源线、贡献权限",
            "冲突：内门新秀围堵、暗中黑手搅局",
            "爆点：功法个性化校准；拍卖截胡；越阶击败天才",
            "卷末：成为宗门探队核心，队伍编制获批",
        ],
    ),
    (
        "第四卷 边关风暴",
        [
            "舞台：边关秘境群、镇魔司哨所",
            "目标：平定妖潮换机密阅览权",
            "爆点：战阵协同初成，以阵破潮；军功在身",
            "高光：林折扛线开路；青禾构筑“口袋阵地”；白砚军需体系搭建",
            "卷末：金丹在望；父亲卷宗出现“主控转移”线索",
        ],
    ),
    (
        "第五卷 天骄榜路",
        [
            "舞台：中域大赛、跨宗联赛",
            "目标：跻身天骄榜、夺跨宗队长",
            "爆点：基础剑法进化专属剑诀；公开越品阶胜强敌",
            "高光：陆惊澜客卿入队，联手斩首；苏临多线分推",
            "卷末：金丹小成，中域扬名",
        ],
    ),
    (
        "第六卷 古井核心",
        [
            "舞台：古井第二/第三层",
            "目标：夺主控碎片、修复器灵记忆",
            "爆点：“风险预演”模块解锁；绝境团灭围堵联队",
            "高光：青禾阵图链接器灵算力；白砚以药诱敌误判",
            "卷末：金丹大成→元婴门前；器灵记忆≈5%",
        ],
    ),
    (
        "第七卷 中域大会",
        [
            "舞台：中域盟会、丹阵大会",
            "目标：正式探索者资格、情报网搭建",
            "冲突：学术派系内斗、暗杀与反制",
            "爆点：丹阵“双修校准”出圈；大会会战封神",
            "卷末：元婴小成；获官方半授权“探索团”",
        ],
    ),
    (
        "第八卷 上宗棋局",
        [
            "舞台：上宗秘境、九大世家",
            "目标：跨域通行权、父亲封存线索",
            "爆点：反将上宗强者为棋；一役震动上宗",
            "高光：苏临博弈校准逼对手自崩；陆惊澜转正式队员",
            "卷末：化神在望；确认父亲被封于“主核备用仓”",
        ],
    ),
    (
        "第九卷 上界风霜",
        [
            "舞台：半开上界、断层天梯",
            "目标：进入上界前沿、适配规则压制",
            "冲突：上界代理势力、规则不利环境",
            "爆点：规则适配校准；压制下夺城；父亲苏醒前夜",
            "卷末：化神稳固；掌握“秩序重织”线索",
        ],
    ),
    (
        "第十卷 断层真相",
        [
            "舞台：主控遗址、守门人记忆库",
            "目标：拼合真相、拟定对抗方案",
            "爆点：父子并肩；器灵记忆大幅回归；九幽教现身",
            "高光：白砚情报拼图揭幕后；青禾阵枢护送父亲回归",
            "卷末：合道前兆→小成；定三阶段反制",
        ],
    ),
    (
        "第十一卷 重启与对赌",
        [
            "舞台：诸域联军、上界前线",
            "目标：收拢权限碎片、实施温和重启",
            "冲突：与极端派对赌“谁的秩序能存活”",
            "爆点：跨域联战；协同校准极限拉满",
            "卷末：合道稳固；器灵记忆≈90%，飞升在即",
        ],
    ),
    (
        "第十二卷 校准诸天",
        [
            "舞台：上界主控层、中枢天庭遗址",
            "目标：最终对决与秩序重建",
            "爆点：群像大合唱；主角以校准权统摄诸道；父子线收束",
            "结局：飞升圆满；太初遗迹转为公域试炼场",
        ],
    ),
]


def set_styles(doc: Document) -> None:
    # Normal style: SimSun 12pt
    normal = doc.styles["Normal"]
    normal.font.name = "SimSun"
    normal.font.size = Pt(12)
    # East Asia font mapping for Normal
    r = normal._element.get_or_add_rPr()
    rFonts = r.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "SimSun")

    # Heading 1: SimHei 16pt (bold)
    h1 = doc.styles["Heading 1"]
    h1.font.name = "SimHei"
    h1.font.size = Pt(16)
    h1.font.bold = True
    r = h1._element.get_or_add_rPr()
    rFonts = r.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "SimHei")

    # Heading 2: SimHei 14pt (bold)
    h2 = doc.styles["Heading 2"]
    h2.font.name = "SimHei"
    h2.font.size = Pt(14)
    h2.font.bold = True
    r = h2._element.get_or_add_rPr()
    rFonts = r.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "SimHei")


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    run._r.append(fldChar1)

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = " PAGE "
    run._r.append(instrText)

    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    run._r.append(fldChar2)

    run2 = paragraph.add_run()
    run2.text = "1"

    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar3)


def build_document() -> bytes:
    doc = Document()

    # Page setup: A4, margins 2.5 cm
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    # Ensure footer page numbers are centered will be handled below for all sections

    # Fonts and styles
    set_styles(doc)

    # Cover page
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run(TITLE)
    run.font.name = "SimHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
    run.font.size = Pt(20)
    run.bold = True

    doc.add_paragraph("")

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_sub.add_run(SUBTITLE)
    run.font.name = "SimHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
    run.font.size = Pt(14)

    # Page break to body
    doc.add_page_break()

    # Add footer page numbers for the new section if a new section was created
    # python-docx page_break doesn't create a new section; ensure footer exists regardless
    for sec in doc.sections:
        if not sec.footer.paragraphs:
            sec.footer.add_paragraph("")
        para = sec.footer.paragraphs[0]
        if len(para.runs) == 0:
            add_page_number(para)

    # TOC hint
    doc.add_paragraph("目录（打开Word后使用“更新域/目录”生成）")

    # Body content with Heading styles
    doc.add_heading("书名", level=1)
    doc.add_paragraph(TITLE)

    doc.add_heading("300字梗概", level=1)
    doc.add_paragraph(SYNOPSIS)

    doc.add_heading("主角团画像（个性鲜明）", level=1)
    for name, desc in CAST:
        doc.add_heading(name, level=2)
        doc.add_paragraph(desc)

    doc.add_heading("十二卷卷纲（约150万字，≈480章）", level=1)
    for vol_title, points in VOLUMES:
        doc.add_heading(vol_title, level=2)
        for pt in points:
            # Use dash-leading lines to match exact content
            doc.add_paragraph(f"- {pt}")

    # Save to bytes
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def b64_chunks(data: bytes, chunk_size: int = 150_000) -> List[str]:
    b64 = base64.b64encode(data).decode("utf-8")
    return [b64[i : i + chunk_size] for i in range(0, len(b64), chunk_size)]


def main():
    data = build_document()
    chunks = b64_chunks(data)
    total = len(chunks)

    print("Filename suggestion: {}".format(FILENAME_SUGGESTION))
    print("How to use:")
    print("1) Concatenate the Base64 parts in order into a single string.")
    print("2) Base64-decode to get the .docx binary.")
    print("3) Save as '{}' and open in Word; then update the Table of Contents if desired.".format(FILENAME_SUGGESTION))
    print("")

    for idx, part in enumerate(chunks, start=1):
        print("==== Base64 Part {}/{} ====".format(idx, total))
        print(part)
        print("")


if __name__ == "__main__":
    main()
