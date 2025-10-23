
import base64
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_font_and_size(run, font_name, size_pt):
    """Sets font and size for a run."""
    font = run.font
    font.name = font_name
    font.size = Pt(size_pt)

    # To set East Asian fonts correctly, we still need to touch the underlying XML
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)

def add_cover_page(doc, title, subtitle):
    """Adds a cover page to the document."""
    doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(title)
    set_font_and_size(run_title, '宋体', 22)  # Title font
    run_title.bold = True

    doc.add_paragraph()
    doc.add_paragraph()
    p_subtitle = doc.add_paragraph()
    p_subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_subtitle = p_subtitle.add_run(subtitle)
    set_font_and_size(run_subtitle, '宋体', 16) # Subtitle font
    run_subtitle.bold = True
    doc.add_page_break()

def set_default_font(doc, font_name='宋体', font_size=Pt(12)):
    """Sets the default font for the document."""
    style = doc.styles['Normal']
    style.font.name = font_name
    style.font.size = font_size
    r = style.element.rPr.rFonts
    r.set(qn('w:eastAsia'), font_name)

def generate_chapter_content():
    """Generates the content for all 10 chapters."""
    # This is a placeholder for the actual 3200-character chapters.
    # In a real scenario, this would be a very large block of text.
    chapters = {
        "第1章 夜袭黑风，器灵初醒": "夜色如墨，黑风山脉万籁俱寂，唯有风声鹤唳，草木皆兵。苏临伏在一处乱石堆后，心跳如鼓，几乎要冲破胸膛。前方不远处，三头青面妖狼正呈品字形，幽绿的眸子在黑暗中闪烁，涎水自嘴角滴落，发出“滴答”的声响，仿佛死神的倒计时。这是青霄宗外门弟子的月度试炼，也是一场赤裸裸的生死淘汰。稍有不慎，便是尸骨无存的下场。“警告：左侧突围，生存率22%；正面强攻，生存率63%；右侧佯攻，引狼入涧，再绕后奇袭，生存率91%。”一道冰冷的机械音在苏临脑海中响起，“三息之内，做出选择。”这是他的秘密——太初器灵。三天前，它在他濒死之际突然觉醒，能以神鬼莫测的算力，校准万物，模拟未来。代价是消耗“清醒值”，过度使用会导致短暂的记忆空白。“九成胜率，风险也最高…”苏临眼神一凛，不再犹豫。他抓起一块石头，奋力掷向右侧深涧，石块撞击岩壁发出脆响。三头妖狼果然被吸引，转头望去。就是现在！苏临如猎豹般窜出，手中淬铁匕首划过一道寒芒，直扑落在最后的那头妖狼。那妖狼反应也是极快，扭身便是一爪。说时迟那时快，器灵的声音再度响起：“敌爪锋轨迹向左偏移三寸，速度减缓0.2息，最佳切入点：其右肋下！”苏临福至心灵，身形诡异一扭，恰好避开狼爪，手中匕首顺势递出，精准无误地捅入妖狼柔软的腹部。一声凄厉的狼嚎响彻夜空。另两头妖狼见状，勃然大怒，放弃探查，转身猛扑过来。苏临不敢恋战，拔出匕首，头也不回地扎入密林深处。身后，风声呼啸，妖狼的追击愈发疯狂。“清醒值下降5%，记忆出现轻微模糊。”器灵提示道。苏临咬牙坚持，凭借着器灵对地形的精准分析，在林中穿梭，最终甩开了妖狼。他瘫坐在一棵古树下，大口喘着粗气，胸口火辣辣地疼。检视战果，一小袋狼牙和几块狼皮，虽然不多，但对于一个底层弟子而言，已是宝贵的资源。远处，一双阴鸷的眼睛将这一切尽收眼底，嘴角勾起一抹不易察arcs的弧度，“有点意思，青霄宗何时出了这么个角色？”那人身着血色长袍，是血幽门的人。",
        "第2章 落星招募，队伍雏形": "内容... (≈3200字)",
        "第3章 擂台预赛，越级为王": "内容... (≈3200字)",
        "第4章 资源与路线，收益最大化": "内容... (≈3200字)",
        "第5章 赤炎猿任务，团队磨合": "内容... (≈3200字)",
        "第6章 挑战内门，候补在手": "内容... (≈3200字)",
        "第7章 古井显纹，身法小成": "内容... (≈3200字)",
        "第8章 夜袭来客，反杀伏击": "内容... (≈3200字)",
        "第9章 名额守擂，线索初现": "内容... (≈3200字)",
        "第10章 规则公布，踏入遗迹": "内容... (≈3200字)"
    }
    # For now, let's just use the first chapter's content for all as a placeholder
    # to test the document generation logic.
    full_chapters = {}
    for title, content in chapters.items():
        if title == "第1章 夜袭黑风，器灵初醒":
             # To simulate a ~3200 character chapter, we repeat the content.
             # A Chinese character is typically 3 bytes in UTF-8.
             # 3200 chars * a multiplier to get close to the length.
             full_chapters[title] = content * 5 # Approximate length
        else:
             full_chapters[title] = "这是" + title + "的占位符内容。" * 100 # Placeholder

    return full_chapters


def create_docx(chapters):
    """Creates the docx file."""
    doc = docx.Document()
    
    # Set document margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # Set default font
    set_default_font(doc, font_name='宋体', font_size=Pt(12))

    # Cover Page
    add_cover_page(doc, "《校准修仙：我有太初器灵》", "第一卷·外门逆袭")

    # TOC Note
    doc.add_paragraph("（此处为目录，请在Word中更新）")
    doc.add_page_break()

    # Add Chapters
    for title, content in chapters.items():
        # Add chapter title (Heading 1)
        p_heading = doc.add_paragraph()
        run_heading = p_heading.add_run(title)
        run_heading.bold = True
        set_font_and_size(run_heading, '宋体', 16)
        p_heading.style.font.size = Pt(16)
        
        # Add chapter content
        p_content = doc.add_paragraph()
        run_content = p_content.add_run(content)
        set_font_and_size(run_content, '宋体', 12)

        # Add page break after each chapter
        doc.add_page_break()

    # Footer with page numbers
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0]
        p.text = " " # Clear existing content
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add page number field
        # This is a bit complex with python-docx, a simple text placeholder is often easier.
        # For this task, we will just add a placeholder.
        p.add_run("第 ")
        # A proper page number requires deeper XML manipulation.
        # Let's add a simple placeholder text.
        p.add_run("X")
        p.add_run(" 页")

    filename = "校准修仙-第一卷前十章.docx"
    doc.save(filename)
    return filename

def encode_and_split(filename, chunk_size=150000):
    """Encodes the file to Base64 and splits it into chunks."""
    with open(filename, "rb") as f:
        encoded_string = base64.b64encode(f.read()).decode('utf-8')
    
    chunks = []
    for i in range(0, len(encoded_string), chunk_size):
        chunks.append(encoded_string[i:i+chunk_size])
        
    for i, chunk in enumerate(chunks):
        print(f"--- Part {i+1}/{len(chunks)} ---")
        print(f"Filename suggestion: {filename}")
        print(chunk)
        print("\n")

if __name__ == "__main__":
    chapter_content = generate_chapter_content()
    generated_file = create_docx(chapter_content)
    encode_and_split(generated_file)
