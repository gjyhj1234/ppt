"""
generate_pptx.py
Converts the 轻松牙医 × 松佰供应链 HTML presentation (index.html) to a
faithful PPTX file (presentation.pptx) using python-pptx.

Run with:  python3 generate_pptx.py
"""

from pptx import Presentation
from pptx.util import Mm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Slide dimensions: A4 landscape 297 × 210 mm ──────────────────────────────
W = Mm(297)
H = Mm(210)

# ── Brand colours ─────────────────────────────────────────────────────────────
BLUE       = RGBColor(0x15, 0x65, 0xC0)
BLUE_MID   = RGBColor(0x19, 0x76, 0xD2)
BLUE_LIGHT = RGBColor(0xE3, 0xF0, 0xFF)
TEAL       = RGBColor(0x00, 0x89, 0x7B)
TEAL_LIGHT = RGBColor(0xE0, 0xF5, 0xF3)
GOLD       = RGBColor(0xE6, 0xA8, 0x17)
TEXT_DARK  = RGBColor(0x1A, 0x23, 0x40)
TEXT_MID   = RGBColor(0x37, 0x45, 0x69)
TEXT_MUTED = RGBColor(0x6C, 0x7A, 0x9C)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
RED        = RGBColor(0xE5, 0x39, 0x35)
AMBER      = RGBColor(0xFB, 0x8C, 0x00)
GREEN_D    = RGBColor(0x2E, 0x7D, 0x32)
RED_D      = RGBColor(0xC6, 0x28, 0x28)
PURPLE     = RGBColor(0x7B, 0x1F, 0xA2)

# ── Helper constants ──────────────────────────────────────────────────────────
HEADER_H = Mm(11)
FOOTER_H = Mm(8)
BODY_Y   = HEADER_H
BODY_H   = H - HEADER_H - FOOTER_H


# ── Low-level helpers ─────────────────────────────────────────────────────────

def hex2rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_rect(slide, x, y, w, h, fill=None, line_color=None, line_width_pt=0):
    """Add a plain rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        x, y, w, h
    )
    shape.line.fill.background()
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width_pt if line_width_pt else 0.75)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, x, y, w, h, text, font_size=12, bold=False,
                color=None, align=PP_ALIGN.LEFT, wrap=True,
                font_name="Microsoft YaHei"):
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    if color:
        run.font.color.rgb = color
    return txb


def add_textbox_multi(slide, x, y, w, h, paragraphs,
                      font_name="Microsoft YaHei"):
    """
    paragraphs: list of dicts:
      {text, font_size, bold, color, align, space_before (pt), space_after (pt)}
    """
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf = txb.text_frame
    tf.word_wrap = True
    for i, pd in enumerate(paragraphs):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = pd.get("align", PP_ALIGN.LEFT)
        if pd.get("space_before"):
            p.space_before = Pt(pd["space_before"])
        if pd.get("space_after"):
            p.space_after = Pt(pd["space_after"])
        run = p.add_run()
        run.text = pd["text"]
        run.font.size = Pt(pd.get("font_size", 12))
        run.font.bold = pd.get("bold", False)
        run.font.name = font_name
        if pd.get("color"):
            run.font.color.rgb = pd["color"]
    return txb


# ── Shared slide chrome (header + footer) ────────────────────────────────────

def add_header(slide, page_label, bg=BLUE, brand_color=WHITE, num_color=None):
    add_rect(slide, 0, 0, W, HEADER_H, fill=bg)
    add_textbox(slide, Mm(14), Mm(2.5), Mm(140), Mm(7),
                "轻松牙医 · 松佰供应链",
                font_size=13, bold=True, color=brand_color)
    add_textbox(slide, W - Mm(30), Mm(2.5), Mm(22), Mm(7),
                page_label, font_size=11,
                color=num_color or RGBColor(0xCC, 0xD9, 0xEE),
                align=PP_ALIGN.RIGHT)


def add_footer(slide, text="严格保密 · 仅供内部使用",
               left_color=BLUE, right_color=TEAL, txt_color=None):
    # gradient-ish footer: split into two rects for a rough gradient
    add_rect(slide, 0, H - FOOTER_H, W // 2, FOOTER_H, fill=left_color)
    add_rect(slide, W // 2, H - FOOTER_H, W - W // 2, FOOTER_H, fill=right_color)
    add_textbox(slide, Mm(14), H - FOOTER_H + Mm(1), Mm(200), Mm(6),
                text, font_size=9,
                color=txt_color or RGBColor(0xCC, 0xE8, 0xE5))


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 1 — WELCOME (cover)
# ─────────────────────────────────────────────────────────────────────────────

def build_slide1(slide):
    # Background gradient (simulated with two overlapping rects)
    add_rect(slide, 0, 0, W, H, fill=RGBColor(0x0D, 0x3B, 0x7A))
    # Subtle teal stripe on right
    add_rect(slide, W * 2 // 3, 0, W // 3, H, fill=RGBColor(0x00, 0x69, 0x5C))

    add_header(slide, "1 / 6",
               bg=RGBColor(0x0D, 0x3B, 0x7A),
               brand_color=RGBColor(0xCC, 0xDD, 0xFF),
               num_color=RGBColor(0x88, 0xAA, 0xCC))

    # Top label
    add_textbox(slide, 0, Mm(42), W, Mm(8),
                "STRATEGIC PARTNERSHIP PROPOSAL",
                font_size=10, color=RGBColor(0xAA, 0xBB, 0xCC),
                align=PP_ALIGN.CENTER)

    # Main title
    add_textbox(slide, 0, Mm(52), W, Mm(22),
                "携手共赢", font_size=42, bold=True,
                color=WHITE, align=PP_ALIGN.CENTER)

    add_textbox(slide, 0, Mm(72), W, Mm(18),
                "共启口腔新局", font_size=36, bold=True,
                color=RGBColor(0x7D, 0xD3, 0xF7), align=PP_ALIGN.CENTER)

    # Gold divider line
    add_rect(slide, W // 2 - Mm(30), Mm(93), Mm(60), Mm(2), fill=GOLD)

    # Subtitle
    add_textbox(slide, 0, Mm(97), W, Mm(10),
                "松佰供应链战略合作推广提案",
                font_size=15, color=RGBColor(0xDD, 0xEE, 0xFF),
                align=PP_ALIGN.CENTER)

    # Meta row
    meta = [("演讲人", "轻松牙医软件"),
            ("议题",   "供应链合作推广方案"),
            ("会议",   "集团总经理会")]
    col_w = Mm(50)
    start_x = (W - col_w * 3) // 2
    for i, (lbl, val) in enumerate(meta):
        cx = start_x + col_w * i
        add_textbox(slide, cx, Mm(148), col_w, Mm(6),
                    lbl, font_size=9,
                    color=RGBColor(0x88, 0x99, 0xAA),
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, cx, Mm(155), col_w, Mm(8),
                    val, font_size=11, bold=True,
                    color=RGBColor(0xDD, 0xEE, 0xFF),
                    align=PP_ALIGN.CENTER)

    add_footer(slide, "严格保密 · 仅供内部使用",
               left_color=RGBColor(0x0D, 0x3B, 0x7A),
               right_color=RGBColor(0x0D, 0x3B, 0x7A),
               txt_color=RGBColor(0x88, 0xAA, 0xCC))


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 2 — WHY SWITCH (pain points)
# ─────────────────────────────────────────────────────────────────────────────

def build_slide2(slide):
    add_header(slide, "2 / 6")

    # Section title
    add_textbox(slide, Mm(14), Mm(13), Mm(120), Mm(9),
                "门诊换软件的四大驱动力",
                font_size=20, bold=True, color=BLUE)
    add_textbox(slide, Mm(136), Mm(14.5), Mm(100), Mm(7),
                "— 为什么现在是最好的切入时机",
                font_size=11, color=TEXT_MUTED)

    cards = [
        {
            "title": "监管合规压力加剧",
            "color": RED,
            "bg":    RGBColor(0xFF, 0xF5, 0xF5),
            "border": RED,
            "items": [
                "医疗专项检查持续列为年度重点",
                "医保稽查随时启动，历史数据均在审查范围",
                "税务核查力度加强，历史账目须经得起核验",
                "数据合规问题或引发补税、罚款乃至刑事风险",
            ],
        },
        {
            "title": "数据安全隐患频现",
            "color": AMBER,
            "bg":    RGBColor(0xFF, 0xF8, 0xEE),
            "border": AMBER,
            "items": [
                "现有软件数据不在门诊自己手中",
                "患者信息泄露事件行业内频繁发生",
                "SaaS厂商因亏损加剧存在倒闭、跑路风险",
                "一旦厂商停运，门诊历史数据面临永久丢失",
            ],
        },
        {
            "title": "供应链迟早被他人截走",
            "color": BLUE_MID,
            "bg":    RGBColor(0xF0, 0xF7, 0xFF),
            "border": BLUE_MID,
            "items": [
                "现有封闭软件无法对接外部供应链系统",
                "不换软件，订单数据永远留在竞对平台",
                "门诊采购渠道一旦被锁定，流失风险极高",
                "先布局者占据先机，晚行动则拱手相让",
            ],
        },
        {
            "title": "年费模式负担持续上升",
            "color": TEAL,
            "bg":    RGBColor(0xF0, 0xFA, 0xF8),
            "border": TEAL,
            "items": [
                "主流软件每年收取年费，成本逐年攀升",
                "续费涨价已成行业惯例，门诊叫苦不迭",
                "年费到期即停用，数据取回难度大",
                "转换窗口稍纵即逝，越早行动阻力越小",
            ],
        },
    ]

    card_w = Mm(62)
    gap    = Mm(4)
    total  = card_w * 4 + gap * 3
    start_x = (W - total) // 2
    card_y = Mm(24)
    card_h = Mm(118)

    for i, c in enumerate(cards):
        cx = start_x + (card_w + gap) * i
        # Background
        add_rect(slide, cx, card_y, card_w, card_h,
                 fill=c["bg"], line_color=c["border"], line_width_pt=0.5)
        # Top accent bar
        add_rect(slide, cx, card_y, card_w, Mm(2), fill=c["color"])
        # Title
        add_textbox(slide, cx + Mm(3), card_y + Mm(4), card_w - Mm(6), Mm(8),
                    c["title"], font_size=12, bold=True, color=TEXT_DARK)
        # Bullet items
        for j, item in enumerate(c["items"]):
            add_textbox(slide,
                        cx + Mm(5), card_y + Mm(14) + Mm(j * 24),
                        card_w - Mm(7), Mm(20),
                        f"· {item}", font_size=10, color=TEXT_MID, wrap=True)

    # Bottom note
    note = ("结论：门诊更换软件的外部驱动力正在快速聚集，合规压力与成本压力双重叠加，"
            "现在是与门诊建立合作的最佳窗口期。")
    add_textbox(slide, Mm(14), Mm(148), W - Mm(28), Mm(10),
                note, font_size=10, color=TEXT_MUTED, wrap=True)

    add_footer(slide)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 3 — ADVANTAGES
# ─────────────────────────────────────────────────────────────────────────────

def build_slide3(slide):
    add_header(slide, "3 / 6")

    add_textbox(slide, Mm(14), Mm(13), Mm(120), Mm(9),
                "轻松牙医的六大核心优势",
                font_size=20, bold=True, color=TEAL)
    add_textbox(slide, Mm(138), Mm(14.5), Mm(100), Mm(7),
                "— 让门诊换得安心、用得放心",
                font_size=11, color=TEXT_MUTED)

    adv_cards = [
        {
            "num": "1", "title": "数据完全自主，安全可控",
            "body": "支持本地化私有部署，数据存储于门诊自己的服务器，任何第三方无法擅自调取或泄露患者信息。",
            "tag": "私有化部署 · 本地存储", "highlight": True,
        },
        {
            "num": "2", "title": "历经三代迭代，功能与主流竞品看齐",
            "body": "轻松牙医已完成三个版本的重大升级，功能覆盖与市面主流软件对标，核心业务场景全面支持，并持续迭代。",
            "tag": "三代升级 · 持续演进", "highlight": False,
        },
        {
            "num": "3", "title": "合规数据管理，有效应对审查",
            "body": "支持数据按年份、项目、科目切割，自动对齐财务账目，分数库管理，助力门诊从容应对医保稽查与税务核查。",
            "tag": "分年管理 · 财务对齐", "highlight": False,
        },
        {
            "num": "4", "title": "灵活定制，满足大连锁及公立医院需求",
            "body": "提供深度定制化开发服务，大型连锁门诊或公立医院可依据自身业务场景，定制专属功能模块与数据字段。",
            "tag": "数据化定制 · 企业级扩展", "highlight": False,
        },
        {
            "num": "5", "title": "历史数据无缝迁移，零感知切换",
            "body": "支持将原有软件历史数据完整导入，门诊无需重新录入，配合三套皮肤主题，切换过渡体验顺畅自然。",
            "tag": "数据迁移 · 三套皮肤", "highlight": True,
        },
        {
            "num": "6", "title": "一次购买，终身使用，无年费负担",
            "body": "本地授权模式，一次性购买永久使用权，告别年费焦虑，显著降低门诊长期运营成本，财务可预期。",
            "tag": "买断制 · 零年费", "highlight": False,
        },
    ]

    cols = 3
    rows = 2
    margin_x = Mm(14)
    gap_x = Mm(4)
    gap_y = Mm(4)
    card_w = (W - margin_x * 2 - gap_x * (cols - 1)) // cols
    card_h = (BODY_H - Mm(12) - gap_y * (rows - 1)) // rows
    start_y = Mm(25)

    for i, c in enumerate(adv_cards):
        col = i % cols
        row = i // cols
        cx = margin_x + (card_w + gap_x) * col
        cy = start_y + (card_h + gap_y) * row
        bg = RGBColor(0xE3, 0xF0, 0xFF) if c["highlight"] else RGBColor(0xF7, 0xFA, 0xFF)
        border = TEAL if c["highlight"] else RGBColor(0xD0, 0xDA, 0xEA)
        add_rect(slide, cx, cy, card_w, card_h, fill=bg,
                 line_color=border, line_width_pt=0.75)
        # Number badge
        num_bg = TEAL if c["highlight"] else BLUE
        add_rect(slide, cx + Mm(3), cy + Mm(3), Mm(6), Mm(6), fill=num_bg)
        add_textbox(slide, cx + Mm(3), cy + Mm(3), Mm(6), Mm(6),
                    c["num"], font_size=9, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        # Title
        add_textbox(slide, cx + Mm(11), cy + Mm(3), card_w - Mm(14), Mm(8),
                    c["title"], font_size=11, bold=True, color=TEXT_DARK, wrap=True)
        # Body
        add_textbox(slide, cx + Mm(3), cy + Mm(12), card_w - Mm(6), card_h - Mm(22),
                    c["body"], font_size=10, color=TEXT_MID, wrap=True)
        # Tag
        add_rect(slide, cx + Mm(3), cy + card_h - Mm(8), card_w - Mm(6), Mm(6),
                 fill=BLUE_LIGHT)
        add_textbox(slide, cx + Mm(3), cy + card_h - Mm(8), card_w - Mm(6), Mm(6),
                    c["tag"], font_size=9, bold=True, color=BLUE,
                    align=PP_ALIGN.CENTER)

    add_footer(slide)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 4 — COOPERATION STEPS
# ─────────────────────────────────────────────────────────────────────────────

def build_slide4(slide):
    add_header(slide, "4 / 6")

    add_textbox(slide, Mm(14), Mm(13), W - Mm(28), Mm(9),
                "四步推广方案 — 我们如何携手落地",
                font_size=20, bold=True, color=BLUE)

    steps = [
        {
            "num": "01", "title": "精准拜访，建立心锚",
            "body": ("锁定目标门诊，由我方与供应商联合上门拜访。"
                     "重点介绍合规风险与数据安全议题，在门诊负责人心中埋下'风险预警'的认知锚点——"
                     "让他们在监管压力来临时第一时间想到我们的解决方案。"),
            "highlight": "我方可配合出席拜访，提供话术与演示支持",
            "accent": RGBColor(0xE8, 0xF4, 0xFD), "acolor": BLUE,
        },
        {
            "num": "02", "title": "组建本地团队，共担成本",
            "body": ("可在当地招聘或培养专职业务人员，负责持续跟进目标客户。"
                     "双方按50%比例共担人员成本，以不亏损为基本原则，确保双方可持续推进，降低各自的投入风险。"),
            "highlight": "成本共担 50%，风险均摊",
            "accent": RGBColor(0xFF, 0xF3, 0xE0), "acolor": AMBER,
        },
        {
            "num": "03", "title": "软件推广，灵活定价",
            "body": ("提供两种合作方式：\n"
                     "① 折扣进货：以五折价格进货，自行销售赚取差价；\n"
                     "② 免费铺设+提成：免费为门诊部署，按门诊年使用额1%提成，每家门诊每年封顶5,000元。"),
            "highlight": "五折进货 或 1% 提成（封顶5,000元/年）",
            "accent": RGBColor(0xE8, 0xF5, 0xE9), "acolor": GREEN_D,
        },
        {
            "num": "04", "title": "供应链激活，拉动首单",
            "body": ("门诊上线供应链后，以'首单1元'的活动策略激励客户完成第一笔下单。"
                     "一方面清理门诊旧库存，另一方面迅速让客户熟悉采购流程，建立使用习惯，加速复购转化。"),
            "highlight": "首单 1 元激活 · 清库存 · 建习惯",
            "accent": RGBColor(0xFC, 0xE4, 0xEC), "acolor": RED_D,
        },
    ]

    card_h = Mm(135)
    card_y = Mm(25)
    arrow_w = Mm(8)
    n = len(steps)
    total_arrow = arrow_w * (n - 1)
    margin_x = Mm(14)
    card_w = (W - margin_x * 2 - total_arrow) // n

    for i, s in enumerate(steps):
        cx = margin_x + (card_w + arrow_w) * i

        # Card background
        add_rect(slide, cx, card_y, card_w, card_h,
                 fill=RGBColor(0xF7, 0xFA, 0xFF),
                 line_color=RGBColor(0xD0, 0xDA, 0xEA), line_width_pt=0.75)

        # Step number label
        add_textbox(slide, cx + Mm(3), card_y + Mm(3), card_w - Mm(6), Mm(6),
                    f"STEP {s['num']}", font_size=9, bold=True, color=TEXT_MUTED)

        # Icon box
        add_rect(slide, cx + Mm(3), card_y + Mm(10), Mm(10), Mm(10),
                 fill=s["accent"])
        add_textbox(slide, cx + Mm(3), card_y + Mm(10), Mm(10), Mm(10),
                    "●", font_size=12, color=s["acolor"], align=PP_ALIGN.CENTER)

        # Title
        add_textbox(slide, cx + Mm(3), card_y + Mm(22), card_w - Mm(6), Mm(10),
                    s["title"], font_size=12, bold=True, color=TEXT_DARK, wrap=True)

        # Body
        add_textbox(slide, cx + Mm(3), card_y + Mm(33), card_w - Mm(6),
                    card_h - Mm(55),
                    s["body"], font_size=10, color=TEXT_MID, wrap=True)

        # Highlight pill at bottom
        add_rect(slide, cx + Mm(3), card_y + card_h - Mm(14),
                 card_w - Mm(6), Mm(11),
                 fill=BLUE_LIGHT)
        add_textbox(slide, cx + Mm(3), card_y + card_h - Mm(14),
                    card_w - Mm(6), Mm(11),
                    s["highlight"], font_size=9, bold=True,
                    color=s["acolor"], align=PP_ALIGN.CENTER, wrap=True)

        # Arrow (between cards)
        if i < n - 1:
            ax = cx + card_w + Mm(1.5)
            add_textbox(slide, ax, card_y + card_h // 2 - Mm(4), arrow_w - Mm(3), Mm(8),
                        "▶", font_size=14, color=RGBColor(0xC0, 0xCF, 0xE0),
                        align=PP_ALIGN.CENTER)

    add_footer(slide)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 5 — SUPPLY CHAIN VALUE (6 cards)
# ─────────────────────────────────────────────────────────────────────────────

def build_slide5(slide):
    add_header(slide, "5 / 6")

    add_textbox(slide, Mm(14), Mm(13), Mm(130), Mm(9),
                "松佰供应链的核心价值",
                font_size=20, bold=True, color=TEAL)
    add_textbox(slide, Mm(148), Mm(14.5), Mm(100), Mm(7),
                "— 软件 + 供应链深度融合，打造持续竞争壁垒",
                font_size=11, color=TEXT_MUTED)

    value_cards = [
        {
            "title": "实时库存感知，精准营销",
            "body": ("软件与供应链数据打通后，可实时掌握每家门诊的库存状态。"
                     "结合消费历史与使用周期，构建立体用户画像，精准推送补货建议，显著提升客单频次。"),
            "bg": RGBColor(0xF0, 0xF9, 0xF8),
            "border": RGBColor(0xB2, 0xDF, 0xDB),
            "tcolor": TEAL,
        },
        {
            "title": "AI 智能定价与商品推荐",
            "body": ("上传历史订单与价格数据后，AI 自动生成建议定价与关联商品推荐，"
                     "帮助供应商在不同场景下实现动态定价，提升利润空间与上架效率。"),
            "bg": RGBColor(0xF0, 0xF7, 0xFF),
            "border": RGBColor(0x90, 0xCA, 0xF9),
            "tcolor": BLUE,
        },
        {
            "title": "一键下单，一键入库",
            "body": ("软件与供应链系统全面打通，门诊可直接在软件内一键下单至供应链端；"
                     "到货后扫码或确认即完成一键入库，采购全程无纸化，彻底告别手工记账。"),
            "bg": RGBColor(0xE8, 0xF5, 0xE9),
            "border": RGBColor(0xA5, 0xD6, 0xA7),
            "tcolor": GREEN_D,
        },
        {
            "title": "持续黏性，提升复购率",
            "body": ("软件与供应链一体化构成天然闭环，门诊日常使用软件即可完成采购。"
                     "深度绑定场景使客户难以流失，月均活跃度与复购率均得到有效提升。"),
            "bg": RGBColor(0xFF, 0xFD, 0xE7),
            "border": RGBColor(0xFF, 0xE0, 0x82),
            "tcolor": RGBColor(0xB8, 0x86, 0x0B),
        },
        {
            "title": "数据主权，拒绝被平台绑架",
            "body": ("门诊数据完全归属门诊本身，供应商和轻松牙医均以合规权限共享洞察，而非单向掌控。"
                     "透明可信赖的数据治理模式是赢得大型门诊客户信任的关键。"),
            "bg": RGBColor(0xFC, 0xE4, 0xEC),
            "border": RGBColor(0xF4, 0x8F, 0xB1),
            "tcolor": RED_D,
        },
        {
            "title": "合规录入，百倍提效",
            "body": ("国家要求库房系统记录生产批准文号、批次号、有效期等完整信息，传统手录耗时极长。"
                     "接入松佰供应链后，商品分类、品名及三证资料自动同步至门诊端，录入效率提升百倍，"
                     "彻底解放库管人员。"),
            "bg": RGBColor(0xF3, 0xE5, 0xF5),
            "border": RGBColor(0xCE, 0x93, 0xD8),
            "tcolor": PURPLE,
        },
    ]

    cols = 3
    rows = 2
    margin_x = Mm(12)
    gap_x = Mm(4)
    gap_y = Mm(4)
    card_w = (W - margin_x * 2 - gap_x * (cols - 1)) // cols
    card_h = (BODY_H - Mm(13) - gap_y * (rows - 1)) // rows
    start_y = Mm(25)

    for i, c in enumerate(value_cards):
        col = i % cols
        row = i // cols
        cx = margin_x + (card_w + gap_x) * col
        cy = start_y + (card_h + gap_y) * row

        add_rect(slide, cx, cy, card_w, card_h, fill=c["bg"],
                 line_color=c["border"], line_width_pt=0.75)
        add_textbox(slide, cx + Mm(3), cy + Mm(3), card_w - Mm(6), Mm(8),
                    c["title"], font_size=11, bold=True, color=TEXT_DARK)
        add_textbox(slide, cx + Mm(3), cy + Mm(12), card_w - Mm(6),
                    card_h - Mm(15),
                    c["body"], font_size=10, color=TEXT_MID, wrap=True)

    add_footer(slide)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE 6 — PROFIT SHARING & CLOSING
# ─────────────────────────────────────────────────────────────────────────────

def build_slide6(slide):
    # Dark gradient background
    add_rect(slide, 0, 0, W, H, fill=RGBColor(0x0D, 0x3B, 0x7A))
    add_rect(slide, W // 2, 0, W // 2, H, fill=RGBColor(0x00, 0x69, 0x5C))

    add_header(slide, "6 / 6",
               bg=RGBColor(0x0D, 0x3B, 0x7A),
               brand_color=RGBColor(0xCC, 0xDD, 0xFF),
               num_color=RGBColor(0x88, 0xAA, 0xCC))

    add_textbox(slide, 0, Mm(13), W, Mm(6),
                "PROFIT SHARING MODEL",
                font_size=10, color=RGBColor(0xAA, 0xBB, 0xCC),
                align=PP_ALIGN.CENTER)

    add_textbox(slide, 0, Mm(20), W, Mm(10),
                "三方共赢 · 利益清晰可期",
                font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, 0, Mm(31), W, Mm(9),
                "分润模式一览",
                font_size=22, bold=True,
                color=RGBColor(0x7D, 0xD3, 0xF7),
                align=PP_ALIGN.CENTER)

    # Gold divider
    add_rect(slide, W // 2 - Mm(25), Mm(42), Mm(50), Mm(2), fill=GOLD)

    profit_cards = [
        {
            "title": "供应商公司",
            "title_color": RGBColor(0x7D, 0xD3, 0xF7),
            "items": [
                "打通门诊软件，直接触达终端采购决策",
                "供应链商城上线后，订单量持续增长",
                "AI 推荐提升客单价，精准营销降低获客成本",
                "门诊黏性提升，老客户流失率显著降低",
            ],
            "bullet_color": RGBColor(0x7D, 0xD3, 0xF7),
        },
        {
            "title": "推广人员",
            "title_color": RGBColor(0xB2, 0xF5, 0xEA),
            "items": [
                "软件销售差价（进货价五折，零售价自定）",
                "免费铺设模式：年均提成 1%，封顶 5,000 元 / 门诊",
                "供应链订单持续产生被动收益",
                "人员成本 50% 由双方分担，个人利润有保障",
            ],
            "bullet_color": RGBColor(0xB2, 0xF5, 0xEA),
        },
        {
            "title": "轻松牙医软件",
            "title_color": RGBColor(0xFF, 0xE0, 0x82),
            "items": [
                "软件装机量快速扩大，品牌影响力持续提升",
                "供应链平台交易量增长，获取平台服务收益",
                "借助供应商渠道快速覆盖区域市场",
                "定制化需求推动产品深化，形成差异化竞争优势",
            ],
            "bullet_color": RGBColor(0xFF, 0xE0, 0x82),
        },
    ]

    margin_x = Mm(14)
    gap = Mm(8)
    card_w = (W - margin_x * 2 - gap * 2) // 3
    card_y = Mm(47)
    card_h = Mm(110)

    for i, pc in enumerate(profit_cards):
        cx = margin_x + (card_w + gap) * i
        # Semi-transparent card (simulated with light overlay)
        add_rect(slide, cx, card_y, card_w, card_h,
                 fill=RGBColor(0x1A, 0x4A, 0x7E),
                 line_color=RGBColor(0x44, 0x77, 0xAA), line_width_pt=0.75)
        # Title
        add_textbox(slide, cx + Mm(4), card_y + Mm(4), card_w - Mm(8), Mm(9),
                    pc["title"], font_size=13, bold=True,
                    color=pc["title_color"])
        # Bullet items
        for j, item in enumerate(pc["items"]):
            add_textbox(slide,
                        cx + Mm(4), card_y + Mm(15) + Mm(j * 22),
                        card_w - Mm(8), Mm(19),
                        f"· {item}", font_size=11,
                        color=RGBColor(0xDD, 0xEE, 0xFF), wrap=True)

    add_footer(slide,
               text="携手共赢，共启口腔新局 — 期待与各位共同开创",
               left_color=RGBColor(0x0D, 0x3B, 0x7A),
               right_color=RGBColor(0x0D, 0x3B, 0x7A),
               txt_color=RGBColor(0x88, 0xAA, 0xCC))


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    # Use blank layout (index 6 in most default themes; fall back to first)
    layouts = prs.slide_layouts
    blank = None
    for lay in layouts:
        if lay.name in ("Blank", "blank", "空白"):
            blank = lay
            break
    if blank is None:
        blank = layouts[6] if len(layouts) > 6 else layouts[0]

    builders = [
        build_slide1,
        build_slide2,
        build_slide3,
        build_slide4,
        build_slide5,
        build_slide6,
    ]

    for builder in builders:
        slide = prs.slides.add_slide(blank)
        # Remove any placeholder shapes that come with the layout
        for ph in list(slide.placeholders):
            sp = ph._element
            sp.getparent().remove(sp)
        builder(slide)

    out = "presentation.pptx"
    prs.save(out)
    print(f"Saved → {out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
