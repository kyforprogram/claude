"""
Palo Alto Networks ライセンス体系まとめ PPTX 生成スクリプト
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ===== カラーパレット (Palo Alto Networks ブランドカラー) =====
COLOR_PAN_RED    = RGBColor(0xFA, 0x58, 0x2F)   # PAN オレンジレッド
COLOR_DARK_GRAY  = RGBColor(0x2C, 0x2C, 0x2C)
COLOR_MID_GRAY   = RGBColor(0x55, 0x55, 0x55)
COLOR_LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
COLOR_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_BLUE       = RGBColor(0x00, 0x6E, 0xC7)
COLOR_GREEN      = RGBColor(0x00, 0x96, 0x64)
COLOR_ORANGE     = RGBColor(0xFF, 0x8C, 0x00)
COLOR_PURPLE     = RGBColor(0x68, 0x21, 0x9E)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def new_prs():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_layout(prs):
    return prs.slide_layouts[6]   # completely blank


# ─────────────────────────────────────────────
# Helper: add filled rectangle
# ─────────────────────────────────────────────
def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def add_text_box(slide, text, left, top, width, height,
                 font_size=14, bold=False, color=COLOR_DARK_GRAY,
                 align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_para(tf, text, font_size=12, bold=False, color=COLOR_DARK_GRAY,
             align=PP_ALIGN.LEFT, indent_level=0):
    p = tf.add_paragraph()
    p.alignment = align
    p.level = indent_level
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p


# ─────────────────────────────────────────────
# Slide builders
# ─────────────────────────────────────────────

def slide_title(prs):
    slide = prs.slides.add_slide(blank_layout(prs))

    # 背景
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_DARK_GRAY)

    # 左アクセントバー
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    # 上部グラデーション風帯
    add_rect(slide, Inches(0.25), 0, SLIDE_W - Inches(0.25), Inches(0.06), COLOR_PAN_RED)

    # メインタイトル
    add_text_box(slide,
                 "Palo Alto Networks",
                 Inches(1.0), Inches(1.5), Inches(11), Inches(1.2),
                 font_size=44, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT)

    add_text_box(slide,
                 "ライセンス体系 完全まとめ",
                 Inches(1.0), Inches(2.6), Inches(11), Inches(1.1),
                 font_size=38, bold=True, color=COLOR_PAN_RED, align=PP_ALIGN.LEFT)

    add_text_box(slide,
                 "STRATA / Prisma SASE / Cortex",
                 Inches(1.0), Inches(3.65), Inches(10), Inches(0.7),
                 font_size=20, bold=False, color=RGBColor(0xCC, 0xCC, 0xCC), align=PP_ALIGN.LEFT)

    # 日付
    add_text_box(slide,
                 "2026年6月",
                 Inches(1.0), Inches(6.5), Inches(4), Inches(0.5),
                 font_size=13, color=RGBColor(0xAA, 0xAA, 0xAA), align=PP_ALIGN.LEFT)

    return slide


def slide_toc(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)
    add_rect(slide, Inches(0.25), 0, SLIDE_W - Inches(0.25), Inches(1.1), COLOR_DARK_GRAY)

    add_text_box(slide, "目次", Inches(0.5), Inches(0.2), Inches(10), Inches(0.75),
                 font_size=28, bold=True, color=COLOR_WHITE)

    items = [
        ("1", "Palo Alto Networks 製品ポートフォリオ概要"),
        ("2", "NGFW (PAシリーズ) ライセンス体系"),
        ("3", "サブスクリプションバンドル (Professional / Enterprise)"),
        ("4", "クラウド提供型セキュリティサービス (CDSS) 詳細"),
        ("5", "VM-Series / CN-Series — Software NGFW Credits"),
        ("6", "Prisma Access (SASE) ライセンス"),
        ("7", "Cortex XDR / XSIAM ライセンス"),
        ("8", "AIOps / Strata Cloud Manager"),
        ("9", "GlobalProtect / SD-WAN"),
        ("10", "サポートティア & ライセンス更新"),
        ("11", "製品別ライセンス比較まとめ"),
    ]

    cols = [items[:6], items[6:]]
    x_starts = [Inches(0.6), Inches(6.9)]
    for col_idx, col_items in enumerate(cols):
        x = x_starts[col_idx]
        for i, (num, label) in enumerate(col_items):
            y = Inches(1.3) + i * Inches(0.85)
            add_rect(slide, x, y, Inches(0.45), Inches(0.45), COLOR_PAN_RED)
            add_text_box(slide, num, x, y + Inches(0.03), Inches(0.45), Inches(0.42),
                         font_size=14, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
            add_text_box(slide, label, x + Inches(0.55), y + Inches(0.05),
                         Inches(5.8), Inches(0.45),
                         font_size=13, color=COLOR_DARK_GRAY)
    return slide


def slide_portfolio(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "1  製品ポートフォリオ概要",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    pillars = [
        ("STRATA\n(ネットワークセキュリティ)",
         COLOR_PAN_RED,
         ["• PA-Series (ハードウェアNGFW)",
          "• VM-Series (仮想NGFW)",
          "• CN-Series (コンテナNGFW)",
          "• Panorama (集中管理)",
          "• Strata Cloud Manager"]),
        ("Prisma SASE\n(クラウド & モバイル)",
         COLOR_BLUE,
         ["• Prisma Access (SWG/ZTNA)",
          "• Prisma SD-WAN",
          "• Autonomous DEM",
          "• ADEM / IoT Security",
          "• SaaS Security"]),
        ("Cortex\n(SOC & 自動化)",
         COLOR_GREEN,
         ["• Cortex XDR (EDR/XDR)",
          "• Cortex XSIAM (AI-SOC)",
          "• Cortex XSOAR (SOAR)",
          "• Cortex Xpanse (ASM)",
          "• Cortex Data Lake"]),
    ]

    for idx, (title, color, bullets) in enumerate(pillars):
        x = Inches(0.55) + idx * Inches(4.2)
        # ヘッダ
        add_rect(slide, x, Inches(1.3), Inches(3.9), Inches(0.9), color)
        add_text_box(slide, title, x + Inches(0.1), Inches(1.32),
                     Inches(3.7), Inches(0.86),
                     font_size=14, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        # ボックス
        add_rect(slide, x, Inches(2.2), Inches(3.9), Inches(4.9), COLOR_LIGHT_GRAY, color)
        tb = slide.shapes.add_textbox(x + Inches(0.15), Inches(2.35),
                                      Inches(3.6), Inches(4.6))
        tb.text_frame.word_wrap = True
        for j, b in enumerate(bullets):
            p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
            p.space_before = Pt(6)
            run = p.add_run()
            run.text = b
            run.font.size = Pt(13)
            run.font.color.rgb = COLOR_DARK_GRAY
    return slide


def slide_ngfw_overview(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "2  NGFW (PAシリーズ) ライセンス体系",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # 左列: ハードウェアライン
    add_text_box(slide, "ハードウェアラインナップ",
                 Inches(0.5), Inches(1.2), Inches(6), Inches(0.5),
                 font_size=16, bold=True, color=COLOR_PAN_RED)

    rows = [
        ("PA-400 シリーズ",  "小規模拠点・ブランチ向け\nPA-410 / 415 / 440 / 445 / 450 / 460"),
        ("PA-800 シリーズ",  "中規模企業向け\nPA-820 / 850"),
        ("PA-1400 シリーズ", "中〜大規模キャンパス向け\nPA-1410 / 1420"),
        ("PA-3400 シリーズ", "エンタープライズ向け\nPA-3410 / 3420 / 3430 / 3440"),
        ("PA-5400 シリーズ", "大規模データセンター向け\nPA-5410 / 5420 / 5430 / 5440 / 5445"),
        ("PA-7500 シリーズ", "超大規模DC・キャリア向け\nPA-7510 / 7520"),
    ]
    for i, (model, desc) in enumerate(rows):
        y = Inches(1.75) + i * Inches(0.82)
        add_rect(slide, Inches(0.5), y, Inches(2.2), Inches(0.72), COLOR_PAN_RED)
        add_text_box(slide, model, Inches(0.55), y + Inches(0.1),
                     Inches(2.1), Inches(0.55),
                     font_size=12, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, Inches(2.72), y, Inches(3.7), Inches(0.72), COLOR_LIGHT_GRAY, COLOR_PAN_RED)
        add_text_box(slide, desc, Inches(2.82), y + Inches(0.04),
                     Inches(3.5), Inches(0.68),
                     font_size=11, color=COLOR_DARK_GRAY)

    # 右列: ライセンス体系
    add_text_box(slide, "ライセンス構造",
                 Inches(7.1), Inches(1.2), Inches(5.7), Inches(0.5),
                 font_size=16, bold=True, color=COLOR_BLUE)

    license_items = [
        ("ハードウェア本体",     COLOR_DARK_GRAY,
         "アプライアンス購入（永続）\n別途サブスクリプション・サポートが必要"),
        ("サブスクリプション",   COLOR_PAN_RED,
         "年間 or 3年契約\nThreat Prevention / URL Filtering /\nWildFire / DNS Security 等"),
        ("サポートサービス",     COLOR_BLUE,
         "Premium Support（必須）\nPlatinum Support（オプション）\n更新: 1年 / 3年単位"),
        ("バンドルライセンス",   COLOR_GREEN,
         "Professional Bundle（4サービスセット）\nEnterprise Bundle（7サービスセット）\n単品より15〜25%割引"),
    ]

    for i, (title, color, body) in enumerate(license_items):
        y = Inches(1.75) + i * Inches(1.28)
        add_rect(slide, Inches(7.1), y, Inches(0.12), Inches(1.1), color)
        add_rect(slide, Inches(7.24), y, Inches(5.7), Inches(1.1), COLOR_LIGHT_GRAY)
        add_text_box(slide, title, Inches(7.35), y + Inches(0.05),
                     Inches(5.4), Inches(0.35),
                     font_size=13, bold=True, color=color)
        add_text_box(slide, body, Inches(7.35), y + Inches(0.38),
                     Inches(5.4), Inches(0.68),
                     font_size=11, color=COLOR_MID_GRAY)
    return slide


def slide_bundles(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "3  サブスクリプションバンドル",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # 凡例ヘッダ
    add_text_box(slide, "PAシリーズ向け2種類のバンドルライセンスが提供されています。単品購入より15〜25%割引。",
                 Inches(0.5), Inches(1.15), Inches(12.5), Inches(0.45),
                 font_size=13, color=COLOR_MID_GRAY, italic=True)

    bundles = [
        {
            "name": "Professional Bundle\n(PRO)",
            "color": COLOR_BLUE,
            "subtitle": "セキュリティ必須4機能セット",
            "price": "単品合計比 約15〜20%割引",
            "services": [
                "✔  Threat Prevention（IPS/アンチウイルス/アンチスパイウェア）",
                "✔  Advanced URL Filtering（URLフィルタリング）",
                "✔  Advanced WildFire（クラウドサンドボックス解析）",
                "✔  DNS Security（DNSベースのC2遮断）",
            ],
            "note": "基本的な脅威対策をカバー",
        },
        {
            "name": "Enterprise Bundle\n(ENT)",
            "color": COLOR_PAN_RED,
            "subtitle": "包括的エンタープライズセキュリティ",
            "price": "単品合計比 約20〜25%割引",
            "services": [
                "✔  PRO Bundle の全機能を含む（上記4サービス）",
                "✔  IoT Security（IoTデバイス可視化・制御）",
                "✔  SaaS Security Inline（SaaSアプリのリアルタイム制御）",
                "✔  SD-WAN（ソフトウェア定義WAN機能）",
            ],
            "note": "ゼロトラスト推進・OT/IoT環境に最適",
        },
    ]

    for idx, b in enumerate(bundles):
        x = Inches(0.5) + idx * Inches(6.3)
        add_rect(slide, x, Inches(1.65), Inches(6.0), Inches(0.72), b["color"])
        add_text_box(slide, b["name"], x + Inches(0.15), Inches(1.67),
                     Inches(3.5), Inches(0.68),
                     font_size=18, bold=True, color=COLOR_WHITE)
        add_text_box(slide, b["subtitle"], x + Inches(3.65), Inches(1.75),
                     Inches(2.2), Inches(0.55),
                     font_size=12, color=COLOR_WHITE, align=PP_ALIGN.RIGHT)

        add_rect(slide, x, Inches(2.39), Inches(6.0), Inches(4.7), COLOR_LIGHT_GRAY, b["color"])

        add_text_box(slide, b["price"], x + Inches(0.2), Inches(2.5),
                     Inches(5.6), Inches(0.4),
                     font_size=12, bold=True, color=b["color"])

        tb = slide.shapes.add_textbox(x + Inches(0.2), Inches(2.95),
                                      Inches(5.6), Inches(3.5))
        tb.text_frame.word_wrap = True
        for j, s in enumerate(b["services"]):
            p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
            p.space_before = Pt(7)
            run = p.add_run()
            run.text = s
            run.font.size = Pt(12.5)
            run.font.color.rgb = COLOR_DARK_GRAY

        add_text_box(slide, f"★ {b['note']}", x + Inches(0.2), Inches(6.7),
                     Inches(5.6), Inches(0.4),
                     font_size=12, bold=True, color=b["color"])

    # 契約期間メモ
    add_rect(slide, Inches(0.5), Inches(7.05), Inches(12.3), Inches(0.32), RGBColor(0xE8, 0xF4, 0xFF))
    add_text_box(slide, "※ 契約期間: 1年 / 3年（3年契約でさらに割引）　　※ SKU例: PAN-PA-440-BND-PRO-3YR / PAN-PA-440-BND-ENT-3YR",
                 Inches(0.6), Inches(7.07), Inches(12.1), Inches(0.28),
                 font_size=10, color=COLOR_MID_GRAY)
    return slide


def slide_cdss(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "4  クラウド提供型セキュリティサービス (CDSS)",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    services = [
        {
            "name": "Advanced Threat\nPrevention",
            "color": COLOR_PAN_RED,
            "body": ("IPS・アンチウイルス・アンチスパイウェア機能を提供。"
                     "インラインML推論によりゼロデイ脅威をリアルタイムブロック。"
                     "WildFireと連携してファイルレス攻撃も検知。"),
        },
        {
            "name": "Advanced URL\nFiltering",
            "color": COLOR_BLUE,
            "body": ("PANDBデータベースを使用したURLカテゴリフィルタリング。"
                     "ML機能により未知の悪性URLをリアルタイム検知・遮断。"
                     "フィッシング・クレデンシャル盗難対策を含む。"),
        },
        {
            "name": "Advanced\nWildFire",
            "color": COLOR_ORANGE,
            "body": ("クラウドベースのサンドボックスで未知マルウェアを解析。"
                     "静的解析 + 動的解析 + ML分析を組み合わせ。"
                     "1日40億以上のサンプルを処理・シグネチャ自動配信。"),
        },
        {
            "name": "DNS Security",
            "color": COLOR_GREEN,
            "body": ("DNSベースのC2通信・データ窃取をリアルタイム遮断。"
                     "インフラ変更不要でDNS通信を保護。"
                     "DNSトンネリング検知・DGAドメイン検知機能搭載。"),
        },
        {
            "name": "SaaS Security\nInline",
            "color": COLOR_PURPLE,
            "body": ("NGFWのインラインでSaaSアプリを可視化・制御。"
                     "Shadow IT検出、リアルタイムのDLP機能。"
                     "Microsoft 365 / Google Workspace等300以上のSaaSに対応。"),
        },
        {
            "name": "IoT Security",
            "color": RGBColor(0x00, 0x7A, 0x7A),
            "body": ("ネットワーク上のIoT/OTデバイスを自動識別・分類。"
                     "デバイスのリスクスコアリングと異常検知。"
                     "NAC連携によるセグメンテーション推奨機能。"),
        },
    ]

    cols = 3
    for idx, svc in enumerate(services):
        col = idx % cols
        row = idx // cols
        x = Inches(0.45) + col * Inches(4.25)
        y = Inches(1.2) + row * Inches(2.8)

        add_rect(slide, x, y, Inches(4.0), Inches(0.65), svc["color"])
        add_text_box(slide, svc["name"], x + Inches(0.1), y + Inches(0.05),
                     Inches(3.8), Inches(0.6),
                     font_size=14, bold=True, color=COLOR_WHITE)

        add_rect(slide, x, y + Inches(0.65), Inches(4.0), Inches(1.95),
                 COLOR_LIGHT_GRAY, svc["color"])
        add_text_box(slide, svc["body"], x + Inches(0.12), y + Inches(0.73),
                     Inches(3.76), Inches(1.8),
                     font_size=11, color=COLOR_DARK_GRAY, wrap=True)

    # 補足
    add_rect(slide, Inches(0.45), Inches(6.82), Inches(12.4), Inches(0.4),
             RGBColor(0xFF, 0xF0, 0xE8))
    add_text_box(slide,
                 "※ CDSS はすべてサブスクリプション契約。PA/VM/CN Series に追加購入可能。バンドルで購入すると割安。",
                 Inches(0.55), Inches(6.86), Inches(12.2), Inches(0.32),
                 font_size=10.5, color=COLOR_MID_GRAY)
    return slide


def slide_software_ngfw_credits(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "5  VM-Series / CN-Series — Software NGFW Credits",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # 概要
    add_rect(slide, Inches(0.45), Inches(1.15), Inches(12.4), Inches(0.75), RGBColor(0xE8, 0xF4, 0xFF))
    add_text_box(slide,
                 "Software NGFW Credits（Flexクレジット）は VM-Series / CN-Series 向けの柔軟なサブスクリプション体系です。\n"
                 "vCPU数・必要なCDSSをクレジットで選択でき、途中でスペック変更が可能です。",
                 Inches(0.6), Inches(1.2), Inches(12.1), Inches(0.65),
                 font_size=12, color=COLOR_DARK_GRAY)

    # 左列
    add_text_box(slide, "クレジット消費モデル",
                 Inches(0.5), Inches(2.05), Inches(6), Inches(0.45),
                 font_size=15, bold=True, color=COLOR_BLUE)

    credit_rows = [
        ("VM-Series vCPU", "vCPU数に応じてクレジット消費\n（メモリ量は問わない）"),
        ("CN-Series", "Kubernetes/コンテナ環境用\n同じクレジットプールを使用"),
        ("CDSS追加", "ATP / URL / WildFire / DNS Security等\n必要なサービスのみ選択可"),
        ("Panorama仮想", "仮想Panoramaアプライアンスにも\nクレジット適用可能"),
    ]
    for i, (label, desc) in enumerate(credit_rows):
        y = Inches(2.55) + i * Inches(1.0)
        add_rect(slide, Inches(0.5), y, Inches(2.0), Inches(0.85), COLOR_BLUE)
        add_text_box(slide, label, Inches(0.55), y + Inches(0.1),
                     Inches(1.9), Inches(0.68),
                     font_size=11, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, Inches(2.52), y, Inches(3.9), Inches(0.85), COLOR_LIGHT_GRAY, COLOR_BLUE)
        add_text_box(slide, desc, Inches(2.62), y + Inches(0.05),
                     Inches(3.7), Inches(0.78),
                     font_size=11, color=COLOR_DARK_GRAY)

    # 右列
    add_text_box(slide, "対応プラットフォーム",
                 Inches(7.1), Inches(2.05), Inches(5.7), Inches(0.45),
                 font_size=15, bold=True, color=COLOR_PAN_RED)

    platforms = [
        ("オンプレミス仮想化", "VMware ESXi, Hyper-V, KVM"),
        ("パブリッククラウド", "AWS, Microsoft Azure, Google Cloud\nOracle Cloud, Alibaba Cloud"),
        ("コンテナ環境", "Kubernetes (CN-Series)\nOpenShift 対応"),
    ]
    for i, (pname, pdesc) in enumerate(platforms):
        y = Inches(2.55) + i * Inches(1.1)
        add_rect(slide, Inches(7.1), y, Inches(5.8), Inches(0.95), COLOR_LIGHT_GRAY, COLOR_PAN_RED)
        add_text_box(slide, pname, Inches(7.2), y + Inches(0.06),
                     Inches(5.5), Inches(0.38),
                     font_size=13, bold=True, color=COLOR_PAN_RED)
        add_text_box(slide, pdesc, Inches(7.2), y + Inches(0.44),
                     Inches(5.5), Inches(0.48),
                     font_size=11.5, color=COLOR_MID_GRAY)

    add_text_box(slide, "契約期間 & 管理",
                 Inches(7.1), Inches(5.65), Inches(5.7), Inches(0.45),
                 font_size=15, bold=True, color=COLOR_GREEN)
    mgmt_items = [
        "• クレジットは購入後12ヶ月 or 36ヶ月で期限切れ（未使用分も失効）",
        "• Customer Support Portal から集中管理",
        "• 途中でvCPU数・CDSSの変更が可能（柔軟性が高い）",
        "• 更新時は追加クレジット購入で継続",
    ]
    tb = slide.shapes.add_textbox(Inches(7.1), Inches(6.15), Inches(5.8), Inches(1.1))
    tb.text_frame.word_wrap = True
    for j, item in enumerate(mgmt_items):
        p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
        p.space_before = Pt(5)
        run = p.add_run()
        run.text = item
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_DARK_GRAY
    return slide


def slide_prisma_access(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "6  Prisma Access (SASE) ライセンス",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    add_text_box(slide,
                 "Prisma Access は Palo Alto Networks のクラウド型 SASE ソリューション。ユーザー数・帯域・ロケーション数に基づくサブスクリプション。",
                 Inches(0.5), Inches(1.15), Inches(12.4), Inches(0.42),
                 font_size=12.5, color=COLOR_MID_GRAY, italic=True)

    # ライセンスカード
    cards = [
        {
            "title": "Prisma Access\n(ベース)",
            "color": COLOR_BLUE,
            "items": [
                "Secure Web Gateway (SWG)",
                "Cloud Access Security Broker (CASB)",
                "Zero Trust Network Access (ZTNA)",
                "Firewall-as-a-Service (FWaaS)",
                "SD-WAN 統合オプション",
                "対象: モバイルユーザー & 拠点",
            ],
        },
        {
            "title": "Prisma Access\n+ Advanced Threat",
            "color": COLOR_PAN_RED,
            "items": [
                "ベースの全機能を含む",
                "Advanced WildFire 統合",
                "DNS Security 統合",
                "Advanced URL Filtering 統合",
                "Advanced Threat Prevention",
                "データ損失防止 (DLP) オプション",
            ],
        },
        {
            "title": "アドオン\nオプション",
            "color": COLOR_GREEN,
            "items": [
                "Autonomous DEM（デジタルエクスペリエンス監視）",
                "Enterprise DLP（データ損失防止）",
                "SaaS Security API（SaaSデータ保護）",
                "IoT Security（IoTデバイス保護）",
                "Browser Isolation（ブラウザ分離）",
                "Prisma SD-WAN 拠点追加",
            ],
        },
    ]
    for idx, c in enumerate(cards):
        x = Inches(0.45) + idx * Inches(4.25)
        add_rect(slide, x, Inches(1.65), Inches(4.0), Inches(0.68), c["color"])
        add_text_box(slide, c["title"], x + Inches(0.1), Inches(1.67),
                     Inches(3.8), Inches(0.64),
                     font_size=14, bold=True, color=COLOR_WHITE)
        add_rect(slide, x, Inches(2.35), Inches(4.0), Inches(3.85), COLOR_LIGHT_GRAY, c["color"])
        tb = slide.shapes.add_textbox(x + Inches(0.15), Inches(2.45),
                                      Inches(3.7), Inches(3.7))
        tb.text_frame.word_wrap = True
        for j, item in enumerate(c["items"]):
            p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
            p.space_before = Pt(6)
            run = p.add_run()
            run.text = f"• {item}"
            run.font.size = Pt(12)
            run.font.color.rgb = COLOR_DARK_GRAY

    # 価格帯
    add_rect(slide, Inches(0.45), Inches(6.28), Inches(12.4), Inches(1.0), RGBColor(0xF0, 0xF0, 0xF8))
    add_text_box(slide, "価格帯・契約条件",
                 Inches(0.6), Inches(6.32), Inches(3), Inches(0.38),
                 font_size=13, bold=True, color=COLOR_BLUE)
    add_text_box(slide,
                 "• ユーザー課金: 約 $10〜$25/ユーザー/月（プラン・ボリュームにより変動）\n"
                 "• 拠点課金: 帯域・ロケーション数に応じて追加課金\n"
                 "• 推奨契約: 3年（2年・1年も可）。複数年コミットで15〜30%割引\n"
                 "• 大規模: 5,000ユーザー以上の場合、年間契約額は7桁（百万ドル超）になることも",
                 Inches(0.6), Inches(6.7), Inches(12.1), Inches(0.55),
                 font_size=11, color=COLOR_DARK_GRAY)
    return slide


def slide_cortex(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "7  Cortex XDR / XSIAM ライセンス",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # ─── Cortex XDR ───
    add_rect(slide, Inches(0.45), Inches(1.15), Inches(8.5), Inches(0.48), COLOR_GREEN)
    add_text_box(slide, "Cortex XDR  — エンドポイント保護 & 拡張検出・対応",
                 Inches(0.55), Inches(1.18), Inches(8.3), Inches(0.42),
                 font_size=14, bold=True, color=COLOR_WHITE)

    xdr_plans = [
        {
            "name": "XDR Prevent",
            "color": RGBColor(0x00, 0x80, 0x40),
            "items": [
                "次世代アンチウイルス (NGAV)",
                "マルウェア・ランサムウェア遮断",
                "エクスプロイト防止",
                "ファイルレス攻撃対策",
                "デバイス制御 / ディスク暗号化",
                "課金: エンドポイント単位/年",
            ],
        },
        {
            "name": "XDR Pro\nper Endpoint",
            "color": RGBColor(0x00, 0x64, 0x30),
            "items": [
                "Prevent の全機能を含む",
                "EDR（エンドポイント検出・対応）",
                "AIを活用した脅威ハンティング",
                "アラート優先度付け・調査自動化",
                "Cortex Data Lake 連携",
                "課金: エンドポイント単位/年",
            ],
        },
        {
            "name": "XDR Pro\nper GB",
            "color": RGBColor(0x00, 0x50, 0x28),
            "items": [
                "Pro per Endpoint の全機能を含む",
                "サードパーティログ取り込み対応",
                "SIEM的なログ分析機能",
                "ネットワーク・クラウドログ対応",
                "課金: 取り込みデータ量 (GB)/日",
                "大規模SOC環境に最適",
            ],
        },
    ]
    for idx, plan in enumerate(xdr_plans):
        x = Inches(0.45) + idx * Inches(2.85)
        y = Inches(1.68)
        add_rect(slide, x, y, Inches(2.72), Inches(0.52), plan["color"])
        add_text_box(slide, plan["name"], x + Inches(0.08), y + Inches(0.04),
                     Inches(2.58), Inches(0.46),
                     font_size=13, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, x, y + Inches(0.52), Inches(2.72), Inches(2.9),
                 COLOR_LIGHT_GRAY, plan["color"])
        tb = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.62),
                                      Inches(2.52), Inches(2.75))
        tb.text_frame.word_wrap = True
        for j, item in enumerate(plan["items"]):
            p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
            p.space_before = Pt(5)
            run = p.add_run()
            run.text = f"• {item}"
            run.font.size = Pt(10.5)
            run.font.color.rgb = COLOR_DARK_GRAY

    # ─── Cortex XSIAM ───
    add_rect(slide, Inches(9.1), Inches(1.15), Inches(3.8), Inches(0.48), COLOR_PURPLE)
    add_text_box(slide, "Cortex XSIAM",
                 Inches(9.2), Inches(1.18), Inches(3.6), Inches(0.42),
                 font_size=14, bold=True, color=COLOR_WHITE)

    xsiam_items = [
        "AI駆動型の次世代SOCプラットフォーム",
        "XDR + SIEM + SOAR + TIP + UEBA を統合",
        "エンドポイント・ネットワーク・クラウド・SaaSのログを一元分析",
        "機械学習による脅威自動トリアージ",
        "XSOAR (SOAR) 機能を内包",
        "課金: データ取り込み量 (GB) ベース",
        "対象: 大規模エンタープライズSOC",
    ]
    add_rect(slide, Inches(9.1), Inches(1.65), Inches(3.8), Inches(4.75),
             COLOR_LIGHT_GRAY, COLOR_PURPLE)
    tb = slide.shapes.add_textbox(Inches(9.2), Inches(1.78), Inches(3.6), Inches(4.55))
    tb.text_frame.word_wrap = True
    for j, item in enumerate(xsiam_items):
        p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
        p.space_before = Pt(7)
        run = p.add_run()
        run.text = f"• {item}"
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_DARK_GRAY

    # ─── 補足 ───
    add_rect(slide, Inches(0.45), Inches(4.63), Inches(8.5), Inches(2.6),
             RGBColor(0xF8, 0xFF, 0xF0), COLOR_GREEN)
    add_text_box(slide, "Cortex Data Lake",
                 Inches(0.6), Inches(4.7), Inches(4), Inches(0.38),
                 font_size=13, bold=True, color=COLOR_GREEN)
    add_text_box(slide,
                 "• XDR / XSIAM のバックエンドストレージ（クラウド）\n"
                 "• ストレージ容量に応じた追加課金\n"
                 "• NGFWログ・エンドポイントログ・サードパーティログを集約\n"
                 "• デフォルト: 30日〜90日保持（プランにより異なる）",
                 Inches(0.6), Inches(5.12), Inches(8.2), Inches(2.0),
                 font_size=11.5, color=COLOR_DARK_GRAY)
    return slide


def slide_aiops_strata(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "8  AIOps / Strata Cloud Manager",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # 移行案内
    add_rect(slide, Inches(0.45), Inches(1.12), Inches(12.4), Inches(0.65),
             RGBColor(0xFF, 0xF0, 0xE0))
    add_text_box(slide,
                 "⚠  2025年5月8日 AIOps for NGFW Premium ライセンス 販売終了（EOS）。"
                 "既存顧客は「Strata Cloud Manager」ライセンスへ段階的に無償移行（2025年3月〜）",
                 Inches(0.6), Inches(1.16), Inches(12.1), Inches(0.58),
                 font_size=12, bold=True, color=RGBColor(0xCC, 0x44, 0x00))

    # SCM / AIOps比較
    tiers = [
        {
            "name": "AIOps Free\n→ SCM Essentials",
            "color": COLOR_GREEN,
            "items": [
                "デバイステレメトリ収集・分析",
                "ベストプラクティス評価 (BPA)",
                "セキュリティポスチャスコア",
                "設定エラー検出・推奨アクション",
                "無償（NGFWサポート契約に付帯）",
            ],
        },
        {
            "name": "AIOps Premium\n→ SCM Advanced",
            "color": COLOR_BLUE,
            "items": [
                "Free/Essentials の全機能を含む",
                "クラウド管理 for NGFW（Cloud Management）",
                "自動設定プッシュ・コンプライアンス管理",
                "AI駆動による予測的推奨",
                "サブスクリプション別途購入（EOS: 2025/5/8）",
            ],
        },
        {
            "name": "Strata Cloud\nManager (SCM)",
            "color": COLOR_PAN_RED,
            "items": [
                "AIOpsの後継統合管理プラットフォーム",
                "PA/VM/CN Series + Prisma Access 統合管理",
                "ゼロトラスト戦略の可視化・運用",
                "AI推奨・自動化ワークフロー",
                "Essentials（無償）/ Advanced（有償）の2階層",
            ],
        },
    ]
    for idx, t in enumerate(tiers):
        x = Inches(0.45) + idx * Inches(4.25)
        add_rect(slide, x, Inches(1.88), Inches(4.0), Inches(0.6), t["color"])
        add_text_box(slide, t["name"], x + Inches(0.1), Inches(1.90),
                     Inches(3.8), Inches(0.58),
                     font_size=13, bold=True, color=COLOR_WHITE)
        add_rect(slide, x, Inches(2.5), Inches(4.0), Inches(3.35), COLOR_LIGHT_GRAY, t["color"])
        tb = slide.shapes.add_textbox(x + Inches(0.15), Inches(2.6),
                                      Inches(3.7), Inches(3.2))
        tb.text_frame.word_wrap = True
        for j, item in enumerate(t["items"]):
            p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
            p.space_before = Pt(8)
            run = p.add_run()
            run.text = f"• {item}"
            run.font.size = Pt(12)
            run.font.color.rgb = COLOR_DARK_GRAY

    # 移行タイムライン
    add_rect(slide, Inches(0.45), Inches(5.95), Inches(12.4), Inches(1.35),
             RGBColor(0xF0, 0xF8, 0xFF))
    add_text_box(slide, "移行タイムライン",
                 Inches(0.6), Inches(6.0), Inches(3), Inches(0.38),
                 font_size=13, bold=True, color=COLOR_BLUE)

    timeline = [
        ("2025年3月",  "AIOps Free → SCM Essentials 自動移行開始（無償）"),
        ("2025年5月", "AIOps for NGFW Premium EOS（販売終了）"),
        ("2025年〜",   "既存 Premium ユーザーは契約終了まで利用継続可"),
        ("以降",       "新規購入は Strata Cloud Manager Advanced で対応"),
    ]
    x_base = Inches(0.6)
    for i, (date, desc) in enumerate(timeline):
        x = x_base + i * Inches(3.05)
        add_rect(slide, x, Inches(6.42), Inches(1.1), Inches(0.32), COLOR_BLUE)
        add_text_box(slide, date, x, Inches(6.43), Inches(1.1), Inches(0.3),
                     font_size=10, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_text_box(slide, desc, x, Inches(6.77), Inches(2.95), Inches(0.45),
                     font_size=10, color=COLOR_DARK_GRAY)
    return slide


def slide_gp_sdwan(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "9  GlobalProtect / SD-WAN",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # GlobalProtect
    add_rect(slide, Inches(0.45), Inches(1.15), Inches(6.15), Inches(0.5), COLOR_ORANGE)
    add_text_box(slide, "GlobalProtect — リモートアクセス VPN / ZTNA",
                 Inches(0.55), Inches(1.18), Inches(6.0), Inches(0.44),
                 font_size=14, bold=True, color=COLOR_WHITE)

    gp_items = [
        ("ベース機能（無償）",
         "• NGFWにバンドル\n"
         "• IPSec/SSL VPN機能\n"
         "• シンプルなリモートアクセス"),
        ("ゲートウェイライセンス\n（有償・必須）",
         "• HIチェック（Host Information Profile）\n"
         "• モバイルアプリ対応（iOS/Android）\n"
         "• IPv6サポート\n"
         "• GlobalProtect Gateway ごとに購入"),
        ("ZTNA 2.0（Prisma Access）",
         "• クラウド型ゼロトラストNA\n"
         "• Prisma Access ライセンスに内包\n"
         "• SASE統合・App-ID/User-ID連携"),
    ]
    for i, (title, body) in enumerate(gp_items):
        x = Inches(0.45) + i * Inches(2.05)
        y = Inches(1.72)
        add_rect(slide, x, y, Inches(1.95), Inches(0.58), COLOR_ORANGE)
        add_text_box(slide, title, x + Inches(0.06), y + Inches(0.04),
                     Inches(1.84), Inches(0.52),
                     font_size=10.5, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, x, y + Inches(0.58), Inches(1.95), Inches(2.65),
                 COLOR_LIGHT_GRAY, COLOR_ORANGE)
        add_text_box(slide, body, x + Inches(0.08), y + Inches(0.66),
                     Inches(1.8), Inches(2.5),
                     font_size=10.5, color=COLOR_DARK_GRAY)

    # SD-WAN
    add_rect(slide, Inches(6.8), Inches(1.15), Inches(6.1), Inches(0.5), COLOR_PURPLE)
    add_text_box(slide, "SD-WAN — ソフトウェア定義WAN",
                 Inches(6.9), Inches(1.18), Inches(5.95), Inches(0.44),
                 font_size=14, bold=True, color=COLOR_WHITE)

    sdwan_rows = [
        ("Prisma SD-WAN",
         "• Palo Alto Networks クラウド型SD-WAN\n"
         "• Prisma Accessと統合（SASE構成）\n"
         "• 拠点数・帯域によるサブスクリプション\n"
         "• AIOps for SD-WAN を内包"),
        ("NGFW SD-WAN\nアドオン",
         "• PAシリーズへのSD-WAN機能追加\n"
         "• Enterprise Bundleに含まれる\n"
         "• PAN-OSベースでNGFW一体化\n"
         "• アプリベースのパス制御"),
    ]
    for i, (title, body) in enumerate(sdwan_rows):
        x = Inches(6.8)
        y = Inches(1.72) + i * Inches(1.75)
        add_rect(slide, x, y, Inches(2.1), Inches(1.6), COLOR_PURPLE)
        add_text_box(slide, title, x + Inches(0.08), y + Inches(0.1),
                     Inches(1.96), Inches(1.45),
                     font_size=12, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, x + Inches(2.12), y, Inches(3.85), Inches(1.6),
                 COLOR_LIGHT_GRAY, COLOR_PURPLE)
        add_text_box(slide, body, x + Inches(2.22), y + Inches(0.08),
                     Inches(3.6), Inches(1.48),
                     font_size=11.5, color=COLOR_DARK_GRAY)

    # 注記
    add_rect(slide, Inches(0.45), Inches(5.3), Inches(12.4), Inches(2.0),
             RGBColor(0xF8, 0xF0, 0xFF))
    add_text_box(slide, "ライセンス要件まとめ (GlobalProtect)",
                 Inches(0.6), Inches(5.36), Inches(6), Inches(0.38),
                 font_size=13, bold=True, color=COLOR_PURPLE)
    add_text_box(slide,
                 "• ゲートウェイ側に有償ライセンスが必要（ポータルは無償で構築可）\n"
                 "• HIPチェック・モバイルアプリを使う場合は必ずライセンス購入\n"
                 "• ライセンスは各Gateway（PA機）ごとに購入（ゲートウェイ数分必要）\n"
                 "• 更新: 年間 or 3年サブスクリプション",
                 Inches(0.6), Inches(5.78), Inches(5.9), Inches(1.45),
                 font_size=11.5, color=COLOR_DARK_GRAY)
    add_text_box(slide, "ライセンス要件まとめ (SD-WAN)",
                 Inches(6.9), Inches(5.36), Inches(5.5), Inches(0.38),
                 font_size=13, bold=True, color=COLOR_PURPLE)
    add_text_box(slide,
                 "• NGFWアドオン版: Enterprise Bundleに含まれるため追加費用なし\n"
                 "• Prisma SD-WAN: 拠点数・帯域に応じた別途ライセンス\n"
                 "• SASE統合ならPrisma SD-WAN + Prisma Accessが推奨",
                 Inches(6.9), Inches(5.78), Inches(5.9), Inches(1.45),
                 font_size=11.5, color=COLOR_DARK_GRAY)
    return slide


def slide_support(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "10  サポートティア & ライセンス更新",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # ─── サポートティア ───
    add_text_box(slide, "サポートティア比較",
                 Inches(0.5), Inches(1.15), Inches(5), Inches(0.4),
                 font_size=15, bold=True, color=COLOR_DARK_GRAY)

    support_tiers = [
        {
            "name": "Premium Support",
            "color": COLOR_BLUE,
            "features": [
                "24x7x365 サポート",
                "電話 & Webケース作成",
                "重大度1: 1時間以内初期応答",
                "PAN-OS アップグレード / RMA",
                "ハードウェア購入時に必須",
                "サブスクリプションに付帯",
            ],
        },
        {
            "name": "Platinum Support",
            "color": COLOR_PAN_RED,
            "features": [
                "Premium の全機能を含む",
                "担当 TAM（テクニカルアカウントマネージャ）",
                "シニアエンジニア直接対応",
                "重大度1: 30分以内応答",
                "年間レビュー・設計支援",
                "大規模エンタープライズ向け",
            ],
        },
    ]
    for idx, t in enumerate(support_tiers):
        x = Inches(0.45) + idx * Inches(4.5)
        add_rect(slide, x, Inches(1.6), Inches(4.2), Inches(0.55), t["color"])
        add_text_box(slide, t["name"], x + Inches(0.1), Inches(1.62),
                     Inches(4.0), Inches(0.5),
                     font_size=15, bold=True, color=COLOR_WHITE)
        add_rect(slide, x, Inches(2.17), Inches(4.2), Inches(2.95),
                 COLOR_LIGHT_GRAY, t["color"])
        tb = slide.shapes.add_textbox(x + Inches(0.15), Inches(2.28),
                                      Inches(3.9), Inches(2.75))
        tb.text_frame.word_wrap = True
        for j, f in enumerate(t["features"]):
            p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
            p.space_before = Pt(7)
            run = p.add_run()
            run.text = f"• {f}"
            run.font.size = Pt(12.5)
            run.font.color.rgb = COLOR_DARK_GRAY

    # ─── 更新サイクル ───
    add_text_box(slide, "ライセンス更新サイクル",
                 Inches(9.2), Inches(1.15), Inches(3.8), Inches(0.4),
                 font_size=15, bold=True, color=COLOR_DARK_GRAY)

    renewal_items = [
        ("1年契約",  COLOR_PAN_RED,   "最も柔軟。年次見直し可\n割引率は低め"),
        ("3年契約",  COLOR_BLUE,      "推奨。15〜25%割引\nコスト最適"),
        ("ELA",     COLOR_GREEN,      "エンタープライズ包括契約\n複数製品を一括契約\n大幅割引（30〜60%）"),
    ]
    for i, (term, color, desc) in enumerate(renewal_items):
        y = Inches(1.6) + i * Inches(1.25)
        add_rect(slide, Inches(9.2), y, Inches(1.2), Inches(1.1), color)
        add_text_box(slide, term, Inches(9.2), y + Inches(0.28),
                     Inches(1.2), Inches(0.55),
                     font_size=14, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, Inches(10.42), y, Inches(2.65), Inches(1.1), COLOR_LIGHT_GRAY, color)
        add_text_box(slide, desc, Inches(10.52), y + Inches(0.1),
                     Inches(2.5), Inches(0.92),
                     font_size=11.5, color=COLOR_DARK_GRAY)

    # ─── 更新管理 ───
    add_rect(slide, Inches(0.45), Inches(5.22), Inches(12.4), Inches(2.05),
             RGBColor(0xF0, 0xF8, 0xF0))
    add_text_box(slide, "ライセンス管理・更新ポイント",
                 Inches(0.6), Inches(5.28), Inches(5), Inches(0.38),
                 font_size=14, bold=True, color=COLOR_GREEN)
    mgmt_content = [
        "Customer Support Portal (support.paloaltonetworks.com) でライセンス一元管理",
        "有効期限60〜90日前に更新通知。期限切れ後は機能が停止（ライセンス認証必要）",
        "デバイス証明書: 2026年2月以降、CDSS接続に必須（PA/VM/CN/Panorama）",
        "Software NGFW Credits: Customer Support Portal からクレジット残高・消費状況を確認可",
        "Panorama: 台数分のデバイスライセンス + 管理ノードのサポート契約が必要",
    ]
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(5.7), Inches(12.1), Inches(1.45))
    tb.text_frame.word_wrap = True
    for j, item in enumerate(mgmt_content):
        p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
        p.space_before = Pt(5)
        run = p.add_run()
        run.text = f"• {item}"
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_DARK_GRAY
    return slide


def slide_summary_table(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_WHITE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)

    add_text_box(slide, "11  製品別ライセンス比較まとめ",
                 Inches(0.5), Inches(0.18), Inches(12), Inches(0.75),
                 font_size=26, bold=True, color=COLOR_WHITE)

    # テーブルヘッダ
    headers = ["製品・サービス", "ライセンス種別", "課金モデル", "契約期間", "主な用途"]
    col_widths = [Inches(2.5), Inches(2.8), Inches(2.5), Inches(1.6), Inches(3.5)]
    x_pos = [Inches(0.25)]
    for w in col_widths[:-1]:
        x_pos.append(x_pos[-1] + w)

    for i, (h, w, x) in enumerate(zip(headers, col_widths, x_pos)):
        add_rect(slide, x, Inches(1.15), w - Inches(0.04), Inches(0.45), COLOR_DARK_GRAY)
        add_text_box(slide, h, x + Inches(0.05), Inches(1.18),
                     w - Inches(0.08), Inches(0.4),
                     font_size=11, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

    rows = [
        ("PA-Series NGFW\n（ハードウェア）",
         "本体 + サブスクリプション\n+ サポート",
         "アプライアンス\n（年/3年サブスク）",
         "1年 or 3年",
         "物理環境・オンプレミスFW"),
        ("VM/CN-Series\n（ソフトウェア）",
         "Software NGFW Credits\n（Flexクレジット）",
         "クレジット消費\n（vCPU基準）",
         "12ヶ月 or 36ヶ月",
         "仮想化・クラウド・コンテナ環境"),
        ("CDSS\n（各種セキュリティ）",
         "ATP / WildFire / URL /\nDNS / SaaS / IoT 個別",
         "デバイス単位\n（年/3年サブスク）",
         "1年 or 3年",
         "NGFW機能拡張・脅威対策"),
        ("Prisma Access\n（SASE）",
         "ベース / Advanced\nThreat / アドオン",
         "ユーザー数 +\n拠点数・帯域",
         "1年/2年/3年",
         "クラウド・モバイル・ゼロトラスト"),
        ("Cortex XDR\n（EDR/XDR）",
         "Prevent / Pro per EP\n/ Pro per GB",
         "エンドポイント数\nor GB/日",
         "1年 or 3年",
         "エンドポイント保護・SOC分析"),
        ("Cortex XSIAM\n（AI-SOC）",
         "統合SOCプラットフォーム",
         "データ取り込み量\n（GB）",
         "1年 or 3年",
         "大規模SOC・SIEM代替"),
        ("GlobalProtect\n（VPN/ZTNA）",
         "Gatewayライセンス\n（基本は無償込み）",
         "ゲートウェイ単位\n（年/3年サブスク）",
         "1年 or 3年",
         "リモートアクセス・ゼロトラスト"),
        ("Strata Cloud\nManager / AIOps",
         "Essentials（無償）\n/ Advanced（有償）",
         "管理対象デバイス数\nor 包括",
         "1年 or 3年",
         "NGFW集中管理・AIポスチャ管理"),
    ]

    row_colors = [COLOR_WHITE, COLOR_LIGHT_GRAY]
    for ri, row in enumerate(rows):
        y = Inches(1.63) + ri * Inches(0.64)
        bg = row_colors[ri % 2]
        for ci, (cell, w, x) in enumerate(zip(row, col_widths, x_pos)):
            add_rect(slide, x, y, w - Inches(0.04), Inches(0.62), bg)
            txt_color = COLOR_PAN_RED if ci == 0 else COLOR_DARK_GRAY
            add_text_box(slide, cell, x + Inches(0.06), y + Inches(0.03),
                         w - Inches(0.1), Inches(0.58),
                         font_size=10, bold=(ci == 0), color=txt_color)

    add_rect(slide, Inches(0.25), Inches(6.77), Inches(12.85), Inches(0.5),
             RGBColor(0xF0, 0xF0, 0xF0))
    add_text_box(slide,
                 "※ 価格は構成・ボリューム・契約期間・パートナー割引により大きく変動します。正確な見積りは Palo Alto Networks または認定パートナーにお問い合わせください。",
                 Inches(0.35), Inches(6.82), Inches(12.6), Inches(0.4),
                 font_size=10, color=COLOR_MID_GRAY, italic=True)
    return slide


def slide_closing(prs):
    slide = prs.slides.add_slide(blank_layout(prs))
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, COLOR_DARK_GRAY)
    add_rect(slide, 0, 0, Inches(0.25), SLIDE_H, COLOR_PAN_RED)
    add_rect(slide, Inches(0.25), 0, SLIDE_W - Inches(0.25), Inches(0.06), COLOR_PAN_RED)

    add_text_box(slide, "Palo Alto Networks ライセンス体系 まとめ",
                 Inches(1.0), Inches(2.0), Inches(11), Inches(0.9),
                 font_size=30, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

    key_points = [
        "STRATA / Prisma SASE / Cortex の3本柱でポートフォリオを構成",
        "NGFWは本体 + サブスクリプション（バンドルで割安）+ サポートの3層構造",
        "VMシリーズはFlexクレジットで柔軟にリソース・機能を選択可能",
        "SASEはPrisma Accessでユーザー・拠点課金（3年契約推奨）",
        "AIOps Premium はEOS（2025年5月）→ Strata Cloud Manager へ移行",
    ]
    tb = slide.shapes.add_textbox(Inches(2.5), Inches(3.1), Inches(8.5), Inches(2.8))
    tb.text_frame.word_wrap = True
    for j, kp in enumerate(key_points):
        p = tb.text_frame.paragraphs[0] if j == 0 else tb.text_frame.add_paragraph()
        p.space_before = Pt(8)
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = f"✓  {kp}"
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)

    add_text_box(slide,
                 "詳細はパートナー / Palo Alto Networks 営業担当まで",
                 Inches(1.0), Inches(6.4), Inches(11), Inches(0.6),
                 font_size=16, color=RGBColor(0xAA, 0xAA, 0xAA), align=PP_ALIGN.CENTER)
    return slide


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────
def main():
    prs = new_prs()

    slide_title(prs)
    slide_toc(prs)
    slide_portfolio(prs)
    slide_ngfw_overview(prs)
    slide_bundles(prs)
    slide_cdss(prs)
    slide_software_ngfw_credits(prs)
    slide_prisma_access(prs)
    slide_cortex(prs)
    slide_aiops_strata(prs)
    slide_gp_sdwan(prs)
    slide_support(prs)
    slide_summary_table(prs)
    slide_closing(prs)

    out_path = "/home/user/claude/paloalto_license_guide.pptx"
    prs.save(out_path)
    print(f"Saved: {out_path}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
