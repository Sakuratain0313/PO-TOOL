# -*- coding: utf-8 -*-
"""
AWAY PO 產生工具
讀「大貨訂單總表」某個 Sheet,依 Destination(出貨地)篩選,
每個(工廠訂單#, 客人PO#, Destination)各出一份 Wilson Group Holdings PO PDF。
"""
import os
import io
import threading
from collections import defaultdict, OrderedDict

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable
)
from reportlab.pdfgen import canvas as canvas_mod

# ========================================================================
# 固定內容(照 Wilson Group Holdings PO 格式規則,可視需要調整)
# ========================================================================
COMPANY_NAME = 'WILSON GROUP HOLDINGS LIMITED'
COMPANY_COUNTRY = 'HONG KONG'
SUPPLIER_BLOCK = [
    'Wilson Leather(Cambodia) Co., LTD.',
    'WILSON LEATHER (CAMBODIA) CO.,LTD',
    'THLORK VILLAGE, TROPAING KONG COMMUNE, SAMRONG',
    'TONG DISTRICT, KAMPONG SPEU PROVINCE, CAMBODIA.',
]
CONSIGNEE_LINES = [
    'JRSK, Inc dbc Away',
    '503 Broadway, 3rd Floor, New York, NY, 10012, United States',
]
PAYMENT_TERM = 'Net 45 days EOM'
TERMS = [
    'GOODS FOR EXPORT TO THE U.S. ONLY.',
    'WILSON GROUP HOLDINGS LIMITED TAKES TITLE AND RISK OF LOSS TO THE FINISHED GOODS AT THE FACTORY DOOR.',
]

# ========================================================================
# 數字轉英文大寫(SAY TOTAL DOLLARS)
# ========================================================================
ONES = ['', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE',
        'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN',
        'SEVENTEEN', 'EIGHTEEN', 'NINETEEN']
TENS = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY']

def three_digit(n):
    words = []
    if n >= 100:
        words.append(ONES[n // 100]); words.append('HUNDRED'); n %= 100
    if n >= 20:
        words.append(TENS[n // 10])
        if n % 10: words.append(ONES[n % 10])
    elif n > 0:
        words.append(ONES[n])
    return ' '.join(words)

def int_to_words(n):
    if n == 0: return 'ZERO'
    parts = []
    if n >= 1_000_000:
        parts.append(three_digit(n // 1_000_000) + ' MILLION'); n %= 1_000_000
    if n >= 1_000:
        parts.append(three_digit(n // 1_000) + ' THOUSAND'); n %= 1_000
    if n > 0:
        parts.append(three_digit(n))
    return ' '.join(parts)

def say_total_dollars(amount):
    dollars = int(round(amount * 100)) // 100
    cents = int(round(amount * 100)) % 100
    words = int_to_words(dollars)
    if cents == 0:
        return f'{words} ONLY.'
    return f'{words} AND CENTS {int_to_words(cents)} ONLY.'

def fmt_date(dt):
    return f'{dt.year}/{dt.month}/{dt.day}'

# ========================================================================
# Excel 欄位辨識(用關鍵字比對表頭,不依賴固定欄位順序,同一份總表換月份 Sheet 也能用)
# ========================================================================
COLUMN_KEYWORDS = {
    'factory_po': ['factory number'],
    'customer_po': ['客人po'],
    'order_date': ['下單日期'],
    'product_code': ['內部款號'],
    'description': ['description'],
    'color_en': ['color', '英文'],
    'balance_qty': ['balance qty'],
    'destination': ['destination'],
    'price': ['price單價'],
    'xf_date': ['po出貨日期'],
}
COLUMN_LABELS = {
    'factory_po': 'Factory number訂單編號',
    'customer_po': '客人po#',
    'order_date': '下單日期',
    'product_code': '內部款號',
    'description': 'Description描述',
    'color_en': 'Color顏色英文名稱',
    'balance_qty': 'Balance qty to Be Shipped 出貨数量',
    'destination': 'Destination 分出货地',
    'price': 'Price單價',
    'xf_date': 'PO出貨日期',
}

def find_column_map(ws, header_row=1):
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v:
            headers[c] = str(v).strip().lower()
    colmap = {}
    for key, keywords in COLUMN_KEYWORDS.items():
        for c, h in headers.items():
            if all(kw.lower() in h for kw in keywords):
                colmap[key] = c
                break
    missing = [COLUMN_LABELS[k] for k in COLUMN_KEYWORDS if k not in colmap]
    return colmap, missing

def scan_destinations(ws, colmap, data_start_row=3):
    col = colmap['destination']
    vals = set()
    for r in range(data_start_row, ws.max_row + 1):
        v = ws.cell(row=r, column=col).value
        if v:
            vals.add(str(v).strip())
    return sorted(vals)

def extract_groups(ws, colmap, selected_destinations, data_start_row=3):
    """回傳 { (factory_po, customer_po, dest): [row_dict, ...] }。
    Balance qty <= 0 或缺單價的品項列會直接排除;整組都被排除的組別不會出現在結果裡。"""
    def g(r, key):
        return ws.cell(row=r, column=colmap[key]).value

    raw_groups = defaultdict(list)
    for r in range(data_start_row, ws.max_row + 1):
        dest = g(r, 'destination')
        if not dest:
            continue
        dest = str(dest).strip()
        if dest not in selected_destinations:
            continue
        qty = g(r, 'balance_qty')
        price = g(r, 'price')
        if not qty or qty <= 0 or price is None:
            continue
        f = g(r, 'factory_po')
        c = g(r, 'customer_po')
        if not f or not c:
            continue
        raw_groups[(str(f).strip(), str(c).strip(), dest)].append({
            'order_date': g(r, 'order_date'),
            'xf_date': g(r, 'xf_date'),
            'code': g(r, 'product_code'),
            'desc': g(r, 'description'),
            'color': g(r, 'color_en'),
            'qty': qty,
            'price': price,
        })
    return raw_groups

# ========================================================================
# PDF 產生(reportlab,純 Python、不依賴 Chrome)
# ========================================================================
STYLE_COMPANY = ParagraphStyle('company', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, leading=19)
STYLE_COUNTRY = ParagraphStyle('country', fontName='Helvetica', fontSize=10, alignment=TA_CENTER, leading=13)
STYLE_POTITLE = ParagraphStyle('potitle', fontName='Helvetica-Bold', fontSize=15, alignment=TA_CENTER, leading=20)
STYLE_INFO_LEFT = ParagraphStyle('infoL', fontName='Helvetica', fontSize=10.5, alignment=TA_LEFT, leading=15)
STYLE_INFO_RIGHT = ParagraphStyle('infoR', fontName='Helvetica', fontSize=10.5, alignment=TA_LEFT, leading=15)
STYLE_TH = ParagraphStyle('th', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)
STYLE_TH_R = ParagraphStyle('thr', fontName='Helvetica-Bold', fontSize=10, alignment=TA_RIGHT)
STYLE_TD = ParagraphStyle('td', fontName='Helvetica', fontSize=10.5, alignment=TA_LEFT)
STYLE_TD_R = ParagraphStyle('tdr', fontName='Helvetica', fontSize=10.5, alignment=TA_RIGHT)
STYLE_TD_BOLD = ParagraphStyle('tdb', fontName='Helvetica-Bold', fontSize=10.5, alignment=TA_LEFT)
STYLE_TOTAL = ParagraphStyle('total', fontName='Helvetica-Bold', fontSize=10.5, alignment=TA_LEFT)
STYLE_TOTAL_R = ParagraphStyle('totalr', fontName='Helvetica-Bold', fontSize=10.5, alignment=TA_RIGHT)
STYLE_BODY = ParagraphStyle('body', fontName='Helvetica', fontSize=10.5, alignment=TA_LEFT, leading=15)
STYLE_VLINE = ParagraphStyle('vline', fontName='Helvetica', fontSize=9, alignment=TA_LEFT)
STYLE_SIG_LABEL = ParagraphStyle('siglabel', fontName='Helvetica', fontSize=10, alignment=TA_LEFT, leading=13)
STYLE_SIG_LINE = ParagraphStyle('sigline', fontName='Helvetica', fontSize=10, alignment=TA_LEFT)

PAGE_W, PAGE_H = A4
MARGIN_L = 16 * mm
MARGIN_R = 16 * mm
MARGIN_TOP = 18 * mm
MARGIN_BOTTOM = 16 * mm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R


def _build_header_flowables(po_no, away_po_no, po_date_str, xf_date_str, page_str):
    left_w = CONTENT_W * 0.60
    right_w = CONTENT_W * 0.38

    flow = []
    flow.append(Paragraph(COMPANY_NAME, STYLE_COMPANY))
    flow.append(Paragraph(COMPANY_COUNTRY, STYLE_COUNTRY))
    flow.append(Spacer(1, 4))
    flow.append(Paragraph('P&nbsp;U&nbsp;R&nbsp;C&nbsp;H&nbsp;A&nbsp;S&nbsp;E&nbsp;&nbsp;&nbsp;O&nbsp;R&nbsp;D&nbsp;E&nbsp;R', STYLE_POTITLE))
    flow.append(Spacer(1, 6))
    flow.append(HRFlowable(width=CONTENT_W, thickness=1.3, color=colors.black, spaceAfter=4))

    left1 = Paragraph(
        f"<b>TO (SUPPLIER) :</b> {SUPPLIER_BLOCK[0]}<br/>" + '<br/>'.join(SUPPLIER_BLOCK[1:]),
        STYLE_INFO_LEFT)
    right1 = Paragraph(f"<b>DATE :</b> {po_date_str}<br/><b>Page :</b> {page_str}", STYLE_INFO_RIGHT)
    t1 = Table([[left1, right1]], colWidths=[left_w, right_w])
    t1.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    flow.append(t1)
    flow.append(Spacer(1, 4))
    flow.append(HRFlowable(width=CONTENT_W, thickness=1.3, color=colors.black, spaceAfter=4))

    left2 = Paragraph(
        f"<b>PO NO :</b> {po_no}<br/>"
        f"<b>REF. SUPPLIER S/C NO :</b> {po_no}<br/>"
        f"<b>ULTIMATE CONSIGNEE :</b> {CONSIGNEE_LINES[0]}<br/>"
        + '<br/>'.join(CONSIGNEE_LINES[1:]),
        STYLE_INFO_LEFT)
    right2 = Paragraph(
        f"<b>XF-DATE :</b> {xf_date_str}<br/>"
        f"<b>PAYMENT TERM :</b> {PAYMENT_TERM}<br/>"
        f"<b>AWAY PO NO :</b> {away_po_no}",
        STYLE_INFO_RIGHT)
    t2 = Table([[left2, right2]], colWidths=[left_w, right_w])
    t2.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    flow.append(t2)
    flow.append(Spacer(1, 4))
    flow.append(HRFlowable(width=CONTENT_W, thickness=1.3, color=colors.black, spaceAfter=6))
    return flow


def _header_height_estimate():
    """實際跑一次 Frame 排版量出 header 高度(比直接加總 wrap() 準,
    因為 Frame 還會另外加上每個 flowable 的 spaceBefore/spaceAfter)。"""
    buf = io.BytesIO()
    dummy_canvas = canvas_mod.Canvas(buf, pagesize=A4)
    flow = _build_header_flowables('AWC-0000/26', 'POUS0000000', '2026/1/1', '2026/1/1', '1 / 9')
    top_y = PAGE_H - MARGIN_TOP
    frame = Frame(MARGIN_L, 0, CONTENT_W, top_y,
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, showBoundary=0)
    for f in flow:
        frame.add(f, dummy_canvas)
    consumed = top_y - frame._y
    return consumed + 6


def _build_item_story(groups_items):
    col_w = [CONTENT_W * 0.19, CONTENT_W * 0.29, CONTENT_W * 0.16, CONTENT_W * 0.16, CONTENT_W * 0.20]
    table_data = [[
        Paragraph('ITEM NO.', STYLE_TH), Paragraph('DESCRIPTION', STYLE_TH),
        Paragraph('QUANTITY', STYLE_TH_R), Paragraph('UNIT-PRICE', STYLE_TH_R),
        Paragraph('AMOUNT (USD)', STYLE_TH_R),
    ]]
    style_cmds = [
        ('LINEABOVE', (0, 0), (-1, 0), 1.3, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]

    total_qty = 0
    total_amount = 0.0
    row_idx = 1
    for (code, desc), colors_list in groups_items.items():
        table_data.append([Paragraph(code, STYLE_TD_BOLD), Paragraph(desc, STYLE_TD_BOLD), '', '', ''])
        style_cmds.append(('TOPPADDING', (0, row_idx), (-1, row_idx), 7))
        row_idx += 1
        for color, qty, price in colors_list:
            amount = round(qty * price, 2)
            total_qty += qty
            total_amount += amount
            table_data.append([
                '', Paragraph(color, STYLE_TD),
                Paragraph(f'{qty:,} PCS', STYLE_TD_R),
                Paragraph(f'{price:.4f}', STYLE_TD_R),
                Paragraph(f'{amount:,.2f}', STYLE_TD_R),
            ])
            row_idx += 1
    total_amount = round(total_amount, 2)

    table_data.append([
        Paragraph('TOTAL :', STYLE_TOTAL), '', Paragraph(f'{total_qty:,} PCS', STYLE_TOTAL_R),
        '', Paragraph(f'USD&nbsp;&nbsp;{total_amount:,.2f}', STYLE_TOTAL_R),
    ])
    style_cmds.append(('LINEABOVE', (0, row_idx), (-1, row_idx), 1.3, colors.black))
    style_cmds.append(('TOPPADDING', (0, row_idx), (-1, row_idx), 5))
    style_cmds.append(('SPAN', (0, row_idx), (1, row_idx)))

    items_table = Table(table_data, colWidths=col_w, repeatRows=1)
    items_table.setStyle(TableStyle(style_cmds))

    say_total = say_total_dollars(total_amount)

    story = [items_table, Spacer(1, 6), Paragraph('V' * 95, STYLE_VLINE), Spacer(1, 4),
             Paragraph(f'<b>SAY TOTAL DOLLARS :</b> {say_total}', STYLE_BODY), Spacer(1, 6)]
    for i, term in enumerate(TERMS, 1):
        story.append(Paragraph(f'{i}. <b>{term}</b>', STYLE_BODY))

    story.append(PageBreak())
    sig_left = [Spacer(1, 34), Paragraph('_' * 46, STYLE_SIG_LINE), Spacer(1, 2),
                Paragraph('ACCEPTED BY (SUPPLIER)', STYLE_SIG_LABEL)]
    sig_right = [Paragraph('<b>ISSUED BY :</b><br/>Wilson Group Holdings Limited', STYLE_SIG_LABEL),
                 Spacer(1, 20), Paragraph('_' * 40, STYLE_SIG_LINE), Spacer(1, 2),
                 Paragraph('AUTHORIZED SIGNATURE', STYLE_SIG_LABEL)]
    sig_table = Table([[sig_left, sig_right]], colWidths=[CONTENT_W * 0.5, CONTENT_W * 0.5])
    sig_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 60),
    ]))
    story.append(sig_table)
    story.append(Spacer(1, 40))
    story.append(Paragraph('PLEASE SIGN BACK AS SOON AS POSSIBLE', STYLE_BODY))
    return story, total_qty, total_amount, say_total


def build_po_pdf(out_path, po_no, away_po_no, po_date_str, xf_date_str, groups_items):
    """groups_items: dict{(code,desc): [(color, qty, price), ...]}(用一般 dict 保序即可)。
    兩階段輸出:先算總頁數,再正式輸出,讓「Page : X / Y」永遠正確。"""
    header_h = _header_height_estimate()
    main_h = PAGE_H - MARGIN_TOP - MARGIN_BOTTOM - header_h - 4

    def make_doc(path, page_str_fn):
        story, total_qty, total_amount, say_total = _build_item_story(groups_items)

        def on_page(c, doc):
            c.saveState()
            page_str = page_str_fn(doc.page)
            header_flow = _build_header_flowables(po_no, away_po_no, po_date_str, xf_date_str, page_str)
            frame = Frame(MARGIN_L, PAGE_H - MARGIN_TOP - header_h, CONTENT_W, header_h,
                          leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, showBoundary=0)
            frame.addFromList(list(header_flow), c)
            c.restoreState()

        doc = BaseDocTemplate(path, pagesize=A4,
                               leftMargin=MARGIN_L, rightMargin=MARGIN_R,
                               topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM)
        frame_main = Frame(MARGIN_L, MARGIN_BOTTOM, CONTENT_W, main_h,
                            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id='main')
        template = PageTemplate(id='po', frames=[frame_main], onPage=on_page)
        doc.addPageTemplates([template])
        doc.build(story)
        return doc.page, total_qty, total_amount, say_total

    tmp_path = out_path + '.tmp_pass1.pdf'
    total_pages, _, _, _ = make_doc(tmp_path, lambda p: f'{p} / ?')
    try:
        os.remove(tmp_path)
    except OSError:
        pass

    _, total_qty, total_amount, say_total = make_doc(out_path, lambda p: f'{p} / {total_pages}')
    return total_qty, total_amount, say_total


# ========================================================================
# 主流程:讀 Excel -> 分組 -> 逐組出 PDF
# ========================================================================
def generate_all(excel_path, sheet_name, selected_destinations, out_dir, log):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name]
    colmap, missing = find_column_map(ws)
    if missing:
        log('❌ 這個 Sheet 缺少必要欄位,無法產生:')
        for m in missing:
            log(f'   - {m}')
        return

    raw_groups = extract_groups(ws, colmap, selected_destinations)
    if not raw_groups:
        log('⚠️ 篩選後沒有任何資料(可能勾選的 Destination 都沒有 Balance qty > 0 的品項)')
        return

    log(f'共 {len(raw_groups)} 組(工廠訂單# + 客人PO# + Destination),開始產生 PDF...\n')
    os.makedirs(out_dir, exist_ok=True)

    ok = 0
    for (po_no, away_po_no, dest), rows in sorted(raw_groups.items()):
        order_dates = sorted({r['order_date'] for r in rows if r['order_date']})
        xf_dates = sorted({r['xf_date'] for r in rows if r['xf_date']})
        po_date_str = fmt_date(order_dates[0]) if order_dates else ''
        xf_date_str = ', '.join(fmt_date(d) for d in xf_dates) if xf_dates else ''

        items = OrderedDict()
        for r in rows:
            key = (r['code'], r['desc'])
            items.setdefault(key, []).append((r['color'], r['qty'], r['price']))

        safe_po = str(po_no).replace('/', '-')
        safe_dest = str(dest).replace(' ', '')
        fname = f'PO_{safe_po}_{away_po_no}_{safe_dest}.pdf'
        out_path = os.path.join(out_dir, fname)

        try:
            qty, amount, say = build_po_pdf(out_path, po_no, away_po_no, po_date_str, xf_date_str, items)
            log(f'✅ {fname}  —  {qty:,} PCS / USD {amount:,.2f}')
            ok += 1
        except Exception as e:
            log(f'❌ {fname} 產生失敗: {e}')

    log(f'\n完成!共產生 {ok} 份 PDF,存於:\n{out_dir}')


# ========================================================================
# GUI
# ========================================================================
class App:
    def __init__(self, root):
        root.title('AWAY PO 產生工具 v1')
        root.geometry('720x640')

        pad = {'padx': 10, 'pady': 6}

        self.v_excel = tk.StringVar()
        self.v_sheet = tk.StringVar()
        self.v_out = tk.StringVar()
        self.dest_vars = {}   # {dest_value: tk.BooleanVar}
        self.colmap = None
        self.ws = None

        # -- Excel 檔案 --
        tk.Label(root, text='Excel 檔案:', anchor='w', width=12).grid(row=0, column=0, **pad, sticky='w')
        tk.Entry(root, textvariable=self.v_excel, width=52).grid(row=0, column=1, **pad)
        tk.Button(root, text='選擇', width=8, command=self.browse_excel).grid(row=0, column=2, padx=(0, 10))

        # -- Sheet 選擇 --
        tk.Label(root, text='Sheet:', anchor='w', width=12).grid(row=1, column=0, **pad, sticky='w')
        self.sheet_combo = ttk.Combobox(root, textvariable=self.v_sheet, width=50, state='readonly')
        self.sheet_combo.grid(row=1, column=1, **pad, sticky='w')
        self.sheet_combo.bind('<<ComboboxSelected>>', lambda e: self.on_sheet_selected())

        # -- Destination 勾選區(動態產生)--
        tk.Label(root, text='Destination\n(出貨地):', anchor='nw', width=12, justify='left').grid(row=2, column=0, **pad, sticky='nw')
        dest_outer = tk.Frame(root, bd=1, relief='sunken')
        dest_outer.grid(row=2, column=1, columnspan=2, padx=10, pady=6, sticky='we')
        self.dest_canvas = tk.Canvas(dest_outer, height=140, width=560)
        dest_scroll = tk.Scrollbar(dest_outer, orient='vertical', command=self.dest_canvas.yview)
        self.dest_frame = tk.Frame(self.dest_canvas)
        self.dest_frame.bind('<Configure>', lambda e: self.dest_canvas.configure(scrollregion=self.dest_canvas.bbox('all')))
        self.dest_canvas.create_window((0, 0), window=self.dest_frame, anchor='nw')
        self.dest_canvas.configure(yscrollcommand=dest_scroll.set)
        self.dest_canvas.pack(side='left', fill='both', expand=True)
        dest_scroll.pack(side='right', fill='y')

        # -- 輸出資料夾 --
        tk.Label(root, text='輸出資料夾:', anchor='w', width=12).grid(row=3, column=0, **pad, sticky='w')
        tk.Entry(root, textvariable=self.v_out, width=52).grid(row=3, column=1, **pad)
        tk.Button(root, text='選擇', width=8, command=self.browse_folder).grid(row=3, column=2, padx=(0, 10))

        # -- 執行按鈕 --
        self.btn = tk.Button(root, text='▶  產生 PDF', font=('Arial', 13, 'bold'),
                              bg='#2e86de', fg='white', width=14, command=self.execute)
        self.btn.grid(row=4, column=0, columnspan=3, pady=12)

        # -- Log --
        self.log_box = tk.Text(root, height=18, width=86, state='disabled', bg='#f0f0f0', font=('Courier', 10))
        self.log_box.grid(row=5, column=0, columnspan=3, padx=10, pady=(0, 10))

    # ---- helpers ----
    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')])
        if not path:
            return
        self.v_excel.set(path)
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            self.sheet_combo['values'] = wb.sheetnames
            if wb.sheetnames:
                self.v_sheet.set(wb.sheetnames[0])
                self.on_sheet_selected()
        except Exception as e:
            messagebox.showerror('讀取失敗', str(e))

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.v_out.set(path)

    def on_sheet_selected(self):
        excel_path = self.v_excel.get().strip()
        sheet_name = self.v_sheet.get().strip()
        if not excel_path or not sheet_name:
            return
        for w in self.dest_frame.winfo_children():
            w.destroy()
        self.dest_vars = {}
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            ws = wb[sheet_name]
            colmap, missing = find_column_map(ws)
            if missing:
                tk.Label(self.dest_frame, fg='red',
                         text='這個 Sheet 缺少必要欄位:\n' + '\n'.join(missing),
                         justify='left').pack(anchor='w')
                return
            dests = scan_destinations(ws, colmap)
            for d in dests:
                var = tk.BooleanVar(value=False)
                self.dest_vars[d] = var
                tk.Checkbutton(self.dest_frame, text=d, variable=var).pack(anchor='w')
        except Exception as e:
            messagebox.showerror('讀取 Sheet 失敗', str(e))

    def log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.config(state='disabled')
        self.log_box.update()

    def execute(self):
        excel_path = self.v_excel.get().strip()
        sheet_name = self.v_sheet.get().strip()
        out_dir = self.v_out.get().strip()
        selected = [d for d, v in self.dest_vars.items() if v.get()]

        if not excel_path or not sheet_name:
            messagebox.showwarning('提示', '請先選擇 Excel 檔案跟 Sheet')
            return
        if not selected:
            messagebox.showwarning('提示', '請至少勾選一個 Destination')
            return
        if not out_dir:
            messagebox.showwarning('提示', '請選擇輸出資料夾')
            return

        self.log_box.config(state='normal')
        self.log_box.delete('1.0', 'end')
        self.log_box.config(state='disabled')
        self.btn.config(state='disabled', text='執行中...')

        def task():
            try:
                generate_all(excel_path, sheet_name, selected, out_dir, self.log)
            except Exception as e:
                self.log(f'❌ 發生錯誤: {e}')
            self.btn.config(state='normal', text='▶  產生 PDF')

        threading.Thread(target=task, daemon=True).start()


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
