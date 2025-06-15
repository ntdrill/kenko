import os
import markdown
from xhtml2pdf import pisa
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def create_perfect_pdf(txt_path, pdf_path, font_dir):
    # 1. ReportLabにフォントを直接登録
    regular_font_path = os.path.join(font_dir, 'NotoSansJP-Regular.ttf').replace('\\', '/')
    bold_font_path = os.path.join(font_dir, 'NotoSansJP-Bold.ttf').replace('\\', '/')
    
    pdfmetrics.registerFont(TTFont('NotoSansJP-Regular', regular_font_path))
    pdfmetrics.registerFont(TTFont('NotoSansJP-Bold', bold_font_path))
    pdfmetrics.registerFontFamily('NotoSansJP', normal='NotoSansJP-Regular', bold='NotoSansJP-Bold')

    # テキストファイルを読み込み
    try:
        with open(txt_path, 'r', encoding='utf-8') as f:
            md_text = f.read()
    except Exception as e:
        print(f"ファイルの読み込みに失敗しました: {txt_path}, error: {e}")
        return

    html_text = markdown.markdown(md_text, extensions=['tables'])

    # 2. CSSのパス指定を単純なフルパスにし、font-weightも定義
    css = f"""
    @font-face {{
        font-family: 'NotoSansJP';
        src: url('{regular_font_path}');
        font-weight: normal;
        font-style: normal;
    }}
    @font-face {{
        font-family: 'NotoSansJP';
        src: url('{bold_font_path}');
        font-weight: bold;
        font-style: normal;
    }}
    body {{
        font-family: 'NotoSansJP';
        font-size: 10pt;
        line-height: 1.5;
    }}
    h1, h2, h3, strong, b {{
        font-weight: bold;
    }}
    table {{
        border-collapse: collapse;
        width: 100%;
        margin-top: 10pt;
        margin-bottom: 10pt;
    }}
    th, td {{
        border: 1px solid #dddddd;
        text-align: left;
        padding: 8px;
    }}
    th {{
        background-color: #f2f2f2;
    }}
    """

    full_html = f"<html><head><meta charset=\"UTF-8\"><style>{css}</style></head><body>{html_text}</body></html>"

    try:
        with open(pdf_path, "w+b") as pdf_file:
            # 3. pisa.CreatePDFの呼び出し方を修正
            pisa_status = pisa.CreatePDF(
                full_html,
                dest=pdf_file,
                encoding='utf-8')
        
        if pisa_status.err:
            print(f"PDFの生成に失敗しました: {pdf_path}, error: {pisa_status.err}")
    except Exception as e:
        print(f"PDFの書き込み中にエラーが発生しました: {pdf_path}, error: {e}")

def main():
    source_dir = 'D:/健康管理/提供資料案'
    output_dir = 'D:/健康管理/提供資料案_pdf'
    font_dir = 'D:/健康管理'

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    files_to_convert = [f for f in os.listdir(source_dir) if f.endswith('.txt')]
    print(f"{len(files_to_convert)}個のファイルを変換します...")

    for filename in files_to_convert:
        txt_path = os.path.join(source_dir, filename).replace('\\', '/')
        pdf_filename = os.path.splitext(filename)[0] + '.pdf'
        pdf_path = os.path.join(output_dir, pdf_filename).replace('\\', '/')
        
        print(f'変換中: {txt_path} -> {pdf_path}')
        create_perfect_pdf(txt_path, pdf_path, font_dir)
    
    print('完了しました。')

if __name__ == '__main__':
    main() 