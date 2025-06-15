import os
import markdown
from weasyprint import HTML, CSS

def create_pdf_with_weasyprint(txt_path, pdf_path, font_dir):
    # テキストファイルを読み込み
    try:
        with open(txt_path, 'r', encoding='utf-8') as f:
            md_text = f.read()
    except Exception as e:
        print(f"ファイルの読み込みに失敗しました: {txt_path}, error: {e}")
        return

    # マークダウンをHTMLに変換
    html_text = markdown.markdown(md_text, extensions=['tables'])

    # WeasyPrint用のCSSを定義。file:/// を使った絶対パス指定が最も確実。
    regular_font_path = os.path.join(font_dir, 'NotoSansJP-Regular.ttf').replace('\\', '/')
    bold_font_path = os.path.join(font_dir, 'NotoSansJP-Bold.ttf').replace('\\', '/')

    css_string = f"""
    @font-face {{
        font-family: 'NotoSansJP';
        src: url('file:///{regular_font_path}');
        font-weight: normal;
        font-style: normal;
    }}
    @font-face {{
        font-family: 'NotoSansJP';
        src: url('file:///{bold_font_path}');
        font-weight: bold;
        font-style: normal;
    }}
    body {{
        font-family: 'NotoSansJP', sans-serif;
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
    
    # HTMLとCSSからPDFを生成
    try:
        html = HTML(string=html_text)
        css = CSS(string=css_string)
        html.write_pdf(pdf_path, stylesheets=[css])
    except Exception as e:
        print(f"PDFの生成中にエラーが発生しました: {pdf_path}, error: {e}")
        # GTKランタイムがない場合のエラーメッセージを補足
        if 'no library called "pango-1.0"' in str(e) or 'DLL not found' in str(e):
             print("\\n*** ヒント: WeasyPrintの実行に必要なGTKランタイムがインストールされていない可能性があります。***")
             print("WindowsにGTKをインストールしてください。詳細はWeasyPrintのドキュメントをご確認ください。")
             print("https://weasyprint.readthedocs.io/en/latest/install.html#windows")

def main():
    source_dir = 'D:/健康管理/提供資料案'
    output_dir = 'D:/健康管理/提供資料案_pdf'
    font_dir = 'D:/健康管理'

    # 出力ディレクトリを（再）作成
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    files_to_convert = [f for f in os.listdir(source_dir) if f.endswith('.txt')]
    print(f"WeasyPrintを使用して{len(files_to_convert)}個のファイルを変換します...")

    for filename in files_to_convert:
        txt_path = os.path.join(source_dir, filename).replace('\\', '/')
        pdf_filename = os.path.splitext(filename)[0] + '.pdf'
        pdf_path = os.path.join(output_dir, pdf_filename).replace('\\', '/')
        
        print(f'変換中: {txt_path} -> {pdf_path}')
        create_pdf_with_weasyprint(txt_path, pdf_path, font_dir)
    
    print('完了しました。')

if __name__ == '__main__':
    main() 