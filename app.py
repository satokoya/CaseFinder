import os
import sqlite3
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename
from pptx import Presentation
import PyPDF2
from openpyxl import load_workbook
from io import BytesIO

# アプリケーションの初期化
app = Flask(__name__)
app.secret_key = 'dev'  # MVP では簡易キー。運用時は環境変数などで管理。

# アップロードされたファイルを保存するディレクトリ
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 許可するファイル拡張子
ALLOWED_EXTENSIONS = {'pptx', 'pdf', 'xlsx', 'xls'}

# SQLite 設定
DB_PATH = 'casefinder.db'

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # Check if cases table exists
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='cases'")
    table_exists = c.fetchone()
    
    if not table_exists:
        # Create new tables with created_at column
        c.execute(
            """
            CREATE TABLE cases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                text_content TEXT NOT NULL,
                customer_name TEXT,
                system_name TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        c.execute(
            """
            CREATE TABLE summaries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                case_id INTEGER NOT NULL,
                problem TEXT,
                solution TEXT,
                outcome TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(case_id) REFERENCES cases(id)
            )
            """
        )
    else:
        # Add missing columns to cases table if they don't exist
        c.execute("PRAGMA table_info(cases)")
        columns = [column[1] for column in c.fetchall()]
        
        if 'created_at' not in columns:
            c.execute("ALTER TABLE cases ADD COLUMN created_at DATETIME DEFAULT CURRENT_TIMESTAMP")
        if 'customer_name' not in columns:
            c.execute("ALTER TABLE cases ADD COLUMN customer_name TEXT")
        if 'system_name' not in columns:
            c.execute("ALTER TABLE cases ADD COLUMN system_name TEXT")
    
    # Ensure summaries table exists (for backward compatibility)
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS summaries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_id INTEGER NOT NULL,
            problem TEXT,
            solution TEXT,
            outcome TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(case_id) REFERENCES cases(id)
        )
        """
    )
    
    conn.commit()
    conn.close()

# Initialize database with the latest schema
init_db()

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(file_path):
    """Extract text content from PDF file."""
    text = []
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text.append(page.extract_text() or '')
        return '\n'.join(text)
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")
        return ""

def extract_text_from_excel(file_path):
    """Extract text content from Excel file."""
    text = []
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                row_text = [str(cell) for cell in row if cell is not None]
                if row_text:  # Only add non-empty rows
                    text.append(' '.join(row_text))
        return '\n'.join(text)
    except Exception as e:
        print(f"Error extracting text from Excel: {e}")
        return ""

# --- Rule-based summary extraction -----------------------------------------

def extract_summary(text: str):
    """Extract problem, solution, and outcome from text using keyword matching."""
    # Split text into lines and clean them
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        return '', '', ''

    # Common slide titles to exclude
    exclude_phrases = ['表紙', '目次', 'agenda', 'contents', 'summary', 'conclusion', 'appendix', 'reference']
    
    # Keywords for each section (expanded list)
    problem_keywords = ['課題', '問題', '現状', '背景', 'challenge', 'problem', 'issue', 'background', 'current situation']
    solution_keywords = ['提案', '解決策', '解決方法', 'アプローチ', 'solution', 'proposal', 'approach', 'recommendation', 'methodology']
    outcome_keywords = ['効果', '成果', '実績', '効果測定', 'benefit', 'result', 'outcome', 'achievement', 'improvement', 'impact', 'kpi', '導入効果', '導入成果']
    
    # Words that indicate confidential or template content to exclude
    exclude_content = ['confidential', 'proprietary', '社外秘', '機密', 'template', 'サンプル']
    
    # Initialize variables
    problem = ''
    solution = ''
    outcome = ''
    
    # First pass: Look for section headers and content
    for i, line in enumerate(lines):
        line_lower = line.lower()
        
        # Skip common slide titles and very short lines (likely headers)
        if (any(phrase in line_lower for phrase in exclude_phrases) or 
            len(line) < 3 or 
            line.endswith(('。', '。', '、', ':')) is False and len(line) < 10):
            continue
            
        # Check for problem section
        if not problem and any(kw in line_lower for kw in problem_keywords):
            # Take the current line and next 2-3 lines if they look like content
            content = [line]
            for j in range(i+1, min(i+4, len(lines))):
                next_line = lines[j].strip()
                if len(next_line) > 10:  # Only include substantial lines
                    content.append(next_line)
            problem = '\n'.join(content[:3])  # Max 3 lines
            continue
            
        # Check for solution section
        if not solution and any(kw in line_lower for kw in solution_keywords):
            content = [line]
            for j in range(i+1, min(i+4, len(lines))):
                next_line = lines[j].strip()
                if len(next_line) > 10:
                    content.append(next_line)
            solution = '\n'.join(content[:3])
            continue
            
        # Check for outcome section
        if not outcome and any(kw in line_lower for kw in outcome_keywords):
            # Skip if the line contains confidential markers
            if any(excl in line_lower for excl in exclude_content):
                continue
                
            content = [line]
            lines_added = 1
            # Look ahead up to 4 lines for content
            for j in range(i+1, min(i+5, len(lines))):
                next_line = lines[j].strip()
                # Skip short lines and confidential content
                if (len(next_line) > 10 and 
                    not any(excl in next_line.lower() for excl in exclude_content)):
                    content.append(next_line)
                    lines_added += 1
                    if lines_added >= 3:  # Max 3 lines total
                        break
            if lines_added > 0:  # Only use if we found some content
                outcome = '\n'.join(content[:3])
    
    # Second pass: If we missed any sections, try to find them in the text
    if not problem or not solution or not outcome:
        # Look for problem (first meaningful paragraph)
        if not problem:
            for line in lines:
                line = line.strip()
                if (len(line) > 20 and 
                    not any(phrase in line.lower() for phrase in exclude_phrases) and
                    not any(kw in line.lower() for kw in solution_keywords + outcome_keywords)):
                    problem = line
                    break
        
        # Look for solution (contains solution keywords or follows problem)
        if not solution:
            for i, line in enumerate(lines):
                line_lower = line.lower()
                if any(kw in line_lower for kw in solution_keywords):
                    content = [line]
                    for j in range(i+1, min(i+3, len(lines))):
                        next_line = lines[j].strip()
                        if len(next_line) > 10:
                            content.append(next_line)
                    solution = '\n'.join(content[:3])
                    break
        
        # Look for outcome (contains outcome keywords or follows solution)
        if not outcome:
            for i, line in enumerate(lines):
                line_lower = line.lower()
                if (any(kw in line_lower for kw in outcome_keywords) and 
                    not any(excl in line_lower for excl in exclude_content)):
                    content = [line]
                    lines_added = 1
                    for j in range(i+1, min(i+4, len(lines))):
                        next_line = lines[j].strip()
                        if (len(next_line) > 10 and 
                            not any(excl in next_line.lower() for excl in exclude_content)):
                            content.append(next_line)
                            lines_added += 1
                            if lines_added >= 3:
                                break
                    if lines_added > 0:
                        outcome = '\n'.join(content[:3])
                        break
    
    # Final fallback if we still don't have all sections
    if not problem and lines:
        problem = lines[0][:200]  # First line, max 200 chars
    if not solution and len(lines) > 1:
        solution = lines[1][:200]
    if not outcome and len(lines) > 2:
        outcome = lines[2][:200]
    
    return problem, solution, outcome

# ---------------------------------------------------------------------------

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
@app.route('/search')
def index():
    """トップページ & 検索結果表示"""
    keyword = request.args.get('q', '').strip()
    conn = get_db_connection()
    cur = conn.cursor()
    if keyword:
        like = f"%{keyword}%"
        cur.execute(
            """
            SELECT id, filename, stored_path, text_content, 
                   strftime('%Y-%m-%d %H:%M:%S', created_at) as created_at 
            FROM cases 
            WHERE filename LIKE ? OR text_content LIKE ? 
            ORDER BY id DESC
            """,
            (like, like),
        )
    else:
        cur.execute(
            """
            SELECT id, filename, stored_path, text_content, 
                   strftime('%Y-%m-%d %H:%M:%S', created_at) as created_at 
            FROM cases 
            ORDER BY id DESC
            """
        )
    cases = cur.fetchall()
    # 全件数取得
    cur.execute("SELECT COUNT(*) FROM cases")
    total_count = cur.fetchone()[0]

    print(f'SEARCH "{keyword}" -> {len(cases)} / {total_count} hit(s)')
    conn.close()
    return render_template('index.html', cases=cases, keyword=keyword, total_count=total_count)

@app.route('/case/<int:case_id>')
def case_detail(case_id):
    conn = get_db_connection()
    case = conn.execute(
        """
        SELECT id, filename, stored_path, text_content, 
               customer_name, system_name,
               strftime('%Y-%m-%d %H:%M:%S', created_at) as created_at 
        FROM cases 
        WHERE id = ?
        """, 
        (case_id,)
    ).fetchone()
    summary = conn.execute(
        """
        SELECT problem, solution, outcome, 
               strftime('%Y-%m-%d %H:%M:%S', created_at) as created_at 
        FROM summaries 
        WHERE case_id = ? 
        ORDER BY id DESC 
        LIMIT 1
        """, 
        (case_id,)
    ).fetchone()
    conn.close()
    if case is None:
        flash('事例が見つかりません')
        return redirect(url_for('index'))
    return render_template('case_detail.html', case=case, summary=summary)

@app.route('/case/<int:case_id>/save', methods=['POST'])
def save_summary(case_id):
    problem = request.form.get('problem', '').strip()
    solution = request.form.get('solution', '').strip()
    outcome = request.form.get('outcome', '').strip()
    if not (problem and solution and outcome):
        flash('全ての項目を入力してください')
        return redirect(url_for('case_detail', case_id=case_id))

    conn = get_db_connection()
    conn.execute(
        "INSERT INTO summaries (case_id, problem, solution, outcome) VALUES (?, ?, ?, ?)",
        (case_id, problem, solution, outcome),
    )
    conn.commit()
    conn.close()
    flash('要約を保存しました')
    return redirect(url_for('case_detail', case_id=case_id))

@app.route('/case/<int:case_id>/auto', methods=['POST'])
def auto_summary(case_id):
    """Generate summary using rule-based extraction."""
    conn = get_db_connection()
    case = conn.execute("SELECT * FROM cases WHERE id = ?", (case_id,)).fetchone()
    conn.close()
    if case is None:
        flash('事例が見つかりません')
        return redirect(url_for('index'))
    problem, solution, outcome = extract_summary(case['text_content'])
    if not (problem or solution or outcome):
        flash('自動要約できませんでした')
        return redirect(url_for('case_detail', case_id=case_id))
    conn = get_db_connection()
    conn.execute("INSERT INTO summaries (case_id, problem, solution, outcome) VALUES (?, ?, ?, ?)",
                 (case_id, problem, solution, outcome))
    conn.commit()
    conn.close()
    flash('自動要約を保存しました')
    return redirect(url_for('case_detail', case_id=case_id))

@app.route('/case/<int:case_id>/save_metadata', methods=['POST'])
def save_metadata(case_id):
    customer_name = request.form.get('customer_name', '').strip()
    system_name = request.form.get('system_name', '').strip()
    
    conn = get_db_connection()
    conn.execute(
        "UPDATE cases SET customer_name = ?, system_name = ? WHERE id = ?",
        (customer_name, system_name, case_id)
    )
    conn.commit()
    conn.close()
    
    return jsonify({'status': 'success'})

@app.route('/case/<int:case_id>/pdf')
def download_pdf(case_id):
    from io import BytesIO
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
    from reportlab.lib.units import mm
    from reportlab.lib import colors

    # 日本語フォントを登録
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    
    # スタイルを定義
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='JapaneseTitle',
        fontName='HeiseiKakuGo-W5',
        fontSize=16,
        leading=20,
        spaceAfter=20,
        alignment=1  # 中央揃え
    ))
    styles.add(ParagraphStyle(
        name='JapaneseHeading',
        fontName='HeiseiKakuGo-W5',
        fontSize=12,
        leading=18,
        spaceAfter=6,
        spaceBefore=12,
        textColor=colors.HexColor('#2C3E50')
    ))
    styles.add(ParagraphStyle(
        name='JapaneseText',
        fontName='HeiseiKakuGo-W5',
        fontSize=10,
        leading=16,
        spaceAfter=12,
        leftIndent=10,
        textColor=colors.HexColor('#2C3E50')
    ))

    conn = get_db_connection()
    case = conn.execute("SELECT * FROM cases WHERE id = ?", (case_id,)).fetchone()
    summary = conn.execute("SELECT * FROM summaries WHERE case_id = ? ORDER BY id DESC LIMIT 1", (case_id,)).fetchone()
    conn.close()
    
    if summary is None:
        flash('要約がありません')
        return redirect(url_for('case_detail', case_id=case_id))

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, 
                          rightMargin=40, leftMargin=40,
                          topMargin=40, bottomMargin=40)
    
    # ストーリー（コンテンツ）を作成
    story = []
    
    # タイトル（顧客名とシステム名があれば使用、なければファイル名を使用）
    title_text = ''
    customer_name = case['customer_name'] if 'customer_name' in case.keys() and case['customer_name'] else None
    system_name = case['system_name'] if 'system_name' in case.keys() and case['system_name'] else None
    
    if customer_name and system_name:
        title_text = f"{customer_name}_{system_name}"
    else:
        title_text = f"事例要約: {case['filename']}"
    
    title = f'<para align=center><b>{title_text}</b></para>'
    story.append(Paragraph(title, styles['JapaneseTitle']))
    story.append(Spacer(1, 20))
    
    # ファイル名をサブタイトルとして表示
    if customer_name and system_name:
        subtitle = f'<para align=center><font size=10>ファイル名: {case["filename"]}</font></para>'
        story.append(Paragraph(subtitle, styles['JapaneseText']))
        story.append(Spacer(1, 15))
    
    # 要約セクション
    sections = [
        ('<b>顧客の課題</b>', summary['problem'] or '要約がありません'),
        ('<b>提供した解決策</b>', summary['solution'] or '要約がありません'),
        ('<b>導入後の成果</b>', summary['outcome'] or '要約がありません')
    ]
    
    for title, content in sections:
        # セクションタイトル
        story.append(Paragraph(title, styles['JapaneseHeading']))
        # コンテンツ（改行を<br/>に変換してから処理）
        if content:
            content = content.replace('\n', '<br/>')
            story.append(Paragraph(content, styles['JapaneseText']))
        else:
            story.append(Paragraph('要約がありません', styles['JapaneseText']))
        story.append(Spacer(1, 10))
    
    # フッターに作成日を追加
    from datetime import datetime
    created_at = datetime.now().strftime('%Y年%m月%d日 %H:%M')
    footer = f'<para align=right><font size=8>生成日時: {created_at}</font></para>'
    story.append(Spacer(1, 20))
    story.append(Paragraph(footer, styles['JapaneseText']))
    
    # PDFを生成
    doc.build(story)
    
    # レスポンスを準備
    buffer.seek(0)
    filename = f"case_summary_{case_id}_{datetime.now().strftime('%Y%m%d')}.pdf"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/pdf'
    )

    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f'summary_{case_id}.pdf', mimetype='application/pdf')


@app.route('/case/<int:case_id>/delete', methods=['POST'])
def delete_case(case_id):
    conn = get_db_connection()
    case = conn.execute('SELECT * FROM cases WHERE id = ?', (case_id,)).fetchone()
    
    if case is None:
        flash('削除するファイルが見つかりませんでした')
        return redirect(url_for('index'))
    
    try:
        # ファイルを削除
        if os.path.exists(case['stored_path']):
            os.remove(case['stored_path'])
        
        # 関連する要約を削除
        conn.execute('DELETE FROM summaries WHERE case_id = ?', (case_id,))
        # ケースを削除
        conn.execute('DELETE FROM cases WHERE id = ?', (case_id,))
        conn.commit()
        flash('ファイルを削除しました')
    except Exception as e:
        conn.rollback()
        flash(f'ファイルの削除中にエラーが発生しました: {str(e)}')
    finally:
        conn.close()
    
    return redirect(url_for('index'))

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash('ファイルが選択されていません')
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash('対応していないファイル形式です')
        return redirect(url_for('index'))

    filename = secure_filename(file.filename)
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    # 重複ファイル名の処理
    if os.path.exists(save_path):
        base, ext = os.path.splitext(filename)
        counter = 1
        while os.path.exists(save_path):
            filename = f"{base}_{counter}{ext}"
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            counter += 1
    
    file.save(save_path)

    # ファイルタイプに応じてテキスト抽出
    file_ext = filename.rsplit('.', 1)[1].lower()
    full_text = ""
    
    try:
        if file_ext == 'pptx':
            # PowerPoint のテキスト抽出
            presentation = Presentation(save_path)
            texts = []
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        texts.append(shape.text)
            full_text = "\n".join(texts)
            
        elif file_ext == 'pdf':
            # PDF のテキスト抽出
            full_text = extract_text_from_pdf(save_path)
            
        elif file_ext in ('xlsx', 'xls'):
            # Excel のテキスト抽出
            full_text = extract_text_from_excel(save_path)
            
        if not full_text.strip():
            raise ValueError("ファイルからテキストを抽出できませんでした")
            
    except Exception as e:
        flash(f'ファイル解析に失敗しました: {str(e)}')
        if os.path.exists(save_path):
            os.remove(save_path)  # 保存したファイルを削除
        return redirect(url_for('index'))

    # DB へ保存
    conn = get_db_connection()
    conn.execute(
        "INSERT INTO cases (filename, stored_path, text_content) VALUES (?, ?, ?)",
        (filename, save_path, full_text),
    )
    conn.commit()
    conn.close()

    flash('アップロードが完了しました')
    return redirect(url_for('index'))

if __name__ == '__main__':
    # 外部からのアクセスを許可し、デバッグモードを有効化
    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)
