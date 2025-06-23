import streamlit as st
import re
import os
import subprocess
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
# === CÁC DÒNG ĐƯỢC THÊM VÀO ĐỂ SỬA LỖI ===
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
# ============================================
import io
import shutil

# --- CLASS CHÍNH: Chứa toàn bộ logic xử lý LaTeX sang Word ---
TABLE_PLACEHOLDER = "__TABLE_PLACEHOLDER_{}__"

class LaTeXToWordConverter:
    """
    Lớp chính để thực hiện việc chuyển đổi từ các bài tập định dạng LaTeX
    sang tài liệu Microsoft Word (.docx).
    """
    def __init__(self):
        """Khởi tạo một thư mục tạm để lưu các file trung gian (ảnh, pdf)."""
        self.temp_dir = tempfile.mkdtemp()

    def extract_exercises(self, latex_content):
        """Trích xuất tất cả các khối bài tập được bao bởi \begin{ex}...\end{ex}."""
        ex_pattern = r'\\begin\{ex\}(.*?)\\end\{ex\}'
        matches = re.finditer(ex_pattern, latex_content, re.DOTALL)
        return [match.group(1).strip() for match in matches if match.group(1).strip()]

    def parse_exercise(self, exercise_content):
        """Phân tích một khối bài tập để lấy ra các thành phần: câu hỏi, lựa chọn, đáp án, hình ảnh, lời giải."""
        parse_target = exercise_content
        immini_match = re.search(r'\\immini\{(.*?)\}', exercise_content, re.DOTALL)
        if immini_match:
            parse_target = immini_match.group(1).strip()

        question_parts = re.split(r'\\choice', parse_target, maxsplit=1)
        question = question_parts[0].strip()

        choices = []
        correct_choice_index = -1
        choice_block_match = re.search(r'\\choice\s*(.*?)(?=\\begin\{tikzpicture\}|\\loigiai|\\end\{ex\}|$)', 
                                       exercise_content, re.DOTALL)
        if choice_block_match:
            choices_text = choice_block_match.group(1)
            choice_pattern = r'\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'
            raw_choices = re.findall(choice_pattern, choices_text)
            for choice in raw_choices:
                choice = choice.strip()
                if choice:
                    if choice.startswith('\\True'):
                        correct_choice_index = len(choices)
                        choice = re.sub(r'^\\True\s*', '', choice).strip()
                    choices.append(choice)
        
        tikz_match = re.search(r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}', exercise_content, re.DOTALL)
        tikz = tikz_match.group(0) if tikz_match else None
        if tikz:
             question = question.replace(tikz, "")

        solution_match = re.search(r'\\loigiai\{(.*?)\}', exercise_content, re.DOTALL)
        solution = solution_match.group(1).strip() if solution_match else None

        return {'question': question, 'choices': choices, 'correct_choice': correct_choice_index, 'tikz': tikz, 'solution': solution}

    def _process_content_for_placeholders(self, content):
        tables = []
        def replacer(match):
            tables.append(match.group(0))
            return TABLE_PLACEHOLDER.format(len(tables) - 1)
        
        pattern = r'\\begin\{tabular\}.*?\\end\{tabular\}'
        content_with_placeholders = re.sub(pattern, replacer, content, flags=re.DOTALL)
        return content_with_placeholders, tables

    def _write_content_block(self, doc, content, prefix=""):
        if not content or not content.strip():
            if prefix:
                doc.add_paragraph().add_run(prefix).bold = True
            return

        content_ph, tables_latex = self._process_content_for_placeholders(content)
        parts = re.split(f'({TABLE_PLACEHOLDER.format("[0-9]+")})', content_ph)
        
        first_text_part = True
        for part in parts:
            if not part: continue
            
            placeholder_match = re.match(f'{TABLE_PLACEHOLDER.format("([0-9]+)")}', part)
            if placeholder_match:
                table_index = int(placeholder_match.group(1))
                if table_index < len(tables_latex):
                    self._latex_table_to_word_table(doc, tables_latex[table_index])
            else:
                prepared_text = self.prepare_latex_for_word(part)
                if prepared_text:
                    if first_text_part:
                        para = doc.add_paragraph()
                        if prefix:
                            para.add_run(prefix).bold = True
                        para.add_run(" " + prepared_text)
                        first_text_part = False
                    else:
                        doc.add_paragraph(prepared_text)

        if first_text_part and prefix:
             doc.add_paragraph().add_run(prefix).bold = True

    def _latex_table_to_word_table(self, doc, latex_table):
        spec_match = re.search(r'\\begin\{tabular\}(\{.*?\})', latex_table)
        body_match = re.search(r'(\{.*?\})(.*)\\end\{tabular\}', latex_table, re.DOTALL)
        if not spec_match or not body_match: return

        col_count = len(re.findall(r'[lcr]', spec_match.group(1)))
        body = body_match.group(2).strip()
        
        rows_data = []
        for line in body.split('\\\\'):
            line = line.replace('\\hline', '').strip()
            if not line: continue
            cells = [self.prepare_latex_for_word(cell.strip()) for cell in line.split('&')]
            if len(cells) > 0:
                while len(cells) < col_count: cells.append('')
                rows_data.append(cells[:col_count])
        
        if not rows_data: return

        table = doc.add_table(rows=len(rows_data), cols=col_count)
        table.style = 'Table Grid'
        
        for i, row_cells in enumerate(rows_data):
            for j, cell_text in enumerate(row_cells):
                if j < len(table.rows[i].cells):
                    table.cell(i, j).text = cell_text
                    if i == 0:
                        for p in table.cell(i, j).paragraphs:
                            for run in p.runs: run.bold = True
    
    def prepare_latex_for_word(self, text):
        text = re.sub(r'\\begin\{(center|align|align\*)\}', '', text, flags=re.DOTALL)
        text = re.sub(r'\\end\{(center|align|align\*)\}', '', text, flags=re.DOTALL)
        text = re.sub(r'\\vspace\{.*?\}', '', text)
        text = re.sub(r'\\item', '•', text)
        text = re.sub(r'\\begin\{itemize\}', '', text)
        text = re.sub(r'\\end\{itemize\}', '', text)
        text = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\textit\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)
        text = text.replace('\\\\', ' ')
        text = text.replace('\\hline', '')
        return re.sub(r'\s+', ' ', text).strip()

    def compile_tikz_to_image(self, tikz_code, filename_base):
        latex_doc = f"\\documentclass[border=5pt]{{standalone}}\n\\usepackage{{tikz}}\n\\usepackage{{amsmath}}\n\\usepackage{{amssymb}}\n\\usetikzlibrary{{arrows.meta}}\n\\begin{{document}}\n{tikz_code}\n\\end{{document}}"
        tex_file = os.path.join(self.temp_dir, f"{filename_base}.tex")
        with open(tex_file, 'w', encoding='utf-8') as f: f.write(latex_doc)
            
        try:
            subprocess.run(['pdflatex', '-interaction=nonstopmode', '-output-directory', self.temp_dir, tex_file], capture_output=True, check=True, timeout=30)
            pdf_file = Path(self.temp_dir) / f"{filename_base}.pdf"
            subprocess.run(['pdftoppm', '-png', '-r', '300', '-singlefile', str(pdf_file), str(pdf_file.with_suffix(''))], capture_output=True, check=True, timeout=30)
            png_file = pdf_file.with_suffix('.png')
            return str(png_file) if png_file.exists() else None
        except FileNotFoundError as e:
            st.error(f"LỖI: Lệnh `{e.filename}` không tồn tại. Yêu cầu cài đặt LaTeX (MiKTeX, TeX Live) và Poppler, sau đó thêm vào PATH hệ thống.")
            return None
        except subprocess.CalledProcessError as e:
            st.error(f"LỖI khi biên dịch TikZ. Hãy kiểm tra lại mã TikZ của bạn. Chi tiết lỗi:\n{e.stderr.decode('utf-8', errors='ignore')}")
            return None
        except Exception as e:
            st.error(f"LỖI không mong muốn: {e}")
            return None
    
    def create_word_document(self, exercises):
        doc = Document()
        doc.add_heading('BÀI TẬP TRẮC NGHIỆM', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        
        for idx, ex in enumerate(exercises, 1):
            self._write_content_block(doc, ex['question'], prefix=f"Câu {idx}.")

            if ex['tikz']:
                image_file = self.compile_tikz_to_image(ex['tikz'], f'tikz_{idx}')
                if image_file:
                    p = doc.add_paragraph()
                    p.add_run().add_picture(image_file, width=Inches(3.5))
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for i, choice in enumerate(ex['choices']):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25)
                label_text = f'{chr(65 + i)}.\t'
                label_run = p.add_run(label_text)
                text_run = p.add_run(self.prepare_latex_for_word(choice))
                
                if ex['correct_choice'] == i:
                    label_run.bold = True
                    label_run.underline = True
                    text_run.underline = True
            
            if ex['solution']:
                doc.add_paragraph()
                self._write_content_block(doc, ex['solution'], prefix="Lời giải:")

            if idx < len(exercises):
                p = doc.add_paragraph()
                p_border = OxmlElement('w:pBdr')
                p_border.set(qn('w:bottom'), '{"w:val": "single", "w:sz": "6", "w:space": "1", "w:color": "auto"}')
                p.get_or_add_pPr().append(p_border)

        return doc

    def cleanup(self):
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)

# --- HÀM MAIN: Dựng giao diện người dùng với Streamlit ---
def main():
    st.set_page_config(page_title="LaTeX to Word Converter", page_icon="📝", layout="wide")
    st.title("Chuyển đổi LaTeX sang Word (Phiên bản Hoàn chỉnh)")
    st.markdown("Công cụ chuyển đổi các bài tập trắc nghiệm từ định dạng LaTeX sang Microsoft Word, đã sửa lỗi và cập nhật đầy đủ tính năng.")
    
    if 'latex_input' not in st.session_state:
        st.session_state.latex_input = r"""\begin{ex}
% Ví dụ 1: Câu hỏi có cả văn bản và nhiều bảng
Hai mẫu số liệu ghép nhóm $M_1, M_2$ có bảng tần số ghép nhóm như sau:
\begin{center}
	$M_1 \quad$\begin{tabular}{|c|c|c|c|c|c|}
		\hline Nhóm & {$[8 ; 10)$} & {$[10 ; 12)$} & {$[12 ; 14)$} & {$[14 ; 16)$} & {$[16 ; 18)$} \\
		\hline Tần số & 3 & 4 & 8 & 6 & 4 \\
		\hline
	\end{tabular}
\end{center}
\begin{center}
	$M_2 \quad$\begin{tabular}{|c|c|c|c|c|c|}
		\hline Nhóm & {$[8 ; 10)$} & {$[10 ; 12)$} & {$[12 ; 14)$} & {$[14 ; 16)$} & {$[16 ; 18)$} \\
		\hline Tần số & 6 & 8 & 16 & 12 & 8 \\
		\hline
	\end{tabular}
\end{center}
Gọi $s_1, s_2$ lần lượt là độ lệch chuẩn của mẫu số liệu ghép nhóm $M_1, M_2$. Phát biểu nào sau đây là đúng?
\choice
{\True $s_1=s_2$}
{$s_1=2 s_2$}
{$2 s_1=s_2$}
{$4 s_1=s_2$}
\loigiai{Dễ thấy tần số của mẫu $M_2$ gấp đôi tần số của mẫu $M_1$ ở mọi nhóm tương ứng, trong khi các giá trị đại diện của nhóm là như nhau. Do đó, phương sai và độ lệch chuẩn của hai mẫu này bằng nhau. Cụ thể, $s_1^2 = s_2^2 \approx 5.11$, suy ra $s_1=s_2$.}
\end{ex}

\begin{ex}
% Ví dụ 2: Câu hỏi có hình vẽ TikZ và lệnh \immini
\immini{
	Cho hàm số $y=\dfrac{a x+b}{c x+d}(c \neq 0, a d-b c \neq 0)$ có đồ thị như hình vẽ bên. Tiệm cận ngang của đồ thị hàm số là:
	\choice
	{$x=-1$}
	{\True $y=\dfrac{1}{2}$}
	{$y=-1$}
	{$x=\dfrac{1}{2}$}}
{
	\begin{tikzpicture}[scale=1.5,>=stealth]
		\draw[->] (-3.5,0)--(2.5,0) node[below]{$x$};
		\draw[->] (0,-2.5)--(0,3.5) node[right]{$y$};
        \node[below left] at (0,0) {$O$};
		\draw[dashed, thin](-1,-2.5)--(-1,3.5);
		\draw[dashed, thin](-3.5,0.5)--(2.5,0.5);
        \node[above left] at (-1, 0) {$-1$};
        \node[below right] at (0, 0.5) {$\frac{1}{2}$};
		\clip (-3.4,-2.4) rectangle (2.4,3.4);
		\draw[blue, thick, samples=200, domain=-3.4:-1.05] plot (\x, {0.5 + 1.5/(\x+1)});
		\draw[blue, thick, samples=200, domain=-0.95:2.4] plot (\x, {0.5 + 1.5/(\x+1)});
	\end{tikzpicture}
}
\end{ex}"""

    col1, col2 = st.columns(2, gap="large")
    
    with col1:
        st.subheader("📥 Dữ liệu đầu vào LaTeX")
        uploaded_file = st.file_uploader("Tải lên file .tex để thay thế nội dung bên dưới:", type=['tex'])
        if uploaded_file is not None:
            st.session_state.latex_input = uploaded_file.read().decode('utf-8')

        st.text_area("Nội dung LaTeX:", key='latex_input', height=500, help="Dán mã LaTeX của bạn vào đây hoặc tải lên một file .tex.")

    with col2:
        st.subheader("📤 Chuyển đổi và Tải về")
        st.write("Nhấn nút bên dưới để bắt đầu quá trình chuyển đổi sang file Word.")
        if st.button("🚀 Chuyển đổi sang Word", type="primary", use_container_width=True):
            if st.session_state.latex_input and st.session_state.latex_input.strip():
                converter = LaTeXToWordConverter()
                try:
                    with st.spinner("⏳ Đang xử lý... Quá trình này có thể mất một lúc nếu có nhiều hình ảnh..."):
                        exercises_raw = converter.extract_exercises(st.session_state.latex_input)
                        if not exercises_raw:
                            st.warning("Không tìm thấy bài tập nào trong mã LaTeX được cung cấp. Hãy chắc chắn bạn đã bọc chúng trong `\\begin{ex}...\\end{ex}`.")
                        else:
                            exercises_parsed = [converter.parse_exercise(ex) for ex in exercises_raw]
                            doc = converter.create_word_document(exercises_parsed)
                            
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            doc_io.seek(0)
                            
                            st.success(f"✅ Chuyển đổi thành công {len(exercises_parsed)} bài tập!")
                            st.download_button(
                                label="📥 Tải xuống file Word (.docx)",
                                data=doc_io.getvalue(),
                                file_name="Bai_tap_da_chuyen_doi.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                except Exception as e:
                    st.error(f"❌ Đã xảy ra lỗi nghiêm trọng trong quá trình xử lý.")
                    st.exception(e)
                finally:
                    converter.cleanup()
            else:
                st.warning("⚠️ Vui lòng nhập nội dung LaTeX hoặc tải file lên trước khi chuyển đổi.")
        
        with st.expander("📖 Xem Hướng dẫn sử dụng và Yêu cầu cài đặt", expanded=False):
            st.markdown("""
            ### Yêu cầu hệ thống (QUAN TRỌNG)
            Để có thể chuyển đổi **hình vẽ TikZ**, máy tính của bạn **bắt buộc** phải cài đặt:
            1.  **Một bản phân phối LaTeX**: Ví dụ như [**MiKTeX**](https://miktex.org/download) (cho Windows), **MacTeX** (cho macOS), hoặc **TeX Live** (cho Linux).
            2.  **Poppler**: Cung cấp công cụ để chuyển PDF sang ảnh. Bạn có thể tải Poppler cho Windows [tại đây](https://github.com/oschwartz10612/poppler-windows/releases/).
            
            **Lưu ý**: Sau khi cài đặt, hãy đảm bảo các thư mục chứa `pdflatex.exe` và `pdftoppm.exe` đã được thêm vào **biến môi trường `PATH`** của hệ thống và khởi động lại máy.

            ### Cấu trúc LaTeX được hỗ trợ
            - **Khối bài tập**: Bắt đầu bằng `\\begin{ex}` và kết thúc bằng `\\end{ex}`.
            - **Câu hỏi**: Nội dung bên trong `\\immini{...}` hoặc phần văn bản tự do trước lệnh `\\choice`.
            - **Lựa chọn**: Đặt sau lệnh `\\choice`, mỗi lựa chọn nằm trong một cặp `{...}`.
            - **Đáp án đúng**: Thêm `\\True` vào đầu lựa chọn đúng, ví dụ: `{\\True Đáp án đúng}`.
            - **Công thức toán**: Sẽ được **giữ nguyên** (ví dụ: `$x^2+y^2=z^2$`).
            - **Hình vẽ TikZ**: Đặt trong môi trường `\\begin{tikzpicture}...\\end{tikzpicture}`.
            - **Bảng**: Đặt trong môi trường `\\begin{tabular}{...}...\\end{tabular}`.
            - **Lời giải**: Đặt trong lệnh `\\loigiai{...}`.
            """)

if __name__ == "__main__":
    main()
