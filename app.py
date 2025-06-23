import streamlit as st
import re
import os
import subprocess
import tempfile
import shutil
from pathlib import Path
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

def load_css():
    """Tải và áp dụng CSS tùy chỉnh để làm đẹp giao diện."""
    st.markdown("""
    <style>
        /* --- Bảng màu Teal Theme --- */
        :root {
            --primary-color: #17a2b8; /* Teal */
            --secondary-color: #f0f2f6; /* Màu nền xám nhạt */
            --text-color: #0c0c0c;
            --card-bg-color: #ffffff;
            --border-radius: 10px;
            --box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }

        /* --- Giao diện tổng thể --- */
        .stApp {
            background-color: var(--secondary-color);
        }

        /* --- Tiêu đề --- */
        h1, h2, h3 {
            color: var(--primary-color);
        }

        /* --- Thiết kế dạng Card cho các cột --- */
        .st-emotion-cache-1e5imcs > div {
            background-color: var(--card-bg-color);
            padding: 25px;
            border-radius: var(--border-radius);
            box-shadow: var(--box-shadow);
        }

        /* --- Nút bấm chính --- */
        .stButton > button {
            background-color: var(--primary-color);
            color: white;
            border-radius: 5px;
            border: none;
            padding: 10px 20px;
            transition: background-color 0.3s ease;
        }
        .stButton > button:hover {
            background-color: #138496; /* Teal đậm hơn khi hover */
        }
        
        /* --- Vùng nhập liệu --- */
        .stTextArea textarea {
            border-radius: var(--border-radius);
            border-color: #ced4da;
        }

        /* --- Expander (Hướng dẫn) --- */
        .stExpander {
            background-color: var(--card-bg-color);
            border-radius: var(--border-radius);
            border: 1px solid #e6e6e6;
        }
    </style>
    """, unsafe_allow_html=True)

class LaTeXToWordConverter:
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()

    def extract_exercises(self, latex_content):
        """Trích xuất tất cả các bài tập từ nội dung LaTeX."""
        pattern = r'\\begin\{ex\}(.*?)\\end\{ex\}'
        return re.findall(pattern, latex_content, re.DOTALL)

    def parse_exercise(self, exercise_content):
        """Phân tích nội dung của một bài tập riêng lẻ."""
        # Trích xuất câu hỏi (nội dung trong \immini{...} hoặc trước \choice)
        question_match = re.search(r'\\immini\{(.*?)\}', exercise_content, re.DOTALL)
        if question_match:
            question_content = question_match.group(1).strip()
        else:
            # Nếu không có \immini, lấy tất cả nội dung trước \choice
            question_content = exercise_content.split('\\choice')[0].strip()

        # Trích xuất các lựa chọn và tìm đáp án đúng
        choices_section_match = re.search(r'\\choice(.*?)(?=\\loigiai|\\begin\{tikzpicture\}|$)', exercise_content, re.DOTALL)
        choices = []
        correct_choice_index = -1
        if choices_section_match:
            # Tìm tất cả các cặp dấu {}
            raw_choices = re.findall(r'\{(.*?)\}', choices_section_match.group(1), re.DOTALL)
            for idx, choice in enumerate(raw_choices):
                choice = choice.strip()
                if choice:
                    if choice.startswith('\\True'):
                        correct_choice_index = len(choices)
                        choice = choice.replace('\\True', '').strip()
                    choices.append(choice)
        
        # Trích xuất hình ảnh TikZ
        tikz_match = re.search(r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}', exercise_content, re.DOTALL)
        tikz_content = tikz_match.group(0) if tikz_match else None

        # Trích xuất bảng biểu từ câu hỏi
        tables = self.extract_and_convert_tables(question_content)
        # Loại bỏ nội dung bảng khỏi câu hỏi để tránh lặp lại
        question_text = re.sub(r'\\begin\{tabular\}.*?\\end\{tabular\}', '', question_content, flags=re.DOTALL).strip()

        # Trích xuất lời giải
        solution_match = re.search(r'\\loigiai\{(.*?)\}', exercise_content, re.DOTALL)
        solution = solution_match.group(1).strip() if solution_match else None

        return {
            'question': question_text,
            'choices': choices,
            'correct_choice': correct_choice_index,
            'tikz': tikz_content,
            'tables': tables,
            'solution': solution
        }

    def extract_and_convert_tables(self, content):
        """Trích xuất bảng LaTeX và chuyển đổi sang định dạng markdown."""
        tables = []
        tabular_pattern = r'\\begin\{tabular\}.*?\\end\{tabular\}'
        
        for table_latex in re.findall(tabular_pattern, content, re.DOTALL):
            # Lấy thông số cột
            col_spec_match = re.search(r'\\begin\{tabular\}\{([^}]+)\}', table_latex)
            col_spec = col_spec_match.group(1) if col_spec_match else ''
            col_count = len(re.findall(r'[lcr]', col_spec))
            
            # Lấy nội dung bảng
            table_content_match = re.search(r'\\begin\{tabular\}\{[^}]*\}(.*?)\\end\{tabular\}', table_latex, re.DOTALL)
            table_content = table_content_match.group(1).strip() if table_content_match else ''
            
            markdown_table = self.latex_table_to_markdown(table_content, col_count)
            tables.append(markdown_table)
            
        return tables

    def latex_table_to_markdown(self, table_content, col_count):
        """Chuyển nội dung bảng LaTeX sang markdown."""
        lines = table_content.strip().split('\\\\')
        lines = [line.strip() for line in lines if line.strip() and '\\hline' not in line]
        
        markdown_rows = []
        for line in lines:
            cells = [self.clean_latex_text(cell.strip()) for cell in line.split('&')]
            while len(cells) < col_count:
                cells.append('')
            markdown_rows.append('| ' + ' | '.join(cells) + ' |')
        
        if markdown_rows and col_count > 0:
            separator = '| ' + ' | '.join(['---'] * col_count) + ' |'
            markdown_rows.insert(1, separator)
            
        return '\n'.join(markdown_rows)

    def create_table_from_markdown(self, doc, markdown_table):
        """Tạo bảng Word từ định dạng markdown."""
        lines = [line for line in markdown_table.strip().split('\n') if '---' not in line]
        if not lines:
            return

        rows_data = [[cell.strip() for cell in line.split('|')[1:-1]] for line in lines]
        if not rows_data or not rows_data[0]:
            return
        
        table = doc.add_table(rows=len(rows_data), cols=len(rows_data[0]))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        for i, row_cells in enumerate(rows_data):
            for j, cell_text in enumerate(row_cells):
                cell = table.cell(i, j)
                cell.text = cell_text
                # In đậm hàng đầu tiên (header)
                if i == 0:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True

    def clean_latex_text(self, text):
        """Làm sạch các lệnh LaTeX khỏi văn bản."""
        text = re.sub(r'\s+', ' ', text) # Chuẩn hóa khoảng trắng
        text = re.sub(r'\$(.*?)\$', r'\1', text)  # Loại bỏ dấu $ bao quanh công thức
        text = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\textit\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\begin\{itemize\}', '', text)
        text = re.sub(r'\\end\{itemize\}', '', text)
        text = re.sub(r'\\item', '\n• ', text) # Chuyển \item thành dấu •
        # Thêm các quy tắc khác nếu cần
        return text.strip()

    def compile_tikz_to_image(self, tikz_code, filename_base):
        """Biên dịch mã TikZ thành file ảnh PNG."""
        latex_doc = f"""
\\documentclass[border=5pt]{{standalone}}
\\usepackage{{tikz}}
\\usepackage{{amsmath}}
\\usepackage{{amssymb}}
\\begin{{document}}
{tikz_code}
\\end{{document}}
"""
        tex_file = os.path.join(self.temp_dir, f"{filename_base}.tex")
        with open(tex_file, 'w', encoding='utf-8') as f:
            f.write(latex_doc)
            
        try:
            # Biên dịch bằng pdflatex
            subprocess.run(
                ['pdflatex', '-interaction=nonstopmode', '-output-directory', self.temp_dir, tex_file],
                capture_output=True, check=True, timeout=30
            )
            
            pdf_file = Path(self.temp_dir) / f"{filename_base}.pdf"
            png_file_base = Path(self.temp_dir) / filename_base
            png_file = Path(self.temp_dir) / f"{filename_base}.png"

            # Chuyển đổi PDF sang PNG bằng pdftoppm (ưu tiên)
            subprocess.run(
                ['pdftoppm', '-png', '-r', '300', '-singlefile', str(pdf_file), str(png_file_base)],
                capture_output=True, check=True, timeout=30
            )
            
            if png_file.exists():
                return str(png_file)
            else:
                raise FileNotFoundError("Image conversion with pdftoppm failed.")

        except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
            st.error(f"Lỗi khi biên dịch TikZ: {e}")
            st.warning("Hãy chắc chắn rằng bạn đã cài đặt MiKTeX/TexLive (với pdflatex) và poppler (với pdftoppm).")
            return None

    def add_formatted_run(self, paragraph, text, is_correct=False):
        """Thêm một đoạn text vào paragraph với định dạng gạch chân và in đậm nếu đúng."""
        run = paragraph.add_run(text)
        if is_correct:
            run.bold = True
            run.font.underline = True
        return run

    def create_word_document(self, exercises):
        """Tạo file Word từ các bài tập đã được phân tích."""
        doc = Document()
        doc.add_heading('BÀI TẬP TRẮC NGHIỆM', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for idx, ex in enumerate(exercises, 1):
            # Thêm câu hỏi
            p = doc.add_paragraph()
            p.add_run(f'Câu {idx}. ').bold = True
            p.add_run(self.clean_latex_text(ex['question']))

            # Thêm bảng biểu trong câu hỏi
            if ex['tables']:
                for table_md in ex['tables']:
                    self.create_table_from_markdown(doc, table_md)
                    doc.add_paragraph() # Thêm khoảng cách

            # Thêm hình ảnh TikZ
            if ex['tikz']:
                image_file = self.compile_tikz_to_image(ex['tikz'], f'tikz_{idx}')
                if image_file and os.path.exists(image_file):
                    p_img = doc.add_paragraph()
                    p_img.add_run().add_picture(image_file, width=Inches(3))
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Thêm các lựa chọn
            for i, choice in enumerate(ex['choices']):
                is_correct = (ex['correct_choice'] == i)
                p_choice = doc.add_paragraph()
                p_choice.paragraph_format.left_indent = Inches(0.5)
                # Thêm nhãn (A, B, C, D) và nội dung lựa chọn
                self.add_formatted_run(p_choice, f'{chr(65 + i)}. ', is_correct)
                self.add_formatted_run(p_choice, self.clean_latex_text(choice), is_correct)

            # Thêm lời giải
            if ex['solution']:
                doc.add_paragraph() # Khoảng cách
                p_sol_header = doc.add_paragraph()
                p_sol_header.add_run('Lời giải:').bold = True
                
                # Xử lý bảng biểu trong lời giải
                solution_tables = self.extract_and_convert_tables(ex['solution'])
                solution_text = re.sub(r'\\begin\{tabular\}.*?\\end\{tabular\}', '', ex['solution'], flags=re.DOTALL)
                
                doc.add_paragraph(self.clean_latex_text(solution_text))
                
                if solution_tables:
                    for table_md in solution_tables:
                        self.create_table_from_markdown(doc, table_md)
                        doc.add_paragraph()

            doc.add_paragraph() # Thêm khoảng trống lớn giữa các câu hỏi
        return doc

    def cleanup(self):
        """Dọn dẹp các file tạm."""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)

def main():
    st.set_page_config(page_title="LaTeX to Word Converter", page_icon="🔄", layout="wide")
    load_css()
    
    st.title("🔄 LaTeX to Word Converter")
    st.markdown("Một công cụ mạnh mẽ để chuyển đổi các bài tập trắc nghiệm từ định dạng LaTeX sang tài liệu Word.")
    
    col1, col2 = st.columns([1, 1], gap="large")
    
    with col1:
        st.subheader("📥 Nhập liệu LaTeX")
        
        uploaded_file = st.file_uploader("Tải lên file .tex của bạn", type=['tex'])
        
        latex_input = st.text_area(
            "Hoặc dán trực tiếp vào đây:",
            height=500,
            value=r"""\begin{ex}
\immini{Cho bảng số liệu sau về điểm thi của một lớp học:
\begin{tabular}{|c|c|c|c|}
\hline
\textbf{STT} & \textbf{Họ tên} & \textbf{Điểm} & \textbf{Xếp loại} \\
\hline
1 & Nguyễn Văn A & 8.5 & Giỏi \\
\hline
2 & Trần Thị B & 7.0 & Khá \\
\hline
3 & Lê Văn C & 9.0 & Giỏi \\
\hline
\end{tabular}

Dựa vào bảng trên, có bao nhiêu học sinh được xếp loại Giỏi?}
\choice
{1 học sinh}
{\True 2 học sinh}
{3 học sinh}
{Không có ai}
\loigiai{
Dựa vào cột "Xếp loại" trong bảng, ta thấy có 2 học sinh là "Nguyễn Văn A" và "Lê Văn C" đạt loại Giỏi.
Vậy đáp án đúng là B.
}
\end{ex}

\begin{ex}
\immini{Dựa vào hình vẽ bên dưới, hãy chọn khẳng định đúng.}
\begin{tikzpicture}[scale=0.8]
\coordinate (M) at (0,0); \coordinate (N) at (2,0); \coordinate (P) at (4,0);
\draw[thick, ->] (-1,0) -- (5,0);
\foreach \pt/\pos in {M/above, N/above, P/above}
  \fill (\pt) circle (2pt) node[\pos] {$\pt$};
\end{tikzpicture}
\choice
{Điểm $M$ nằm giữa hai điểm $N$ và $P$.}
{\True Điểm $N$ nằm giữa hai điểm $M$ và $P$.}
{Điểm $P$ nằm giữa hai điểm $M$ và $N$.}
{Không có điểm nào nằm giữa.}
\loigiai{Quan sát hình vẽ, ta thấy thứ tự các điểm trên đường thẳng là M, N, P. Do đó, điểm N nằm giữa hai điểm M và P.}
\end{ex}
"""
        )
        
        if uploaded_file:
            latex_input = uploaded_file.read().decode('utf-8')
            st.info("Đã tải nội dung từ file. Bạn có thể chỉnh sửa trong ô bên trên.")

    with col2:
        st.subheader("📤 Chuyển đổi & Tải xuống")
        
        if st.button("🚀 Chuyển đổi sang Word", type="primary", use_container_width=True):
            if latex_input.strip():
                try:
                    with st.spinner("🧙‍♂️ Đang thực hiện phép màu..."):
                        converter = LaTeXToWordConverter()
                        
                        exercises_raw = converter.extract_exercises(latex_input)
                        if not exercises_raw:
                            st.warning("Không tìm thấy bài tập nào với cấu trúc `\\begin{ex}...\\end{ex}`.")
                            return

                        exercises_parsed = [converter.parse_exercise(ex) for ex in exercises_raw]
                        doc = converter.create_word_document(exercises_parsed)
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        converter.cleanup()

                    st.success(f"✅ Chuyển đổi thành công {len(exercises_parsed)} câu hỏi!")
                    
                    st.download_button(
                        label="📥 Tải xuống file Word",
                        data=doc_io,
                        file_name="Bai_tap_da_chuyen_doi.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Đã xảy ra lỗi: {str(e)}")
            else:
                st.warning("⚠️ Vui lòng nhập nội dung LaTeX hoặc tải file lên.")

    st.markdown("---")
    with st.expander("📖 Hướng dẫn sử dụng và Lưu ý quan trọng"):
        st.markdown("""
        ### Cấu trúc LaTeX được hỗ trợ:
        - **Toàn bộ bài tập**: Phải được bao trong môi trường `\\begin{ex}...\\end{ex}`.
        - **Nội dung câu hỏi**: Đặt trong `\\immini{...}`. Bảng biểu và text nên đặt chung trong đây.
        - **Các lựa chọn**: Bắt đầu bằng lệnh `\\choice`, mỗi lựa chọn đặt trong một cặp dấu `{...}`.
        - **Đáp án đúng**: Đánh dấu đáp án đúng bằng cách thêm `\\True` vào đầu lựa chọn đó. Ví dụ: `{\\True Đáp án đúng}`.
        - **Hình vẽ TikZ**: Đặt trong môi trường `\\begin{tikzpicture}...\\end{tikzpicture}`.
        - **Lời giải**: Đặt trong `\\loigiai{...}`.
        
        ### 🚨 Yêu cầu hệ thống (Rất quan trọng):
        Để chức năng chuyển đổi hình ảnh **TikZ** hoạt động, bạn cần cài đặt các phần mềm sau trên máy tính của mình:
        1.  **Một bản phân phối LaTeX**: Chẳng hạn như [**MiKTeX**](https://miktex.org/download) (cho Windows) hoặc [**TexLive**](https://www.tug.org/texlive/) (cho Linux/Mac).
        2.  **Poppler**: Một thư viện để xử lý file PDF.
            -   **Windows**: Tải từ [trang này](https://github.com/oschwartz10612/poppler-windows/releases/), giải nén và **thêm đường dẫn đến thư mục `bin` vào biến môi trường PATH của hệ thống**.
            -   **Mac (dùng Homebrew)**: `brew install poppler`
            -   **Linux (Ubuntu/Debian)**: `sudo apt-get install poppler-utils`

        Nếu không cài đặt các phần mềm trên, ứng dụng vẫn sẽ chuyển đổi được văn bản và bảng biểu, nhưng sẽ báo lỗi khi xử lý hình vẽ TikZ.
        """)

if __name__ == "__main__":
    main()
