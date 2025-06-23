import streamlit as st
import re
import os
import subprocess
import tempfile
from pathlib import Path
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import shutil

class LaTeXToWordConverter:
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()

    def extract_exercises(self, latex_content):
        """Extract all exercises from LaTeX content"""
        # Pattern to match \begin{ex} ... \end{ex}
        ex_pattern = r'\\begin\{ex\}(.*?)\\end\{ex\}'
        matches = re.finditer(ex_pattern, latex_content, re.DOTALL)
        
        exercises = [match.group(1).strip() for match in matches if match.group(1).strip()]
        return exercises

    def parse_exercise(self, exercise_content):
        """Parse individual exercise content"""
        # Extract question (content before \choice or inside \immini)
        question_match = re.search(r'\\immini\{(.*?)\}', exercise_content, re.DOTALL)
        if question_match:
            question_content = question_match.group(1).strip()
        else:
            # If \immini is not found, take everything before \choice
            parts = re.split(r'\\choice', exercise_content, maxsplit=1)
            question_content = parts[0].strip()

        # Extract choices and find correct answer
        choices = []
        correct_choice_index = -1
        
        # Find the choice section more accurately
        choice_match = re.search(r'\\choice\s*(.*?)(?=\\begin\{tikzpicture\}|\\begin\{tabular\}|\\loigiai|\\end\{ex\}|$)', 
                                 exercise_content, re.DOTALL)
        
        if choice_match:
            choices_text = choice_match.group(1)
            
            # More precise pattern to extract individual choices, handling nested braces
            choice_pattern = r'\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}'
            raw_choices = re.findall(choice_pattern, choices_text)
            
            # Process each choice
            for choice in raw_choices:
                choice = choice.strip()
                if choice:
                    # Check if this is the correct answer
                    if choice.startswith('\\True'):
                        correct_choice_index = len(choices)  # Current index before appending
                        # Remove \True marker
                        choice = re.sub(r'^\\True\s*', '', choice).strip()
                    choices.append(choice)
        
        # Extract TikZ picture
        tikz_match = re.search(r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}', exercise_content, re.DOTALL)
        tikz_content = tikz_match.group(0) if tikz_match else None
        
        # Extract tables from the entire exercise content
        tables = self.extract_and_convert_tables(exercise_content)
        
        # Extract solution
        solution_match = re.search(r'\\loigiai\{(.*?)\}', exercise_content, re.DOTALL)
        solution = solution_match.group(1).strip() if solution_match else None
        
        return {
            'question': question_content,
            'choices': choices,
            'correct_choice': correct_choice_index,
            'tikz': tikz_content,
            'tables': tables,
            'solution': solution
        }

    def extract_and_convert_tables(self, content):
        """Extract LaTeX tables and convert to markdown format"""
        tables = []
        tabular_pattern = r'\\begin\{tabular\}(\{[^}]*\})(.*?)\\end\{tabular\}'
        
        for match in re.finditer(tabular_pattern, content, re.DOTALL):
            column_spec = match.group(1)
            table_content = match.group(2)
            
            # Parse column specification
            col_count = len(re.findall(r'[lcr]', column_spec))
            
            # Convert table to markdown
            markdown_table = self.latex_table_to_markdown(table_content, col_count)
            tables.append(markdown_table)
            
        return tables

    def latex_table_to_markdown(self, table_content, col_count):
        """Convert LaTeX table content to markdown table, preserving math formulas"""
        lines = table_content.strip().split('\\\\')
        lines = [line.strip() for line in lines if line.strip() and not line.strip().startswith('\\hline')]
        
        markdown_rows = []
        for line in lines:
            line = line.replace('\\hline', '').strip()
            if not line:
                continue
                
            cells = line.split('&')
            # Use the new preparation function that preserves LaTeX math
            cells = [self.prepare_latex_for_word(cell.strip()) for cell in cells]
            
            # Pad or trim row to match column count
            while len(cells) < col_count:
                cells.append('')
            cells = cells[:col_count]
            
            markdown_row = '| ' + ' | '.join(cells) + ' |'
            markdown_rows.append(markdown_row)
        
        # Add markdown header separator
        if markdown_rows:
            separator = '| ' + ' | '.join(['---'] * col_count) + ' |'
            markdown_rows.insert(1, separator)
        
        return '\n'.join(markdown_rows)

    def create_table_from_markdown(self, doc, markdown_table):
        """Create a Word table from markdown format"""
        lines = [line for line in markdown_table.strip().split('\n') if line]
        if not lines:
            return

        # Filter out separator line
        data_lines = [line for line in lines if '---' not in line]
        if len(data_lines) == 0:
            return

        rows_data = []
        for line in data_lines:
            cells = [cell.strip() for cell in line.split('|')[1:-1]]
            rows_data.append(cells)
        
        if not rows_data:
            return

        num_rows = len(rows_data)
        num_cols = len(rows_data[0]) if rows_data else 0

        table = doc.add_table(rows=num_rows, cols=num_cols)
        table.style = 'Table Grid'
        
        for i, row_cells in enumerate(rows_data):
            for j, cell_text in enumerate(row_cells):
                if j < len(table.rows[i].cells):
                    table.cell(i, j).text = cell_text
                    # Bold the header row (first row)
                    if i == 0:
                        for para in table.cell(i, j).paragraphs:
                            for run in para.runs:
                                run.bold = True
    
    def prepare_latex_for_word(self, text):
        """
        Cleans LaTeX text for Word output, with a key change:
        - KEEPS math environments ($...$, $$...$$) intact.
        - Removes specified environments like center, align.
        """
        # Remove specific environments but keep their content
        text = re.sub(r'\\begin\{center\}', '', text, flags=re.DOTALL)
        text = re.sub(r'\\end\{center\}', '', text, flags=re.DOTALL)
        text = re.sub(r'\\begin\{align\*?\}', '', text, flags=re.DOTALL)
        text = re.sub(r'\\end\{align\*?\}', '', text, flags=re.DOTALL)
        
        # General text cleaning
        text = re.sub(r'\\item', '•', text)
        text = re.sub(r'\\begin\{itemize\}', '', text)
        text = re.sub(r'\\end\{itemize\}', '', text)
        
        # Handle formatting commands but keep content
        text = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\textit\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)

        # Remove commands that are typically just for spacing or line breaks in LaTeX
        text = text.replace('\\\\', '') # Remove double backslash
        text = text.replace('\\hline', '') # Remove hline
        
        # Normalize whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text

    def compile_tikz_to_image(self, tikz_code, filename_base):
        """Compile TikZ code to a PNG image."""
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
            # Compile with pdflatex
            subprocess.run(
                ['pdflatex', '-interaction=nonstopmode', '-output-directory', self.temp_dir, tex_file],
                capture_output=True, check=True, timeout=30
            )
            
            # Convert PDF to PNG
            pdf_file = Path(self.temp_dir) / f"{filename_base}.pdf"
            png_file = Path(self.temp_dir) / f"{filename_base}.png"
            
            # Use pdftoppm for conversion
            subprocess.run(
                ['pdftoppm', '-png', '-r', '300', '-singlefile', str(pdf_file), str(pdf_file.with_suffix(''))],
                capture_output=True, check=True, timeout=30
            )
            
            if png_file.exists():
                return str(png_file)
            else:
                return None
                
        except FileNotFoundError as e:
            st.error(f"Lỗi: Lệnh `{e.filename}` không tìm thấy. Hãy chắc chắn rằng bạn đã cài đặt một bản phân phối LaTeX (như MiKTeX, TeX Live) và Poppler, và chúng đã được thêm vào PATH hệ thống.")
            return None
        except subprocess.CalledProcessError as e:
            st.error(f"Lỗi khi biên dịch TikZ. Hãy kiểm tra lại code TikZ của bạn.")
            st.code(e.stderr.decode('utf-8', errors='ignore'), language='bash')
            return None
        except Exception as e:
            st.error(f"Đã xảy ra lỗi không mong muốn: {e}")
            return None

    def add_underline_to_run(self, run):
        """Add underline formatting to a run"""
        run.font.underline = True

    def create_word_document(self, exercises):
        """Create Word document from parsed exercises"""
        doc = Document()
        doc.add_heading('Bài tập', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for idx, exercise in enumerate(exercises, 1):
            # Clean the question text while preserving formulas
            # First remove table source code from question to avoid duplication
            question_text = re.sub(r'\\begin\{tabular\}.*?\\end\{tabular\}', '', exercise['question'], flags=re.DOTALL)
            question_text_prepared = self.prepare_latex_for_word(question_text)

            # Add question
            question_para = doc.add_paragraph()
            question_para.add_run(f'Câu {idx}. ').bold = True
            question_para.add_run(question_text_prepared)

            # Add tables found in the question, if any
            if exercise.get('tables'):
                for table_markdown in exercise['tables']:
                    # Remove the table if it's also inside the solution to avoid double printing
                    if not (exercise.get('solution') and table_markdown in self.extract_and_convert_tables(exercise['solution'])):
                       self.create_table_from_markdown(doc, table_markdown)
                       doc.add_paragraph()

            # Add TikZ image if exists
            if exercise['tikz']:
                image_file = self.compile_tikz_to_image(exercise['tikz'], f'tikz_{idx}')
                if image_file and os.path.exists(image_file):
                    doc.add_picture(image_file, width=Inches(3))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add choices
            for i, choice in enumerate(exercise['choices']):
                choice_para = doc.add_paragraph(style='List Paragraph')
                choice_label = f'{chr(65 + i)}. '
                
                is_correct = (exercise.get('correct_choice', -1) == i)
                
                label_run = choice_para.add_run(choice_label)
                text_run = choice_para.add_run(self.prepare_latex_for_word(choice))
                
                if is_correct:
                    label_run.bold = True
                    self.add_underline_to_run(label_run)
                    self.add_underline_to_run(text_run)
            
            # Add solution if exists
            if exercise['solution']:
                doc.add_paragraph()
                solution_header = doc.add_paragraph()
                solution_header.add_run('Lời giải:').bold = True
                
                # Separate text from tables in the solution
                solution_tables_md = self.extract_and_convert_tables(exercise['solution'])
                solution_text_only = re.sub(r'\\begin\{tabular\}.*?\\end\{tabular\}', '', exercise['solution'], flags=re.DOTALL)
                
                if solution_text_only.strip():
                    doc.add_paragraph(self.prepare_latex_for_word(solution_text_only))

                # Add tables from the solution
                for table_md in solution_tables_md:
                    self.create_table_from_markdown(doc, table_md)
                    doc.add_paragraph()

            doc.add_paragraph()

        return doc

    def cleanup(self):
        """Clean up temporary files"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)

def main():
    st.set_page_config(page_title="LaTeX to Word Converter", page_icon="📝")
    
    st.title("🔄 LaTeX to Word Converter")
    st.markdown("Chuyển đổi bài tập LaTeX (giữ nguyên công thức toán) sang định dạng Word.")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📥 Input LaTeX")
        
        latex_input = st.text_area(
            "Nhập code LaTeX của bạn:",
            height=400,
            value=r"""\begin{ex}
\immini{Cho phương trình $x^2 - 2(m-1)x + m^2 - 3 = 0$. Tìm $m$ để phương trình có hai nghiệm phân biệt.
\begin{center}
$x_1, x_2$ là hai nghiệm.
\end{center}
\choice
{$m < 2$}
{\True $m < 2$}
{$m > -2$}
{$m = 2$}
}
\loigiai{
Để phương trình có hai nghiệm phân biệt, ta cần $\Delta' > 0$.
\begin{align*}
\Delta' &= (m-1)^2 - (m^2 - 3) \\
&= m^2 - 2m + 1 - m^2 + 3 \\
&= -2m + 4
\end{align*}
$\Delta' > 0 \Leftrightarrow -2m + 4 > 0 \Leftrightarrow 2m < 4 \Leftrightarrow m < 2$.

Vậy với $m < 2$ thì phương trình có hai nghiệm phân biệt.
}
\end{ex}

\begin{ex}
\immini{Dựa vào hình vẽ (Hình b), hãy chọn khẳng định đúng?
\choice
{ Điểm $M$ nằm giữa $2$ điểm $N$ và $P$}
{\True Điểm $N$ nằm giữa $2$ điểm $M$ và $P$}
{ Điểm $P$ nằm giữa $2$ điểm $M$ và $N$}
{ Hai điểm $M$ và $P$ nằm cùng phía đối với điểm $N$}
}{\begin{tikzpicture}[scale=1]
\coordinate (M) at (0.5, 0);
\coordinate (N) at (2.5, 0);
\coordinate (P) at (4.5, 0);
\draw[thick] (0, 0) -- (5.5, 0);
\foreach \pt/\angle in {M/90, N/90, P/90} {
\draw[fill=white] (\pt) circle (1.5pt) +(\angle:3mm) node{$\pt$};
}
\node[below=5mm of N] {Hình $b$};
\end{tikzpicture}
}
\loigiai{
Theo hình vẽ, các điểm $M, N, P$ thẳng hàng và $N$ nằm giữa $M$ và $P$.
Đáp án đúng là B.
}
\end{ex}"""
        )
        
        uploaded_file = st.file_uploader("Hoặc tải lên file .tex", type=['tex'])
        if uploaded_file:
            latex_input = uploaded_file.read().decode('utf-8')

    with col2:
        st.subheader("📤 Output")
        
        if st.button("🔄 Chuyển đổi sang Word", type="primary"):
            if latex_input:
                converter = LaTeXToWordConverter()
                try:
                    with st.spinner("Đang xử lý..."):
                        exercises_raw = converter.extract_exercises(latex_input)
                        exercises_parsed = [converter.parse_exercise(ex) for ex in exercises_raw]
                        
                        doc = converter.create_word_document(exercises_parsed)
                        
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        st.success("✅ Chuyển đổi thành công!")
                        st.download_button(
                            label="📥 Tải xuống file Word",
                            data=doc_io.getvalue(),
                            file_name="exercises_converted.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                        st.info(f"✅ Đã chuyển đổi {len(exercises_parsed)} câu hỏi.")
                        with st.expander("📊 Chi tiết chuyển đổi"):
                            for i, ex in enumerate(exercises_parsed, 1):
                                st.write(f"**Câu {i}:**")
                                st.write(f"- Số lựa chọn: {len(ex['choices'])}")
                                if ex['correct_choice'] >= 0:
                                    st.write(f"- Đáp án đúng: {chr(65 + ex['correct_choice'])}")
                                if ex['tikz']:
                                    st.write("- Có hình TikZ ✓")
                                if ex['tables']:
                                    st.write(f"- Có {len(ex['tables'])} bảng ✓")
                                if ex['solution']:
                                    st.write("- Có lời giải ✓")
                                st.write("---")
                                
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")
                finally:
                    # Always clean up temp files
                    converter.cleanup()
            else:
                st.warning("⚠️ Vui lòng nhập nội dung LaTeX")

    with st.expander("📖 Hướng dẫn sử dụng và Yêu cầu"):
        st.markdown("""
        ### Yêu cầu hệ thống (QUAN TRỌNG)
        Để có thể chuyển đổi hình vẽ TikZ, máy tính của bạn **bắt buộc** phải cài đặt:
        1.  **Một bản phân phối LaTeX**: Ví dụ như [**MiKTeX**](https://miktex.org/download) (cho Windows), **MacTeX** (cho macOS), hoặc **TeX Live** (cho Linux).
        2.  **Poppler**: Cung cấp công cụ để chuyển PDF sang ảnh. Bạn có thể tải [tại đây](https://github.com/oschwartz10612/poppler-windows/releases/).
        
        **Lưu ý**: Sau khi cài đặt, hãy đảm bảo các thư mục chứa `pdflatex.exe` và `pdftoppm.exe` đã được thêm vào biến môi trường `PATH` của hệ thống.
        
        ### Cấu trúc LaTeX được hỗ trợ
        1.  **Câu hỏi**: Đặt trong `\\begin{ex}...\\end{ex}`
        2.  **Nội dung câu hỏi**: Trong `\\immini{...}` hoặc trước `\\choice`
        3.  **Các lựa chọn**: Sau `\\choice`, mỗi lựa chọn trong `{...}`
        4.  **Đáp án đúng**: Đánh dấu bằng `{\\True ...}` - sẽ được **in đậm và gạch chân** trong Word
        5.  **Công thức toán**: Sẽ được **giữ nguyên** (ví dụ: `$x^2+y^2=z^2$`). Bạn có thể dùng add-in MathType của Word để chuyển đổi chúng sau.
        6.  **Hình vẽ TikZ**: Trong `\\begin{tikzpicture}...\\end{tikzpicture}`
        7.  **Bảng**: Trong `\\begin{tabular}{...}...\\end{tabular}` sẽ được chuyển thành bảng Word.
        8.  **Lời giải**: Trong `\\loigiai{...}`
        9.  **Các môi trường bị loại bỏ**: `\\begin{center}`, `\\begin{align}`, `\\begin{align*}` sẽ bị xóa nhưng nội dung bên trong được giữ lại.
        """)

    st.markdown("---")
    st.markdown("💡 **Tip**: Bạn có thể dán nhiều câu hỏi cùng lúc, mỗi câu trong một môi trường `\\begin{ex}...\\end{ex}`.")

if __name__ == "__main__":
    main()
