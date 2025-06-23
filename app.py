import streamlit as st
import re
import os
import subprocess
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import shutil

# Placeholder pattern for processed elements
TABLE_PLACEHOLDER = "__TABLE_PLACEHOLDER_{}__"

class LaTeXToWordConverter:
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()

    def extract_exercises(self, latex_content):
        """Extract all exercises from LaTeX content"""
        ex_pattern = r'\\begin\{ex\}(.*?)\\end\{ex\}'
        matches = re.finditer(ex_pattern, latex_content, re.DOTALL)
        return [match.group(1).strip() for match in matches if match.group(1).strip()]

    def parse_exercise(self, exercise_content):
        """Parse individual exercise content to extract all its components."""
        # The main text content to be parsed for question and choices
        parse_target = exercise_content
        
        # If \immini exists, it contains the primary question and choices
        immini_match = re.search(r'\\immini\{(.*?)\}', exercise_content, re.DOTALL)
        if immini_match:
            parse_target = immini_match.group(1).strip()

        # The question is everything before \choice within the parse_target
        question_parts = re.split(r'\\choice', parse_target, maxsplit=1)
        question = question_parts[0].strip()

        # Choices are parsed from the full exercise content to be robust
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
        
        # TikZ picture is extracted from the full content
        tikz_match = re.search(r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}', exercise_content, re.DOTALL)
        tikz = tikz_match.group(0) if tikz_match else None
        if tikz:
             question = question.replace(tikz, "") # Clean tikz from question if it was captured

        # Solution is extracted from the full content
        solution_match = re.search(r'\\loigiai\{(.*?)\}', exercise_content, re.DOTALL)
        solution = solution_match.group(1).strip() if solution_match else None

        return {
            'question': question,
            'choices': choices,
            'correct_choice': correct_choice_index,
            'tikz': tikz,
            'solution': solution
        }

    def latex_table_to_word_table(self, doc, table_content):
        """Directly convert LaTeX tabular content to a Word table."""
        # Get column specification and count
        col_spec_match = re.match(r'\{([^}]+)\}', table_content)
        col_count = 0
        if col_spec_match:
            col_count = len(re.findall(r'[lcr]', col_spec_match.group(1)))
        
        # Get table body
        body_match = re.search(r'\}(.*)', table_content, re.DOTALL)
        if not body_match:
            return None
        
        body = body_match.group(1).strip()
        
        lines = body.split('\\\\')
        rows_data = []
        for line in lines:
            line = line.replace('\\hline', '').strip()
            if not line:
                continue
            cells = [self.prepare_latex_for_word(cell.strip()) for cell in line.split('&')]
            if len(cells) > 0:
                while len(cells) < col_count:
                    cells.append('')
                rows_data.append(cells[:col_count])
        
        if not rows_data:
            return None

        # Create table in Word
        table = doc.add_table(rows=len(rows_data), cols=col_count)
        table.style = 'Table Grid'
        
        for i, row_cells in enumerate(rows_data):
            for j, cell_text in enumerate(row_cells):
                if j < len(table.rows[i].cells):
                    table.cell(i, j).text = cell_text
                    if i == 0: # Bold header
                        for p in table.cell(i, j).paragraphs:
                            for run in p.runs:
                                run.bold = True
        return table

    def process_content_and_placeholders(self, content):
        """
        Finds all tables in content, replaces them with placeholders,
        and returns the modified content and a list of table contents.
        """
        tables = []
        
        def replacer(match):
            table_content = match.group(0)
            # The full tabular content including \begin, spec, and \end
            full_table_latex = f"\\begin{{tabular}}{table_content}"
            tables.append(full_table_latex)
            placeholder = TABLE_PLACEHOLDER.format(len(tables) - 1)
            return placeholder

        # Regex to find content from column spec to end of tabular
        pattern = r'(\{.*?\}.*?\\end\{tabular\})'
        # We replace the tabular environment with a placeholder
        content_with_placeholders = re.sub(pattern, replacer, content, flags=re.DOTALL)
        
        return content_with_placeholders, tables
    
    def add_content_to_doc(self, doc, content_with_placeholders, tables):
        """Adds text and tables to the doc according to placeholders."""
        # Split text by placeholders and add content sequentially
        parts = re.split(f'({TABLE_PLACEHOLDER.format("[0-9]+")})', content_with_placeholders)
        
        for part in parts:
            if not part:
                continue
            
            placeholder_match = re.match(f'{TABLE_PLACEHOLDER.format("([0-9]+)")}', part)
            if placeholder_match:
                table_index = int(placeholder_match.group(1))
                if table_index < len(tables):
                    self.latex_table_to_word_table(doc, tables[table_index])
                    doc.add_paragraph() # Add space after table
            else:
                # This is a regular text part
                prepared_text = self.prepare_latex_for_word(part)
                if prepared_text:
                    doc.add_paragraph(prepared_text)

    def prepare_latex_for_word(self, text):
        """Cleans LaTeX text for Word output, preserving math formulas."""
        # Remove environments but keep content
        text = re.sub(r'\\begin\{(center|align|align\*)\}', '', text, flags=re.DOTALL)
        text = re.sub(r'\\end\{(center|align|align\*)\}', '', text, flags=re.DOTALL)
        text = re.sub(r'\\vspace\{.*?\}', '', text) # Remove vspace

        # General cleaning
        text = re.sub(r'\\item', '•', text)
        text = re.sub(r'\\begin\{itemize\}', '', text)
        text = re.sub(r'\\end\{itemize\}', '', text)
        text = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\textit\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\text\{([^}]*)\}', r'\1', text)

        text = text.replace('\\\\', '')
        text = text.replace('\\hline', '')
        
        return re.sub(r'\s+', ' ', text).strip()

    def compile_tikz_to_image(self, tikz_code, filename_base):
        """Compile TikZ code to a PNG image."""
        latex_doc = f"""
\\documentclass[border=5pt]{{standalone}}
\\usepackage{{tikz}}
\\usepackage{{amsmath}}
\\usepackage{{amssymb}}
\\usetikzlibrary{{arrows.meta}}
\\begin{{document}}
{tikz_code}
\\end{{document}}
"""
        tex_file = os.path.join(self.temp_dir, f"{filename_base}.tex")
        with open(tex_file, 'w', encoding='utf-8') as f:
            f.write(latex_doc)
            
        try:
            subprocess.run(
                ['pdflatex', '-interaction=nonstopmode', '-output-directory', self.temp_dir, tex_file],
                capture_output=True, check=True, timeout=30
            )
            pdf_file = Path(self.temp_dir) / f"{filename_base}.pdf"
            png_file = Path(self.temp_dir) / f"{filename_base}.png"
            subprocess.run(
                ['pdftoppm', '-png', '-r', '300', '-singlefile', str(pdf_file), str(pdf_file.with_suffix(''))],
                capture_output=True, check=True, timeout=30
            )
            return str(png_file) if png_file.exists() else None
        except FileNotFoundError as e:
            st.error(f"Lỗi: Lệnh `{e.filename}` không tìm thấy. Hãy chắc chắn rằng bạn đã cài đặt LaTeX (MiKTeX, TeX Live) và Poppler, và đã thêm chúng vào PATH hệ thống.")
            return None
        except subprocess.CalledProcessError as e:
            st.error(f"Lỗi khi biên dịch TikZ. Kiểm tra code TikZ hoặc log lỗi bên dưới.")
            st.code(e.stderr.decode('utf-8', errors='ignore'), language='bash')
            return None
        except Exception as e:
            st.error(f"Lỗi không mong muốn: {e}")
            return None
    
    def create_word_document(self, exercises):
        """Create Word document from parsed exercises using placeholder method."""
        doc = Document()
        doc.add_heading('Bài tập', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for idx, ex in enumerate(exercises, 1):
            para = doc.add_paragraph()
            para.add_run(f'Câu {idx}. ').bold = True
            
            # Process question content (text and tables)
            # First, remove choices and solution from the question text to avoid duplication
            question_text = ex['question']
            
            content_with_placeholders, tables = self.process_content_and_placeholders(question_text)
            self.add_content_to_doc(para, content_with_placeholders, tables)

            # Add TikZ image if it exists
            if ex['tikz']:
                image_file = self.compile_tikz_to_image(ex['tikz'], f'tikz_{idx}')
                if image_file:
                    para = doc.add_paragraph()
                    para.add_run().add_picture(image_file, width=Inches(3))
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add choices
            for i, choice in enumerate(ex['choices']):
                para = doc.add_paragraph(style='List Paragraph')
                label_run = para.add_run(f'{chr(65 + i)}. ')
                text_run = para.add_run(self.prepare_latex_for_word(choice))
                if ex['correct_choice'] == i:
                    label_run.bold = True
                    label_run.underline = True
                    text_run.underline = True
            
            # Add solution if it exists
            if ex['solution']:
                doc.add_paragraph()
                doc.add_paragraph().add_run('Lời giải:').bold = True
                
                sol_content_placeholders, sol_tables = self.process_content_and_placeholders(ex['solution'])
                self.add_content_to_doc(doc, sol_content_placeholders, sol_tables)

            doc.add_paragraph()
        return doc

    def cleanup(self):
        """Clean up temporary files"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)

def main():
    st.set_page_config(page_title="LaTeX to Word Converter", page_icon="📝")
    st.title("🔄 LaTeX to Word Converter (Nâng cấp)")
    st.markdown("Chuyển đổi bài tập LaTeX sang Word, hỗ trợ bảng và cấu trúc lồng nhau.")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📥 Input LaTeX")
        latex_input = st.text_area(
            "Nhập code LaTeX của bạn:",
            height=500,
            value=r"""\begin{ex}
Hai mẫu số liệu ghép nhóm $M_1, M_2$ có bảng tần số ghép nhóm như sau:
\begin{center}
	$M_1 \quad$\begin{tabular}{|c|c|c|c|c|c|}
		\hline Nhóm & {$[8 ; 10)$} & {$[10 ; 12)$} & {$[12 ; 14)$} & {$[14 ; 16)$} & {$[16 ; 18)$} \\
		\hline Tần số & 3 & 4 & 8 & 6 & 4 \\
		\hline
	\end{tabular}
\end{center}\vspace{2mm}
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
\end{ex}

\begin{ex}
\immini{
	Cho hàm số $y=\dfrac{a x+b}{c x+d}(c \neq 0, a d-b c \neq 0)$ có đồ thị như hình vẽ bên. Tiệm cận ngang của đồ thị hàm số là:
	\choice
	{$x=-1$}
	{\True $y=\dfrac{1}{2}$}
	{$y=-1$}
	{$x=\dfrac{1}{2}$}}
{
	\begin{tikzpicture}[scale=1.5,>=stealth, line join=round, line cap=round]
		\tikzset{declare function={xmin=-3.5;xmax=2.5;ymin=-2.5;ymax=3.5;},smooth,samples=450}
		\draw[->] (xmin,0)--(xmax,0) node[below]{$ x $};
		\draw[->] (0,ymin)--(0,ymax) node[right]{$ y $};
        \node[below left] at (0,0) {$O$};
		\draw[dashed, thin](-1,ymin)--(-1,ymax) node[above, xshift=-0.4cm]{$x=-1$};
		\draw[dashed, thin](xmin,0.5)--(xmax,0.5) node[right]{$y=\frac{1}{2}$};
		\clip (xmin+.1,ymin+.1) rectangle (xmax-.1,ymax-.1);
		\draw[blue, thick] plot[domain=xmin:-1.05] (\x, {(0.5*(\x)-1)/((\x)+1)});
		\draw[blue, thick] plot[domain=-0.95:xmax] (\x, {(0.5*(\x)-1)/((\x)+1)});
	\end{tikzpicture}
}
\end{ex}
"""
        )
        
    with col2:
        st.subheader("📤 Output")
        if st.button("🔄 Chuyển đổi sang Word", type="primary"):
            if latex_input:
                converter = LaTeXToWordConverter()
                try:
                    with st.spinner("Đang xử lý... Vui lòng chờ..."):
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
                except Exception as e:
                    st.error(f"❌ Lỗi: {str(e)}")
                finally:
                    converter.cleanup()
            else:
                st.warning("⚠️ Vui lòng nhập nội dung LaTeX")

if __name__ == "__main__":
    main()
