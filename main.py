from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import io
import zipfile
import os
import tempfile
from collections import defaultdict
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from typing import List
import shutil
from pathlib import Path

app = FastAPI(title="Seating Arrangement Generator", version="1.0.0")

# Create uploads directory if it doesn't exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("output", exist_ok=True)

def add_borders_to_word_table(table):
    """Add borders to Word table"""
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            borders = OxmlElement('w:tcBorders')
            for edge in ('top', 'left', 'bottom', 'right'):
                tag = OxmlElement(f'w:{edge}')
                tag.set(qn('w:val'), 'single')
                tag.set(qn('w:sz'), '4')
                tag.set(qn('w:space'), '0')
                tag.set(qn('w:color'), '000000')
                borders.append(tag)
            tcPr.append(borders)

def fill_seats_columnwise(paper_queue, paper_groups, rows, cols):
    """Fill seats column-wise with alternating papers"""
    room = [["" for _ in range(cols)] for _ in range(rows)]
    dept_map = [["" for _ in range(cols)] for _ in range(rows)]
    paper_map = [["" for _ in range(cols)] for _ in range(rows)]
    seat_order = [(r, c) for c in range(cols) for r in range(rows)]
    seat_index = 0
    
    while seat_index < len(seat_order) and paper_queue:
        p1 = paper_queue[0]
        p2 = paper_queue[1] if len(paper_queue) > 1 else None
        
        for i in range(seat_index, len(seat_order)):
            r, c = seat_order[i]
            current_paper = None
            
            if c % 2 == 0 and paper_groups[p1]:
                current_paper = p1
            elif c % 2 == 1 and p2 and paper_groups[p2]:
                current_paper = p2
                
            if current_paper:
                roll, dept = paper_groups[current_paper].pop(0)
                room[r][c] = roll
                dept_map[r][c] = dept
                paper_map[r][c] = current_paper
                seat_index += 1
                
                if not paper_groups[current_paper]:
                    paper_queue.remove(current_paper)
                break
            else:
                seat_index += 1
                
    return room, dept_map, paper_map

def get_column_departments(dept_map, col, rows):
    """Get departments for a specific column"""
    dept_set = set()
    for r in range(rows):
        if dept_map[r][col]:
            dept_set.add(dept_map[r][col].upper())
    return "/".join(sorted(dept_set))

@app.get("/", response_class=HTMLResponse)
async def read_root():
    """Serve the main HTML interface"""
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Seating Arrangement Generator</title>
        <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
            .form-group { margin-bottom: 15px; }
            label { display: block; margin-bottom: 5px; font-weight: bold; }
            input, textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
            button { background-color: #007bff; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
            button:hover { background-color: #0056b3; }
            .example { font-size: 12px; color: #666; margin-top: 2px; }
            .file-input { border: 2px dashed #ddd; padding: 20px; text-align: center; }
        </style>
    </head>
    <body>
        <h1>📋 Seating Arrangement Generator</h1>
        <form id="uploadForm" enctype="multipart/form-data">
            <div class="form-group">
                <label for="excel_file">📊 Upload Student Excel File:</label>
                <input type="file" id="excel_file" name="excel_file" accept=".xlsx,.xls" required>
                <div class="example">Excel should have columns: Name, RollNo, Paper Code</div>
            </div>
            
            <div class="form-group">
                <label for="template_file">📄 Upload Word Template:</label>
                <input type="file" id="template_file" name="template_file" accept=".docx" required>
            </div>
            
            <div class="form-group">
                <label for="mappings">🗺️ Paper Mappings:</label>
                <textarea id="mappings" name="mappings" rows="3" placeholder="ICT202-16407722-16407255-ECE,MAT201-16407100-16407200-CSE" required></textarea>
                <div class="example">Format: PaperCode-Last8Digits-Last8Digits-Department (comma separated)</div>
            </div>
            
            <div class="form-group">
                <label for="room_specs">🏫 Room Specifications:</label>
                <textarea id="room_specs" name="room_specs" rows="2" placeholder="Room1:48:6x8,Room2:40:5x8" required></textarea>
                <div class="example">Format: RoomName:Capacity:RowsxCols (comma separated)</div>
            </div>
            
            <div class="form-group">
                <label for="exam_date">📅 Exam Date:</label>
                <input type="text" id="exam_date" name="exam_date" placeholder="31-05-2025" required>
            </div>
            
            <div class="form-group">
                <label for="exam_time">⏰ Exam Time:</label>
                <input type="text" id="exam_time" name="exam_time" placeholder="10:00 AM – 1:00 PM" required>
            </div>
            
            <button type="submit">🚀 Generate Seating Arrangement</button>
        </form>
        
        <div id="result" style="margin-top: 20px;"></div>
        
        <script>
            document.getElementById('uploadForm').addEventListener('submit', async function(e) {
                e.preventDefault();
                
                const formData = new FormData();
                formData.append('excel_file', document.getElementById('excel_file').files[0]);
                formData.append('template_file', document.getElementById('template_file').files[0]);
                formData.append('mappings', document.getElementById('mappings').value);
                formData.append('room_specs', document.getElementById('room_specs').value);
                formData.append('exam_date', document.getElementById('exam_date').value);
                formData.append('exam_time', document.getElementById('exam_time').value);
                
                document.getElementById('result').innerHTML = '<p>⏳ Processing... Please wait...</p>';
                
                try {
                    const response = await fetch('/generate-seating', {
                        method: 'POST',
                        body: formData
                    });
                    
                    if (response.ok) {
                        const blob = await response.blob();
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.style.display = 'none';
                        a.href = url;
                        a.download = 'Final_Seating_Documents.zip';
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(url);
                        document.getElementById('result').innerHTML = '<p>✅ Files generated successfully! Download started.</p>';
                    } else {
                        const error = await response.text();
                        document.getElementById('result').innerHTML = '<p style="color: red;">❌ Error: ' + error + '</p>';
                    }
                } catch (error) {
                    document.getElementById('result').innerHTML = '<p style="color: red;">❌ Error: ' + error.message + '</p>';
                }
            });
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@app.post("/generate-seating")
async def generate_seating(
    excel_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
    mappings: str = Form(...),
    room_specs: str = Form(...),
    exam_date: str = Form(...),
    exam_time: str = Form(...)
):
    """Generate seating arrangement files"""
    try:
        # Create temporary directory for this request
        temp_dir = tempfile.mkdtemp()
        
        # Read Excel file
        excel_content = await excel_file.read()
        df = pd.read_excel(io.BytesIO(excel_content))
        df.columns = df.columns.str.strip().str.lower()
        df = df[['name', 'rollno', 'paper code']]
        df['rollno'] = df['rollno'].astype(str).str.zfill(11)
        df['paper code'] = df['paper code'].str.strip()
        df['last8'] = df['rollno'].str[-8:]
        
        # Read Word template
        template_content = await template_file.read()
        template_bytes = io.BytesIO(template_content)
        
        # Parse mappings
        input_map = mappings.split(",")
        paper_last8_dept_map = {}
        for entry in input_map:
            parts = entry.strip().split("-")
            if len(parts) < 3:
                continue
            paper = parts[0].strip()
            dept = parts[-1].strip()
            for last8 in parts[1:-1]:
                paper_last8_dept_map[(paper, last8.strip())] = dept
        
        # Filter valid papers and assign departments
        valid_papers = {k[0] for k in paper_last8_dept_map}
        df = df[df['paper code'].isin(valid_papers)]
        df['department'] = df.apply(lambda row: paper_last8_dept_map.get((row['paper code'], row['last8'])), axis=1)
        df = df[df['department'].notna()]
        
        # Parse room specifications
        parsed_rooms = []
        for spec in room_specs.split(","):
            parts = spec.strip().split(":")
            name = parts[0]
            layout = parts[2] if len(parts) == 3 else "6x8"
            rows, cols = map(int, layout.lower().split("x"))
            parsed_rooms.append((name, rows, cols))
        
        # Group students by paper
        paper_groups = defaultdict(list)
        for _, row in df.iterrows():
            paper_groups[row['paper code']].append((row['rollno'], row['department']))
        
        if not paper_groups:
            raise HTTPException(status_code=400, detail="No valid students found for the given mappings")
        
        # Color palette for Excel
        color_palette = ["F8CBAD", "DDEBF7", "C6E0B4", "F4B084", "FFD966", "D9D2E9", "B4C6E7", "E2EFDA"]
        paper_colors = {paper: color_palette[i % len(color_palette)] for i, paper in enumerate(paper_groups)}
        
        # Generate documents
        final_docx_files = []
        wb = Workbook()
        wb.remove(wb.active)
        paper_queue = list(paper_groups.keys())
        
        for room_name, rows, cols in parsed_rooms:
            if not any(paper_groups.values()):
                break
                
            room, dept_map, paper_map = fill_seats_columnwise(paper_queue, paper_groups, rows, cols)
            
            # Generate Word document
            doc = Document(template_bytes)
            
            # Update document content
            found_time = False
            for p in doc.paragraphs:
                if 'DATE:' in p.text:
                    p.text = f'DATE: {exam_date}'
                elif 'TIME' in p.text.upper():
                    p.text = f'TIME: {exam_time}'
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    if p.runs:
                        p.runs[0].bold = True
                    found_time = True
                elif 'ROOM NO.' in p.text.upper():
                    p.clear()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run1 = p.add_run("SEATING ARRANGEMENT FOR ")
                    run2 = p.add_run(f"Room No. {room_name}")
                    run2.bold = True
                    run2.font.size = Pt(14)
            
            if not found_time:
                p = doc.add_paragraph(f'TIME: {exam_time}')
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.runs[0].bold = True
            
            # Add summary
            summary_count = defaultdict(int)
            for r in range(rows):
                for c in range(cols):
                    roll = room[r][c]
                    dept = dept_map[r][c]
                    paper = paper_map[r][c]
                    if roll and dept and paper:
                        summary_count[(dept, paper)] += 1
            
            # Remove existing paper code paragraph
            for para in doc.paragraphs:
                if "PAPER CODE" in para.text:
                    para.clear()
                    break
            
            # Add summary lines
            for (dept, paper), count in summary_count.items():
                line = f"{dept.upper()} (PAPER CODE {paper}) – {{{count}}}"
                para = doc.add_paragraph(line)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.runs[0].bold = True
            
            # Create seating table
            table = doc.add_table(rows=rows + 1, cols=cols)
            try:
                table.style = 'Table Grid'
            except KeyError:
                pass
            
            # Add column headers
            for c in range(cols):
                dept = get_column_departments(dept_map, c, rows)
                col_label = "ROW-1" if c < cols // 2 else "ROW-2"
                table.cell(0, c).text = f"{dept}\n{col_label}"
                para = table.cell(0, c).paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if para.runs:
                    para.runs[0].bold = True
            
            # Fill table with roll numbers
            for r in range(rows):
                for c in range(cols):
                    table.cell(r + 1, c).text = room[r][c]
                    table.cell(r + 1, c).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            add_borders_to_word_table(table)
            
            # Save Word document
            word_file = os.path.join(temp_dir, f"{room_name}_Seating.docx")
            doc.save(word_file)
            final_docx_files.append(word_file)
            
            # Create Excel sheet
            sheet = wb.create_sheet(title=room_name)
            
            # Add column headers to Excel
            for c in range(cols):
                dept = get_column_departments(dept_map, c, rows)
                col_label = "ROW-1" if c < cols // 2 else "ROW-2"
                sheet.cell(row=1, column=c + 2, value=f"{dept}\n{col_label}")
            
            # Fill Excel with data and colors
            for r in range(rows):
                sheet.cell(row=r + 2, column=1, value=f"Row {r+1}")
                for c in range(cols):
                    roll = room[r][c]
                    paper = paper_map[r][c]
                    cell = sheet.cell(row=r + 2, column=c + 2, value=roll)
                    if paper:
                        fill_color = paper_colors.get(paper, "FFFFFF")
                        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        # Save Excel file
        excel_file_path = os.path.join(temp_dir, "Seating_Summary.xlsx")
        wb.save(excel_file_path)
        
        # Create zip file
        zip_path = os.path.join(temp_dir, "Final_Seating_Documents.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for f in final_docx_files:
                zipf.write(f, os.path.basename(f))
            zipf.write(excel_file_path, "Seating_Summary.xlsx")
        
        # Return zip file
        return FileResponse(
            path=zip_path,
            filename="Final_Seating_Documents.zip",
            media_type="application/zip"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)