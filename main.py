# http://127.0.0.1:8000/docs#/default/generate_seating_plan_generate_post



# from fastapi import FastAPI, UploadFile, Form
# from fastapi.responses import StreamingResponse
# import pandas as pd
# from collections import defaultdict
# from io import BytesIO
# from openpyxl import Workbook
# from openpyxl.styles import PatternFill

# app = FastAPI()

# @app.post("/generate-seating/")
# async def generate_seating_plan(
#     file: UploadFile,
#     mapping: str = Form(...),  # Format: ICT202-16407722-16407255-ECE, ICT204-16401521-CSE
#     room_specs: str = Form(...)  # Format: Room1:72:9x8, Room2
# ):
#     # 📥 Read Excel file
#     content = await file.read()
#     df = pd.read_excel(BytesIO(content))

#     # 🧹 Clean and prepare
#     df.columns = df.columns.str.strip().str.lower()
#     df = df[['name', 'rollno', 'paper code']]
#     df['rollno'] = df['rollno'].astype(str).str.zfill(11)
#     df['paper code'] = df['paper code'].str.strip()

#     # ✍ Parse mapping
#     input_map = mapping.split(",")
#     paper_last8_dept_map = {}
#     for entry in input_map:
#         parts = entry.strip().split("-")
#         if len(parts) < 3:
#             continue
#         paper = parts[0].strip()
#         dept = parts[-1].strip()
#         roll_last8s = parts[1:-1]
#         for last8 in roll_last8s:
#             paper_last8_dept_map[(paper, last8.strip())] = dept

#     valid_papers = {k[0] for k in paper_last8_dept_map}
#     df = df[df['paper code'].isin(valid_papers)]

#     df['last8'] = df['rollno'].str[-8:]
#     df['department'] = df.apply(lambda row: paper_last8_dept_map.get((row['paper code'], row['last8']), None), axis=1)
#     df = df[df['department'].notna()]
#     df['display'] = df['rollno'] + " (" + df['department'] + ")"

#     # 🏫 Parse room specs
#     parsed_rooms = []
#     for spec in room_specs.split(","):
#         parts = spec.strip().split(":")
#         room_name = parts[0]
#         if len(parts) == 3:
#             rows, cols = map(int, parts[2].lower().split("x"))
#         else:
#             rows, cols = 6, 8
#         parsed_rooms.append((room_name, rows, cols))

#     # 📦 Group by paper
#     paper_groups = defaultdict(list)
#     for _, row in df.iterrows():
#         paper_groups[row['paper code']].append(row['display'])

#     # 🎨 Colors
#     color_palette = [
#         "BDD7EE", "FCE4D6", "E2EFDA", "FFF2CC", "D9E1F2", "F8CBAD",
#         "DDEBF7", "C6E0B4", "F4B084", "FFD966", "D9D2E9", "B4C6E7"
#     ]
#     paper_colors = {paper: color_palette[i % len(color_palette)] for i, paper in enumerate(paper_groups)}

#     # 🧠 Column-wise seating
#     def fill_seating_columnwise(paper_queue, paper_groups, rows, cols):
#         room = [["" for _ in range(cols)] for _ in range(rows)]
#         paper_map = [["" for _ in range(cols)] for _ in range(rows)]
#         seat_order = [(r, c) for c in range(cols) for r in range(rows)]
#         seat_index = 0

#         while seat_index < len(seat_order) and paper_queue:
#             p1 = paper_queue[0]
#             p2 = paper_queue[1] if len(paper_queue) > 1 else None

#             for i in range(seat_index, len(seat_order)):
#                 r, c = seat_order[i]
#                 current_paper = None
#                 if c % 2 == 0 and paper_groups[p1]:
#                     current_paper = p1
#                 elif c % 2 == 1 and p2 and paper_groups[p2]:
#                     current_paper = p2

#                 if current_paper:
#                     roll = paper_groups[current_paper].pop(0)
#                     room[r][c] = roll
#                     paper_map[r][c] = current_paper
#                     seat_index += 1
#                     if not paper_groups[current_paper]:
#                         paper_queue.remove(current_paper)
#                     break
#                 else:
#                     seat_index += 1
#         return room, paper_map

#     # 🏫 Build room layouts
#     rooms = []
#     paper_queue = list(paper_groups.keys())

#     for room_name, rows, cols in parsed_rooms:
#         if not any(paper_groups.values()):
#             break
#         room, paper_map = fill_seating_columnwise(paper_queue, paper_groups, rows, cols)
#         rooms.append((room_name, room, paper_map, rows, cols))

#     # 💾 Excel output
#     wb = Workbook()
#     wb.remove(wb.active)

#     for room_name, room, paper_map, rows, cols in rooms:
#         ws = wb.create_sheet(title=room_name)

#         # Headers
#         for col_num in range(cols):
#             ws.cell(row=1, column=col_num+2, value=f"Col {col_num+1}")
#         for row_num in range(rows):
#             ws.cell(row=row_num+2, column=1, value=f"Row {row_num+1}")
#             for col_num in range(cols):
#                 val = room[row_num][col_num]
#                 cell = ws.cell(row=row_num+2, column=col_num+2, value=val)
#                 paper = paper_map[row_num][col_num]
#                 if paper:
#                     fill_color = paper_colors[paper]
#                     cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

#         # 📊 Inline summary
#         paper_counter = defaultdict(int)
#         for r in range(rows):
#             for c in range(cols):
#                 paper = paper_map[r][c]
#                 if paper:
#                     paper_counter[paper] += 1
#         start_row = rows + 4
#         ws.cell(row=start_row - 1, column=1, value="📋 Paper-wise Student Count")
#         for i, (paper, count) in enumerate(paper_counter.items()):
#             ws.cell(row=start_row + i, column=1, value=f"{paper}: {count}")
#         ws.cell(row=start_row + len(paper_counter) + 1, column=1,
#                 value=f"🧮 Total students: {sum(paper_counter.values())}")

#     # 🔁 StreamingResponse instead of saving
#     output = BytesIO()
#     wb.save(output)
#     output.seek(0)

#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": "attachment; filename=seating_plan.xlsx"}
#     )





# # End 


# 📦 Install all necessary dependencies
# !pip install fastapi nest-asyncio pyngrok uvicorn python-docx openpyxl --quiet

# 🧠 Runtime setup
import nest_asyncio
from pyngrok import ngrok
import uvicorn

nest_asyncio.apply()

# 🔐 STEP 1: Set your Ngrok Auth Token (get it from https://dashboard.ngrok.com/get-started/your-authtoken)
ngrok.set_auth_token("2yAxXlXEyCxUXbQWe3O1nBydZTp_TDQzgdjZLwopBAa6qiFZ")  # <-- Replace this with your actual token!

# 🚀 FastAPI app definition
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import StreamingResponse
import pandas as pd
import io
import zipfile
from collections import defaultdict
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import PatternFill

app = FastAPI()

def set_table_borders(table):
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

def get_column_departments(dept_map, col, rows):
    return "/".join(sorted({dept_map[r][col].upper() for r in range(rows) if dept_map[r][col]}))

def fill_columnwise(paper_queue, paper_groups, rows, cols):
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

@app.post("/generate-seating/")
async def generate_seating_plan(
    student_excel: UploadFile = File(...),
    docx_template: UploadFile = File(...),
    mappings: str = Form(...),
    rooms: str = Form(...),
    date: str = Form(...),
    time: str = Form(...),
):
    df = pd.read_excel(io.BytesIO(await student_excel.read()))
    df.columns = df.columns.str.strip().str.lower()
    df = df[['name', 'rollno', 'paper code']]
    df['rollno'] = df['rollno'].astype(str).str.zfill(11)
    df['paper code'] = df['paper code'].str.strip()
    df['last8'] = df['rollno'].str[-8:]

    docx_template_bytes = await docx_template.read()

    mapping_entries = [m.strip() for m in mappings.split(",") if m.strip()]
    room_specs = [r.strip() for r in rooms.split(",") if r.strip()]

    paper_last8_dept_map = {}
    for entry in mapping_entries:
        parts = entry.split("-")
        paper = parts[0].strip()
        dept = parts[-1].strip()
        for last8 in parts[1:-1]:
            paper_last8_dept_map[(paper, last8.strip())] = dept

    df = df[df['paper code'].isin({k[0] for k in paper_last8_dept_map})]
    df['department'] = df.apply(lambda r: paper_last8_dept_map.get((r['paper code'], r['last8'])), axis=1)
    df = df[df['department'].notna()]

    parsed_rooms = []
    for spec in room_specs:
        parts = spec.split(":")
        name = parts[0]
        layout = parts[2] if len(parts) == 3 else "6x8"
        rows, cols = map(int, layout.lower().split("x"))
        parsed_rooms.append((name, rows, cols))

    paper_groups = defaultdict(list)
    for _, row in df.iterrows():
        paper_groups[row['paper code']].append((row['rollno'], row['department']))
    paper_colors = {p: c for p, c in zip(paper_groups.keys(), ["F8CBAD", "DDEBF7", "C6E0B4", "F4B084"]*5)}

    final_docx_files = []
    wb = Workbook()
    wb.remove(wb.active)
    paper_queue = list(paper_groups.keys())

    for room_name, rows, cols in parsed_rooms:
        if not any(paper_groups.values()): break
        room, dept_map, paper_map = fill_columnwise(paper_queue, paper_groups, rows, cols)
        doc = Document(io.BytesIO(docx_template_bytes))

        for p in doc.paragraphs:
            if 'DATE:' in p.text:
                p.text = f'DATE: {date}'
            elif 'TIME' in p.text.upper():
                p.text = f'TIME: {time}'
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if p.runs: p.runs[0].bold = True
            elif 'ROOM NO.' in p.text.upper():
                p.clear(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run("Seating arrangment for")
                p.add_run(f"Room No. {room_name}").bold = True

        summary_count = defaultdict(int)
        for r in range(rows):
            for c in range(cols):
                roll, dept, paper = room[r][c], dept_map[r][c], paper_map[r][c]
                if roll and dept and paper:
                    summary_count[(dept, paper)] += 1

        for (dept, paper), count in summary_count.items():
            para = doc.add_paragraph(f"{dept.upper()} (PAPER CODE {paper}) – {{{count}}}")
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.runs[0].bold = True

        table = doc.add_table(rows=rows + 1, cols=cols)
        for c in range(cols):
            dept = get_column_departments(dept_map, c, rows)
            table.cell(0, c).text = f"{dept}\n{'ROW-1' if c < cols//2 else 'ROW-2'}"
            table.cell(0, c).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        for r in range(rows):
            for c in range(cols):
                table.cell(r + 1, c).text = room[r][c]
                table.cell(r + 1, c).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        set_table_borders(table)
        buffer = io.BytesIO()
        doc.save(buffer)
        final_docx_files.append((f"{room_name}_Seating.docx", buffer.getvalue()))

        sheet = wb.create_sheet(title=room_name)
        for c in range(cols):
            dept = get_column_departments(dept_map, c, rows)
            sheet.cell(row=1, column=c+2, value=f"{dept}\n{'ROW-1' if c < cols//2 else 'ROW-2'}")
        for r in range(rows):
            sheet.cell(row=r+2, column=1, value=f"Row {r+1}")
            for c in range(cols):
                cell = sheet.cell(row=r+2, column=c+2, value=room[r][c])
                paper = paper_map[r][c]
                if paper:
                    cell.fill = PatternFill(start_color=paper_colors.get(paper, "FFFFFF"), end_color=paper_colors.get(paper, "FFFFFF"), fill_type="solid")

    excel_stream = io.BytesIO(); wb.save(excel_stream)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as z:
        for name, content in final_docx_files:
            z.writestr(name, content)
        z.writestr("Seating_Summary.xlsx", excel_stream.getvalue())

    return StreamingResponse(io.BytesIO(zip_buffer.getvalue()), media_type="application/x-zip-compressed",
                             headers={"Content-Disposition": "attachment; filename=seating_documents.zip"})

# 🌐 Launch the FastAPI app and expose it via ngrok
public_url = ngrok.connect(8000)
print("Your app is live at:", public_url)
uvicorn.run(app, host="0.0.0.0", port=8000)