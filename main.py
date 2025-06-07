# http://127.0.0.1:8000/docs#/default/generate_seating_plan_generate_post

from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import StreamingResponse
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import io

app = FastAPI()

@app.post("/generate")
async def generate_seating_plan(
    excel_file: UploadFile,
    input_map: str = Form(...),
    room_specs: str = Form(...)
):
    df = pd.read_excel(excel_file.file)

    # Clean and prepare data
    df.columns = df.columns.str.strip().str.lower()
    df = df[['name', 'rollno', 'paper code']]
    df['rollno'] = df['rollno'].astype(str).str.zfill(11)
    df['paper code'] = df['paper code'].str.strip()

    # Parse input_map
    paper_last8_dept_map = {}
    for entry in input_map.split(","):
        parts = entry.strip().split("-")
        if len(parts) < 3:
            continue
        paper = parts[0].strip()
        dept = parts[-1].strip()
        roll_last8s = parts[1:-1]
        for last8 in roll_last8s:
            paper_last8_dept_map[(paper, last8.strip())] = dept

    # Filter valid papers
    valid_papers = {k[0] for k in paper_last8_dept_map}
    df = df[df['paper code'].isin(valid_papers)]

    # Match department and format display
    df['last8'] = df['rollno'].str[-8:]
    df['department'] = df.apply(lambda row: paper_last8_dept_map.get((row['paper code'], row['last8']), None), axis=1)
    df = df[df['department'].notna()]
    df['display'] = df['rollno'] + " (" + df['department'] + ")"

    # Parse room specs
    parsed_rooms = []
    for spec in room_specs.split(","):
        parts = spec.strip().split(":")
        room_name = parts[0]
        if len(parts) == 3:
            rows, cols = map(int, parts[2].lower().split("x"))
        else:
            rows, cols = 6, 8
        parsed_rooms.append((room_name, rows, cols))

    # Group by paper
    paper_groups = defaultdict(list)
    for _, row in df.iterrows():
        paper_groups[row['paper code']].append(row['display'])

    # Assign colors
    color_palette = [
        "BDD7EE", "FCE4D6", "E2EFDA", "FFF2CC", "D9E1F2", "F8CBAD",
        "DDEBF7", "C6E0B4", "F4B084", "FFD966", "D9D2E9", "B4C6E7"
    ]
    paper_colors = {paper: color_palette[i % len(color_palette)] for i, paper in enumerate(paper_groups)}

    def fill_seating_columnwise(paper_queue, paper_groups, rows, cols):
        room = [["" for _ in range(cols)] for _ in range(rows)]
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
                    roll = paper_groups[current_paper].pop(0)
                    room[r][c] = roll
                    paper_map[r][c] = current_paper
                    seat_index += 1
                    if not paper_groups[current_paper]:
                        paper_queue.remove(current_paper)
                    break
                else:
                    seat_index += 1
        return room, paper_map

    # Build room layouts
    rooms = []
    paper_queue = list(paper_groups.keys())

    for room_name, rows, cols in parsed_rooms:
        if not any(paper_groups.values()):
            break
        room, paper_map = fill_seating_columnwise(paper_queue, paper_groups, rows, cols)
        rooms.append((room_name, room, paper_map, rows, cols))

    # Create styled Excel output
    wb = Workbook()
    wb.remove(wb.active)

    for room_name, room, paper_map, rows, cols in rooms:
        ws = wb.create_sheet(title=room_name)
        for col_num in range(cols):
            ws.cell(row=1, column=col_num+2, value=f"Col {col_num+1}")
        for row_num in range(rows):
            ws.cell(row=row_num+2, column=1, value=f"Row {row_num+1}")
            for col_num in range(cols):
                val = room[row_num][col_num]
                cell = ws.cell(row=row_num+2, column=col_num+2, value=val)
                paper = paper_map[row_num][col_num]
                if paper:
                    fill_color = paper_colors[paper]
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

    # Save to memory
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=seating_plan.xlsx"}
    )
