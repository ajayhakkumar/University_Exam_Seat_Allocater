import pandas as pd
from collections import defaultdict
import os
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image



class ExamSeatAllocator:
    def __init__(self, excel_file):
        """
        Initialize with path to a single Excel file containing 'Students' and 'Halls' sheets.
        """
        self.students = self._load_students(excel_file)
        self.halls = self._load_halls(excel_file)
        self.allocations = {}
        self.hall_seats = {
            hall['hall_code']: [[None for _ in range(hall['cols'])] for _ in range(hall['rows'])]
            for hall in self.halls
        }
        self.hall_courses = {hall['hall_code']: set() for hall in self.halls}

    def _load_halls(self, excel_file):
        """Load hall details from 'Halls' sheet in Excel file, skipping incomplete rows."""
        try:
            df = pd.read_excel(excel_file, sheet_name='Halls')
            required_columns = ['Hall Code', 'Hall Name', 'Block', 'Row', 'Column', 'Total Capacity']
            if not all(col in df.columns for col in required_columns):
                raise ValueError(f"Halls sheet must contain columns: {', '.join(required_columns)}")

            halls = []
            skipped = 0
            for _, row in df.iterrows():
                if any(pd.isna(row[col]) for col in required_columns):
                    skipped += 1
                    continue
                try:
                    halls.append({
                        'hall_code': str(row['Hall Code']).strip(),
                        'hall_name': str(row['Hall Name']).strip(),
                        'internal_external': str(row.get('Internal/External', '')).strip(),
                        'block': str(row['Block']).strip(),
                        'rows': int(row['Row']),
                        'cols': int(row['Column']),
                        'total_capacity': int(row['Total Capacity'])
                    })
                except Exception as e:
                    skipped += 1
                    continue
            if skipped:
                print(f"Skipped {skipped} incomplete hall rows.")
            return halls
        except Exception as e:
            raise ValueError(f"Error reading halls sheet: {e}")

    def _load_students(self, excel_file):
        """Load student details from 'Students' sheet in Excel file, skipping incomplete rows."""
        try:
            df = pd.read_excel(excel_file, sheet_name='Students')
            required_columns = ['Student Reg.No.', 'Department', 'Course Code', 'Course Title']
            if not all(col in df.columns for col in required_columns):
                raise ValueError(f"Students sheet must contain columns: {', '.join(required_columns)}")

            students = []
            skipped = 0
            for _, row in df.iterrows():
                if any(pd.isna(row[col]) for col in required_columns):
                    skipped += 1
                    continue
                try:
                    students.append({
                        'reg_no': str(row['Student Reg.No.']).strip(),
                        'department': str(row['Department']).strip(),
                        'course_code': str(row['Course Code']).strip(),
                        'course_title': str(row['Course Title']).strip(),
                    })
                except Exception as e:
                    skipped += 1
                    continue
            if skipped:
                print(f"Skipped {skipped} incomplete student rows.")
            return students
        except Exception as e:
            raise ValueError(f"Error reading students sheet: {e}")

    def validate_inputs(self):
        """Validate hall capacities and student count."""
        total_seats = sum(hall['rows'] * hall['cols'] for hall in self.halls)
        for hall in self.halls:
            expected_capacity = hall['rows'] * hall['cols']
            if hall['total_capacity'] > expected_capacity:
                raise ValueError(
                    f"Hall {hall['hall_code']} total_capacity {hall['total_capacity']} "
                    f"cannot exceed grid capacity {expected_capacity}"
                )
        if len(self.students) > total_seats:
            raise ValueError(f"Not enough seats ({total_seats}) for {len(self.students)} students.")

    def allocate_seats(self):
        """Rewrite: Clean proportional distribution + smart round-robin allocation."""

        self.validate_inputs()

        # Group students by course
        students_by_course = defaultdict(list)
        for s in sorted(self.students, key=lambda x: (x['course_code'], x['reg_no'])):
            students_by_course[s['course_code']].append(s)

        # Sort courses by size (largest first)
        courses_sorted = sorted(
            students_by_course.keys(),
            key=lambda c: len(students_by_course[c]),
            reverse=True
        )

        # Prepare remaining courses (mutable)
        remaining = {c: list(students_by_course[c]) for c in courses_sorted}

        # NOTE: per-course alternation (split-then-alternate) removed.
        # Allocation will rely on the round-robin logic in the per-hall
        # allocation loop to mix students across courses. The `remaining`
        # dict therefore keeps each course's students in their sorted order
        # (ascending by reg_no) as prepared earlier.

        # Compute hall selection in 25-seat blocks and proportional students per selected halls
        total_capacity = sum(h['rows'] * h['cols'] for h in self.halls)
        total_students = len(self.students)

        # Number of 25-seat blocks required (round up)
        from math import ceil
        required_blocks = ceil(total_students / 25) if total_students > 0 else 0
        required_seats = required_blocks * 25

        # Select the minimal number of halls (in the order provided) such that their
        # combined grid capacity >= required_seats. If not enough capacity overall,
        # select all halls (allocation will detect shortage later).
        selected_halls = []
        selected_capacity = 0
        for hall in self.halls:
            cap = hall['rows'] * hall['cols']
            selected_halls.append(hall)
            selected_capacity += cap
            if selected_capacity >= required_seats:
                break

        selected_codes = {h['hall_code'] for h in selected_halls}

        # Allocate students only among selected halls, proportionally to their capacities.
        students_per_hall = {h['hall_code']: 0 for h in self.halls}
        if selected_capacity > 0 and total_students > 0:
            assigned_count = 0
            for hall in selected_halls:
                cap = hall['rows'] * hall['cols']
                # proportional share relative to selected_capacity
                count = int((cap / selected_capacity) * total_students)
                students_per_hall[hall['hall_code']] = count
                assigned_count += count

            # Fix rounding mismatch across selected halls
            remaining_to_place = total_students - assigned_count
            for hall in selected_halls:
                if remaining_to_place <= 0:
                    break
                students_per_hall[hall['hall_code']] += 1
                remaining_to_place -= 1

        # Halls not selected remain with zero allocation target

        # -------- NEW SEAT ALLOCATION LOGIC -------- #

        for hall in self.halls:
            hall_code = hall['hall_code']
            target = students_per_hall[hall_code]

            if target <= 0:
                continue

            rows = hall['rows']
            cols = hall['cols']
            grid_capacity = rows * cols
            hall_capacity = hall['total_capacity']

            usable_capacity = min(grid_capacity, hall_capacity, target)
            hall_students = []

            # Start with 2 largest remaining courses
            current_courses = []
            for c in courses_sorted:
                if remaining[c]:
                    current_courses.append(c)
                if len(current_courses) == 2:
                    break

            self.hall_courses[hall_code] = set(current_courses)

            # All other courses for future addition
            future_courses = [c for c in courses_sorted if c not in current_courses]

            # ROUND-ROBIN ALLOCATION
            while len(hall_students) < usable_capacity:

                # Cycle through current courses
                for course in list(current_courses):

                    if len(hall_students) >= usable_capacity:
                        break

                    if remaining[course]:
                        hall_students.append(remaining[course].pop(0))
                    else:
                        # COURSE EMPTY → immediately replace with next course
                        if future_courses:
                            next_course = future_courses.pop(0)
                            current_courses.append(next_course)
                            self.hall_courses[hall_code].add(next_course)
                        # Remove empty course
                        current_courses.remove(course)

                # If no courses available but seats remain → break
                if not current_courses:
                    break

            # Assign seats in S-pattern
            student_idx = 0
            for r in range(rows):
                col_range = range(cols) if r % 2 == 0 else range(cols - 1, -1, -1)
                for c in col_range:
                    if student_idx >= len(hall_students):
                        break

                    student = hall_students[student_idx]
                    reg = student['reg_no']
                    self.hall_seats[hall_code][r][c] = reg

                    self.allocations[reg] = {
                        'hall_code': hall_code,
                        'hall_name': hall['hall_name'],
                        'block': hall['block'],
                        'row': r + 1,
                        'col': c + 1,
                        'course_code': student['course_code'],
                        'course_title': student['course_title'],
                        'department': student['department'],
                    }
                    student_idx += 1

        # Identify unallocated students
        unallocated = sum(len(v) for v in remaining.values())
        if unallocated > 0:
            raise ValueError(f"{unallocated} students could not be allocated.")

    def print_seating_plan(self):
        """Print seating plan only for halls that have allocations."""
        for hall in self.halls:
            hall_code = hall['hall_code']
            if not any(self.hall_seats[hall_code][r][c] for r in range(hall['rows']) for c in range(hall['cols'])):
                continue  # skip halls with no allocated students

            print(f"\nHall: {hall['hall_name']} ({hall_code}, Block: {hall['block']})")
            courses_list = sorted(self.hall_courses[hall_code])
            print(f"Courses: {', '.join(courses_list) if courses_list else '-'}")
            for row_idx, row in enumerate(self.hall_seats[hall_code], 1):
                row_display = []
                for reg_no in row:
                    if reg_no:
                        course = self.allocations[reg_no]['course_code']
                        row_display.append(f"{reg_no} ({course})")
                    else:
                        row_display.append("-")
                print(f"Row {row_idx}: {' '.join(row_display)}")


    def export_pdf_seating_plan(self, output_dir="pdf_reports", filename="Seating_Plans.pdf"):
        """Export seating plans for all halls into a single PDF with headers & course summary (skip unallocated halls)."""
        os.makedirs(output_dir, exist_ok=True)
        styles = getSampleStyleSheet()

        # Single PDF file path
        filepath = os.path.join(output_dir, filename)
        doc = SimpleDocTemplate(filepath, pagesize=landscape(A4),
                                leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
        elements = []

        # Track last seat number per base hall code for continuous numbering
        last_seat_per_base = {}


        for idx, hall in enumerate(self.halls):
            hall_code = hall['hall_code']
            hall_name = hall['hall_name']
            block = hall['block']

            # Skip unallocated halls
            allocated_reg_nos = [self.hall_seats[hall_code][r][c] for r in range(hall['rows']) for c in range(hall['cols']) if self.hall_seats[hall_code][r][c]]
            if not allocated_reg_nos:
                continue

            # Header image
            header_img_path = os.path.join(os.path.dirname(__file__), "Header_Hall_Seating.jpg")
            if os.path.exists(header_img_path):
                elements.append(Image(header_img_path, width=660, height=114))

            # Header
            elements.append(Paragraph(f"<b>Block: {block} | Hall Code: {hall_code}</b>", styles['Heading3']))
            # Prefer allocator-level exam_date/session, otherwise use first allocation date for this hall
            hall_exam_date = getattr(self, 'exam_date', None)
            if not hall_exam_date:
                hall_exam_date = None
                for reg_no, alloc in self.allocations.items():
                    if alloc.get('hall_code') == hall_code and alloc.get('date'):
                        hall_exam_date = alloc.get('date')
                        break
            hall_exam_date = hall_exam_date or ""
            hall_session = getattr(self, 'session', None) or ""
            hall_session_str = f" | Session: {hall_session}" if hall_session else ""
            elements.append(Paragraph(f"<b>Date of Exam:</b> {hall_exam_date}{hall_session_str}", styles['Heading3']))

           

            # Setup variables for this hall
            num_cols = 5  # Change this to match your desired column count
            seat_data = []
            total_seats = len(allocated_reg_nos)
            reg_nos = allocated_reg_nos[:]
            num_rows = (total_seats + num_cols - 1) // num_cols

            # Determine base hall code for continuous seat numbering. Use right-most split so
            # halls like 'HALL-1' and 'HALL-2' share the same base code (e.g. 'HALL').
            base = hall_code.rsplit('-', 1)[0] if '-' in hall_code else hall_code
            seat_number = last_seat_per_base.get(base, 0) + 1

            # S-format seat numbering: seat number is based on physical position, not allocation order.
            # If this hall is a continuation of a previous hall with the same base (e.g. Hall-2 after Hall-1),
            # start numbering the first row in right-to-left order so the visual flow continues correctly.
            continue_in_reverse = base in last_seat_per_base
            seat_no_grid = [[None for _ in range(num_cols)] for _ in range(num_rows)]
            for row in range(num_rows):
                # determine whether this row should be left-to-right
                if (row % 2 == 0 and not continue_in_reverse) or (row % 2 == 1 and continue_in_reverse):
                    # left-to-right
                    for col in range(num_cols):
                        seat_no_grid[row][col] = seat_number
                        seat_number += 1
                else:
                    # right-to-left
                    for col in range(num_cols-1, -1, -1):
                        seat_no_grid[row][col] = seat_number
                        seat_number += 1

            # Update last seat for this base
            last_seat_per_base[base] = seat_number - 1

            # Fill seat_data with S-format seat numbers and reg_nos
            idx = 0
            for row in range(num_rows):
                row_cells = []
                for col in range(num_cols):
                    if idx < total_seats:
                        row_cells.append([str(seat_no_grid[row][col]), reg_nos[idx]])
                        idx += 1
                    else:
                        row_cells.append(["", ""])
                seat_data.append(row_cells)

            # Header row
            header_row = []
            for i in range(num_cols):
                header_row.append("Seat No.")
                header_row.append("Reg. No.")

            # Flatten rows
            table_data = [header_row]
            for row in seat_data:
                flat_row = []
                for cell in row:
                    flat_row.append(cell[0])  # Seat No
                    flat_row.append(cell[1])  # Reg No
                table_data.append(flat_row)

            # Assign colors to courses
            color_list = [colors.lightblue, colors.lightgreen, colors.lightyellow, colors.lightpink, colors.lightcyan, colors.lightyellow, colors.lightcoral]
            course_color = {}
            for i, course in enumerate(sorted(self.hall_courses[hall_code])):
                course_color[course] = color_list[i % len(color_list)]

            # Create table
            seat_table = Table(table_data, colWidths=[50, 90] * num_cols, rowHeights=[25] * len(table_data))
            styles_list = [
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ]
            # Add background colors for reg_no cells
            for row_idx in range(1, len(table_data)):
                for col_idx in range(1, len(table_data[row_idx]), 2):
                    reg_no = table_data[row_idx][col_idx]
                    if reg_no and reg_no in self.allocations:
                        course = self.allocations[reg_no]['course_code']
                        bg_color = course_color.get(course, colors.white)
                        styles_list.append(('BACKGROUND', (col_idx, row_idx), (col_idx, row_idx), bg_color))
            seat_table.setStyle(TableStyle(styles_list))
            elements.append(seat_table)
            elements.append(Spacer(1, 5))

            # Summary Table (Hall no., Sub Code, Branch, Register number, No. of Students, Total)
            elements.append(Paragraph("<b>Hall Summary</b>", styles['Heading3']))
            summary_data = [["Hall no.", "Sub Code", "Branch", "Register number", "No. of Students", "Total"]]
            grouped = defaultdict(lambda: defaultdict(list))
            for reg_no, alloc in self.allocations.items():
                if alloc['hall_code'] == hall_code:
                    grouped[alloc['course_code']][alloc['department']].append(reg_no)

            total_students = 0
            for course_code, branches in grouped.items():
                for branch, regnos in branches.items():
                    regnos_sorted = sorted(regnos)
                    # Format register numbers as ranges and singles, separated by commas
                    ranges = []
                    start = prev = regnos_sorted[0]
                    for r in regnos_sorted[1:]:
                        try:
                            if int(r) == int(prev) + 1:
                                prev = r
                            else:
                                ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                                start = prev = r
                        except ValueError:
                            ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                            start = prev = r
                    ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                    # Format: 2 register number ranges per line
                    regno_lines = []
                    for i in range(0, len(ranges), 2):
                        regno_lines.append(", ".join(ranges[i:i+2]))
                    regno_str = "\n".join(regno_lines)
                    count = len(regnos_sorted)
                    total_students += count
                    summary_data.append([
                        hall_code,
                        course_code,
                        branch,
                        regno_str,
                        f"{count:02}",
                        ""
                    ])
            # Add total in last row
            if len(summary_data) > 1:
                summary_data[-1][-1] = str(total_students)

            # Calculate dynamic row heights based on register number lines
            rowHeights = []
            for i, row in enumerate(summary_data):
                if i == 0:  # header row
                    height = 25
                else:
                    regno_str = row[3]
                    lines = regno_str.count('\n') + 1
                    height = 20 + (lines - 1) * 10  # 10 units per additional line
                rowHeights.append(height)

            summary_table = Table(summary_data, repeatRows=1, colWidths=[70, 70, 70, 350, 100, 50], rowHeights=rowHeights)
            # Build style for summary table and add background color for course code column per course
            summary_style = [
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (3, 0), (3, -1), 'MIDDLE'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                ('FONTSIZE', (0,0), (-1,-1), 12),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ]

            # Color the course code cell (column index 1) per row using course_color mapping
            for row_idx in range(1, len(summary_data)):
                try:
                    course_cell = summary_data[row_idx][1]
                    if course_cell and course_cell in course_color:
                        bg = course_color.get(course_cell, colors.white)
                        summary_style.append(('BACKGROUND', (1, row_idx), (1, row_idx), bg))
                except Exception:
                    # ignore rows that don't match expected format
                    pass

            summary_table.setStyle(TableStyle(summary_style))
            elements.append(summary_table)
            elements.append(Spacer(1, 12))

            # Add page break if not last hall
            if idx != len(self.halls) - 1:
                elements.append(PageBreak())

        # Build single PDF (no footer on pages)
        doc.build(elements)
        print(f"PDF Exported: {filepath}")




    def export_master_seating_plan(self, output_dir="pdf_reports", output_file="Master_Seating_Plan.pdf"):
        """Generate a master seating plan report with hall, department, course, regno ranges, totals."""
        styles = getSampleStyleSheet()

        def footer(canvas, doc):
            canvas.saveState()
            canvas.setFont('Helvetica', 10)
            canvas.drawRightString(doc.pagesize[0] - 50, 30, "Chief Superintendent")
            canvas.restoreState()

        doc = SimpleDocTemplate(os.path.join(output_dir, output_file), pagesize=portrait(A4),
                                leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=22, onPage=footer)
        elements = []

        # Header image
        header_img_path = "Header_Hall_Seating.jpg"
        if os.path.exists(header_img_path):
            elements.append(Image(header_img_path, width=550, height=119))

        # Header: prefer allocator-level exam_date/session, otherwise fall back to any allocation date
        master_date = getattr(self, 'exam_date', None)
        if not master_date:
            for reg_no, alloc in self.allocations.items():
                if alloc.get('date'):
                    master_date = alloc.get('date')
                    break
        master_date = master_date or ""
        master_session = getattr(self, 'session', None) or ""
        master_session_str = f" | Session: {master_session}" if master_session else ""
        elements.append(Paragraph(f"<b>Date of Exam:</b> {master_date}{master_session_str}", styles['Heading3']))
        elements.append(Spacer(1, 12))

        # Collect data by block
        block_data = defaultdict(list)
        hall_summary = defaultdict(int)
        grand_total = 0

        # Group allocations: hall -> dept -> course
        grouped = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        for reg_no, alloc in self.allocations.items():
            grouped[alloc['hall_code']][alloc['department']][alloc['course_code']].append(reg_no)

        # Prepare rows per block
        for hall in self.halls:
            hall_code = hall['hall_code']
            hall_name = hall['hall_name']
            block = hall['block']

            if hall_code not in grouped:
                continue

            for dept, courses in grouped[hall_code].items():
                for course_code, regnos in courses.items():
                    regnos_sorted = sorted(regnos)

                    # ✅ Create continuous ranges of register numbers
                    ranges = []
                    start = prev = regnos_sorted[0]
                    for r in regnos_sorted[1:]:
                        try:
                            if int(r) == int(prev) + 1:  # continuous
                                prev = r
                            else:
                                ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                                start = prev = r
                        except ValueError:
                            ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                            start = prev = r
                    ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                    # Format: 2 register number ranges per line
                    regno_lines = []
                    for i in range(0, len(ranges), 2):
                        regno_lines.append(", ".join(ranges[i:i+2]))
                    regno_range_str = "\n".join(regno_lines)

                    count = len(regnos_sorted)
                    grand_total += count
                    hall_summary[hall_code] += count

                    block_data[block].append([
                        hall_name,
                        dept,
                        course_code,
                        regno_range_str,
                        str(count)
                    ])

        # Create tables per block
        for block in sorted(block_data.keys()):
            elements.append(Paragraph(f"<b>Block: {block}</b>", styles['Heading2']))
            elements.append(Spacer(1, 6))
            table_data = [["Hall Name", "Department", "Course Code", "Register No. From–To", "Total Students"]] + block_data[block]
            table = Table(table_data, repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(table)
            elements.append(Spacer(1, 12))

        # Summary Section on new page: repeat header image and date for clarity
        elements.append(PageBreak())
        # Repeat header image on summary page if available
        header_img_path = "Header_Master Seating.jpg"
        if os.path.exists(header_img_path):
            elements.append(Image(header_img_path, width=550, height=119))
        elements.append(Paragraph("<b>Date of Exam:</b> 16-09-2025", styles['Heading3']))
        elements.append(Spacer(1, 6))

        # Department & course totals (left) and Hall totals (right) in two columns
        dept_data = [["Department", "Course Code", "Total Students"]]
        dept_course_totals = defaultdict(lambda: defaultdict(int))
        for reg_no, alloc in self.allocations.items():
            dept = alloc['department']
            course_code = alloc['course_code']
            dept_course_totals[dept][course_code] += 1
        for dept, courses in dept_course_totals.items():
            for course_code, total in courses.items():
                dept_data.append([dept, course_code, str(total)])
        dept_table = Table(dept_data, hAlign="LEFT")
        dept_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ]))

        # Hall totals
        hall_data = [["Hall", "Total Students"]]
        for hall in self.halls:
            hall_code = hall['hall_code']
            if hall_code in hall_summary:
                hall_data.append([f"{hall['hall_name']}", str(hall_summary[hall_code])])
        hall_table = Table(hall_data, hAlign="LEFT")
        hall_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ]))

        # Create nested tables so both sections appear side-by-side in two columns
        left_nested = Table([[Paragraph('<b>By Department & Course:</b>', styles['Heading3'])], [dept_table]], colWidths=[270])
        right_nested = Table([[Paragraph('<b>By Hall:</b>', styles['Heading3'])], [hall_table]], colWidths=[270])
        # Parent table with two columns
        two_col = Table([[left_nested, right_nested]], colWidths=[270, 270])
        two_col.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ]))

        elements.append(two_col)
        elements.append(Spacer(1, 12))

        # Grand total
        elements.append(Paragraph(f"<b>Grand Total Students: {grand_total}</b>", styles['Heading2']))

        # Build master PDF and ensure footer is drawn on every page
        doc.build(elements, onFirstPage=footer, onLaterPages=footer)
        print(f"Master PDF Exported: {output_file}")


# Example Usage
if __name__ == "__main__":
    try:
        allocator = ExamSeatAllocator(excel_file='exam_seat_allocation 19 FN.xlsx')
        allocator.allocate_seats()
        allocator.print_seating_plan()
        allocator.export_pdf_seating_plan()
        allocator.export_master_seating_plan()

    except Exception as e:
        print(f"Error: {e}")
