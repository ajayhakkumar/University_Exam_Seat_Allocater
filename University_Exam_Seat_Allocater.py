import pandas as pd
from collections import defaultdict, deque
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
        """Allocate seats using the user's requested algorithm:
      1) Split student list into first and second half.
      2) Compute proportional students per hall.
      3) Round-robin allocate alternating from first and second half into halls; place in S-pattern.
        """

        self.validate_inputs()

        # 1) directly split for allocation
        students_list = list(self.students)

        # === Select minimal set of halls in 25-seat blocks ===
        # Decide required capacity as next multiple of 25 that can hold all students
        total_students = len(students_list)
        required_capacity = ((total_students + 24) // 25) * 25

        # Compute grid capacity per hall (rows * cols). Choose halls that minimize
        # the number of halls needed by picking largest-capacity halls first, but
        # preserve the original ordering later for placement.
        halls_by_capacity = sorted(self.halls, key=lambda h: h['rows'] * h['cols'], reverse=True)
        selected_codes = set()
        cap_acc = 0
        for h in halls_by_capacity:
            cap_acc += h['rows'] * h['cols']
            selected_codes.add(h['hall_code'])
            if cap_acc >= required_capacity:
                break

        # Preserve original ordering of halls but only keep selected ones
        halls_to_use = [h for h in self.halls if h['hall_code'] in selected_codes]
        # If for some reason we couldn't reach required_capacity (should be caught earlier), raise
        if sum(h['rows'] * h['cols'] for h in halls_to_use) < total_students:
            raise ValueError("Not enough selected hall seats to allocate all students after 25-seat-block selection.")

        # 2) Split into first and second half
        half = (total_students + 1) // 2
        first_half = deque(students_list[:half])
        second_half = deque(students_list[half:])
        # 3) Compute proportional students per selected hall based on grid capacity
        total_capacity = sum(h['rows'] * h['cols'] for h in halls_to_use)
        students_per_hall = {}
        allocated_count = 0
        for hall in halls_to_use:
            cap = hall['rows'] * hall['cols']
            count = int((cap / total_capacity) * total_students) if total_capacity > 0 else 0
            students_per_hall[hall['hall_code']] = count
            allocated_count += count

        # Distribute any rounding remainder to selected halls in original order
        rem = total_students - allocated_count
        for hall in halls_to_use:
            if rem <= 0:
                break
            students_per_hall[hall['hall_code']] += 1
            rem -= 1

        # 4) Allocate per selected hall in original order, pulling alternately from first and second half
        for hall in halls_to_use:
            hall_code = hall['hall_code']
            target = students_per_hall.get(hall_code, 0)
            if target <= 0:
                continue

            rows = hall['rows']
            cols = hall['cols']
            grid_capacity = rows * cols
            hall_capacity = hall['total_capacity']
            usable = min(grid_capacity, hall_capacity, target)

            hall_students = []

            # Round-robin: take one from first_half then one from second_half repeatedly
            while len(hall_students) < usable and (first_half or second_half):
                if first_half and len(hall_students) < usable:
                    hall_students.append(first_half.popleft())
                if second_half and len(hall_students) < usable:
                    hall_students.append(second_half.popleft())

            # If not enough students (both halves empty), continue
            if not hall_students:
                continue

            # Record courses used in this hall
            self.hall_courses[hall_code] = set(s['course_code'] for s in hall_students)

            # Place students in S-pattern (sequential order from hall_students)
            idx = 0
            for r in range(rows):
                if idx >= len(hall_students):
                    break
                col_range = range(cols) if r % 2 == 0 else range(cols - 1, -1, -1)
                for c in col_range:
                    if idx >= len(hall_students):
                        break
                    student = hall_students[idx]
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
                    idx += 1

        # After filling halls, check for unallocated students
        remaining = list(first_half) + list(second_half)
        if remaining:
            rem_regs = [s['reg_no'] for s in remaining]
            raise ValueError(f"Could not allocate all students. Remaining: {len(remaining)}. RegNos: {', '.join(rem_regs)}")

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

           

            # Use the hall's actual grid for printing so seat numbers and
            # registration numbers align with the allocation grid.
            rows = hall['rows']
            cols = hall['cols']

            # Determine base hall code for continuous seat numbering. Use right-most split so
            # halls like 'HALL-1' and 'HALL-2' share the same base code (e.g. 'HALL').
            base = hall_code.rsplit('-', 1)[0] if '-' in hall_code else hall_code
            seat_number = last_seat_per_base.get(base, 0) + 1

            # If this hall continues a previous base, reverse the first-row direction
            continue_in_reverse = base in last_seat_per_base

            # Build a seat number grid matching the physical layout (rows x cols)
            seat_no_grid = [[None for _ in range(cols)] for _ in range(rows)]
            for r in range(rows):
                # Determine traversal direction for this row matching S-pattern
                left_to_right = (r % 2 == 0 and not continue_in_reverse) or (r % 2 == 1 and continue_in_reverse)
                if left_to_right:
                    for c in range(cols):
                        seat_no_grid[r][c] = seat_number
                        seat_number += 1
                else:
                    for c in range(cols - 1, -1, -1):
                        seat_no_grid[r][c] = seat_number
                        seat_number += 1

            # Update last seat for this base
            last_seat_per_base[base] = seat_number - 1

            # Header row: Seat No. / Reg. No. repeated per column
            header_row = []
            for _ in range(cols):
                header_row.extend(["Seat No.", "Reg. No."])

            # Build table_data by reading reg_nos directly from self.hall_seats so
            # the printed reg numbers match actual allocated positions
            table_data = [header_row]
            for r in range(rows):
                flat_row = []
                for c in range(cols):
                    seat_no = seat_no_grid[r][c]
                    reg_no = self.hall_seats[hall_code][r][c] or ""
                    flat_row.append(str(seat_no) if seat_no is not None else "")
                    flat_row.append(reg_no)
                table_data.append(flat_row)

            # Assign colors to courses
            color_list = [colors.lightblue, colors.lightgreen, colors.lightyellow, colors.lightpink, colors.lightcyan, colors.lightyellow, colors.lightcoral]
            course_color = {}
            for i, course in enumerate(sorted(self.hall_courses[hall_code])):
                course_color[course] = color_list[i % len(color_list)]

            # Create table
            seat_table = Table(table_data, colWidths=[50, 90] * cols, rowHeights=[25] * len(table_data))
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
                    # Format: 1 register number range per line (no comma)
                    regno_lines = []
                    for i in range(0, len(ranges), 1):
                        regno_lines.append(", ".join(ranges[i:i+1]))
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

            summary_table = Table(summary_data, repeatRows=1, colWidths=[50, 50, 50, 250, 70, 70], rowHeights=rowHeights)
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

        doc = SimpleDocTemplate(os.path.join(output_dir, output_file), pagesize=landscape(A4),
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
        blocks = sorted(block_data.keys())
        for bi, block in enumerate(blocks):
            elements.append(Paragraph(f"<b>Block: {block}</b>", styles['Heading2']))
            elements.append(Spacer(1, 6))

            # Build table with an extra column for "Total in Hall" placed as the last column
            # Header
            table_rows = [["Hall Name", "Department", "Course Code", "Register No. From–To", "Total Students", "Total in Hall"]]
            # Append block rows (each row: [hall_name, dept, course_code, regno_range_str, count])
            for r in block_data[block]:
                # r is [hall_name, dept, course_code, regno_range_str, str(count)]
                # Keep count as integer for totals
                try:
                    count_val = int(r[4])
                except Exception:
                    count_val = 0
                # place course count under 'Total Students' (index 4), leave 'Total in Hall' (index 5) blank
                table_rows.append([r[0], r[1], r[2], r[3], str(count_val), ""])

            # Compute vertical spans for identical consecutive hall names and fill "Total in Hall" in first row of each hall
            spans = []
            hall_ranges = []
            i = 1
            block_total = 0
            while i < len(table_rows):
                hall_name = table_rows[i][0]
                start = i
                # Sum course counts which are stored in column index 4 ('Total Students')
                hall_sum = int(table_rows[i][4]) if table_rows[i][4] else 0
                i += 1
                while i < len(table_rows) and table_rows[i][0] == hall_name:
                    try:
                        hall_sum += int(table_rows[i][4])
                    except Exception:
                        pass
                    i += 1
                end = i - 1
                # Fill total in first row for this hall into the last column (index 5)
                table_rows[start][5] = str(hall_sum)
                block_total += hall_sum
                # Record the hall row range for later styling (background) and spanning
                hall_ranges.append((start, end))
                # If multiple rows for same hall, span the Hall Name column vertically
                # Also span the 'Total in Hall' column (last column index 5) so the hall total
                # appears as a single merged cell beside the group's rows.
                if end > start:
                    spans.append(('SPAN', (0, start), (0, end)))
                    spans.append(('SPAN', (5, start), (5, end)))
                # continue from next

            # Add final summary row for this block: place the label in the left-most spanned cell
            table_rows.append(["Total Sum of Students", "", "", "", "", str(block_total)])
            # If this is the last block, also append a Grand Total row showing totals across all blocks
            if bi == len(blocks) - 1:
                table_rows.append(["Grand Total (All Blocks)", "", "", "", "", str(grand_total)])

            # Create table and apply styles including spans
            table = Table(table_rows, repeatRows=1, colWidths=[90, 90, 90, 350, 80, 80])
            summary_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
            ]

            # Add spans computed earlier
            for sp in spans:
                summary_style.append(sp)

            # Apply alternating background colors per hall range (do not color header or final total row)
            # Choose two subtle alternating colors
            alt_colors = [colors.lightblue, colors.lightgreen]
            for hi, (s, e) in enumerate(hall_ranges):
                try:
                    bg = alt_colors[hi % len(alt_colors)]
                    # Apply background across all data columns for the hall rows
                    summary_style.append(('BACKGROUND', (0, s), (5, e), bg))
                except Exception:
                    pass

            # Style the final total row(s)
            last_row_idx = len(table_rows) - 1
            # Span first four columns for the label on the last row (Grand Total or Block Total if single)
            summary_style.append(('SPAN', (0, last_row_idx), (3, last_row_idx)))
            summary_style.append(('BACKGROUND', (0, last_row_idx), (-1, last_row_idx), colors.lightgrey))
            summary_style.append(('FONTNAME', (0, last_row_idx), (-1, last_row_idx), 'Helvetica-Bold'))
            # If we added both block total and grand total, also style the block total row (second last)
            if last_row_idx >= 1 and table_rows[last_row_idx - 1][0] == 'Total Sum of Students':
                prev_idx = last_row_idx - 1
                summary_style.append(('SPAN', (0, prev_idx), (3, prev_idx)))
                summary_style.append(('BACKGROUND', (0, prev_idx), (-1, prev_idx), colors.lightgrey))
                summary_style.append(('FONTNAME', (0, prev_idx), (-1, prev_idx), 'Helvetica-Bold'))

            table.setStyle(TableStyle(summary_style))
            elements.append(table)
            elements.append(Spacer(1, 12))
            # Start each block on a new page except after the last block
            if bi != len(blocks) - 1:
                elements.append(PageBreak())

        # Summary Section on new page: repeat header image and date for clarity
        elements.append(PageBreak())
        # Repeat header image on summary page if available
        header_img_path = "Header_Master Seating.jpg"
        if os.path.exists(header_img_path):
            elements.append(Image(header_img_path, width=550, height=119))
        # Reuse previously computed master_date/master_session so the summary page
        # prints the dynamic date/session rather than a hard-coded value.
        elements.append(Paragraph(f"<b>Date of Exam:</b> {master_date}{master_session_str}", styles['Heading3']))
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
        # Append a Grand Total row for the Department & Course table
        # Use grand_total computed earlier (sum of students across blocks)
        try:
            dept_grand_total = int(grand_total)
        except Exception:
            # Fallback: sum the dept_course_totals
            dept_grand_total = sum(sum(int(v) for v in courses.values()) for courses in dept_course_totals.values())
        dept_data.append(["Grand Total", "", str(dept_grand_total)])

        # Create dept table and style header and grand-total row specially
        dept_table = Table(dept_data, hAlign="LEFT")
        dept_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ]
        last_dept_row = len(dept_data) - 1
        dept_style.append(('BACKGROUND', (0, last_dept_row), (-1, last_dept_row), colors.lightgrey))
        dept_style.append(('FONTNAME', (0, last_dept_row), (-1, last_dept_row), 'Helvetica-Bold'))
        dept_table.setStyle(TableStyle(dept_style))

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

        # If either of the summary tables is long, avoid nesting them side-by-side
        # inside a single Table cell because ReportLab cannot split a very tall
        # nested flowable across the page frame. Instead, fall back to stacking
        # them vertically.
        max_rows_side_by_side = 20
        dept_rows = len(dept_data)
        hall_rows = len(hall_data)

        try:
            if dept_rows <= max_rows_side_by_side and hall_rows <= max_rows_side_by_side:
                # Parent table with two columns
                two_col = Table([[left_nested, right_nested]], colWidths=[270, 270])
                two_col.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0,0), (-1,-1), 0),
                    ('RIGHTPADDING', (0,0), (-1,-1), 12),
                ]))
                elements.append(two_col)
            else:
                # Stack vertically to allow proper page splitting
                elements.append(Paragraph('<b>By Department & Course:</b>', styles['Heading3']))
                elements.append(dept_table)
                elements.append(Spacer(1, 12))
                elements.append(Paragraph('<b>By Hall:</b>', styles['Heading3']))
                elements.append(hall_table)
        except Exception:
            # On any unexpected issue, fall back to vertical stacking
            elements.append(Paragraph('<b>By Department & Course:</b>', styles['Heading3']))
            elements.append(dept_table)
            elements.append(Spacer(1, 12))
            elements.append(Paragraph('<b>By Hall:</b>', styles['Heading3']))
            elements.append(hall_table)
        elements.append(Spacer(1, 12))

        # Build master PDF and ensure footer is drawn on every page
        doc.build(elements, onFirstPage=footer, onLaterPages=footer)
        print(f"Master PDF Exported: {output_file}")


# Example Usage
if __name__ == "__main__":
    try:
        allocator = ExamSeatAllocator(excel_file='exam_seat_allocation 19 AN.xlsx')
        allocator.allocate_seats()
        allocator.print_seating_plan()
        allocator.export_pdf_seating_plan()
        allocator.export_master_seating_plan()

    except Exception as e:
        print(f"Error: {e}")
