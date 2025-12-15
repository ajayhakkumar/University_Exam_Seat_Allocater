import pandas as pd
from collections import defaultdict
import os
from reportlab.lib.pagesizes import A4, portrait, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image



class ExamSeatAllocator:
    def __init__(self, students_file, halls_file):
        """
        Initialize with paths to Excel files for students and halls.
        """
        self.students = self._load_students(students_file)
        self.halls = self._load_halls(halls_file)
        self.allocations = {}
        self.hall_seats = {
            hall['hall_code']: [[None for _ in range(hall['cols'])] for _ in range(hall['rows'])]
            for hall in self.halls
        }
        self.hall_courses = {hall['hall_code']: set() for hall in self.halls}

    def _load_halls(self, halls_file):
        """Load hall details from Excel file."""
        try:
            df = pd.read_excel(halls_file)
            required_columns = ['Hall Code', 'Hall Name', 'Block', 'Row', 'Column', 'Total Capacity']
            if not all(col in df.columns for col in required_columns):
                raise ValueError(f"Halls file must contain columns: {', '.join(required_columns)}")

            halls = []
            for _, row in df.iterrows():
                halls.append({
                    'hall_code': str(row['Hall Code']).strip(),
                    'hall_name': str(row['Hall Name']).strip(),
                    'block': str(row['Block']).strip(),
                    'rows': int(row['Row']),
                    'cols': int(row['Column']),
                    'total_capacity': int(row['Total Capacity'])
                })
            return halls
        except Exception as e:
            raise ValueError(f"Error reading halls file: {e}")

    def _load_students(self, students_file):
        """Load student details from Excel file."""
        try:
            df = pd.read_excel(students_file)
            required_columns = ['Student Reg.No.', 'Department', 'Course Code', 'Course Title', 'Date']
            if not all(col in df.columns for col in required_columns):
                raise ValueError(f"Students file must contain columns: {', '.join(required_columns)}")

            students = []
            for _, row in df.iterrows():
                students.append({
                    'reg_no': str(row['Student Reg.No.']).strip(),
                    'department': str(row['Department']).strip(),
                    'course_code': str(row['Course Code']).strip(),
                    'course_title': str(row['Course Title']).strip(),
                    'date': str(row['Date']).strip(),
                })
            return students
        except Exception as e:
            raise ValueError(f"Error reading students file: {e}")

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
        """Allocate students using round-robin across courses, with global balancing for remaining students."""
        self.validate_inputs()

        # Sort students by reg_no ascending inside each course
        students_sorted = sorted(self.students, key=lambda x: (x['course_code'], x['reg_no']))

        # Group by course
        students_by_course = defaultdict(list)
        for s in students_sorted:
            students_by_course[s['course_code']].append(s)

        # Sort courses by student count (largest first)
        courses_by_size = sorted(students_by_course.keys(), key=lambda c: len(students_by_course[c]), reverse=True)

        remaining_courses = {c: list(students_by_course[c]) for c in courses_by_size}

        # First pass: Allocate to halls using round-robin
        for hall in self.halls:
            hall_code = hall['hall_code']
            grid_capacity = hall['rows'] * hall['cols']
            hall_capacity = hall['total_capacity']

            # ✅ Adjust capacity if mismatch (use only declared total_capacity)
            usable_capacity = min(hall_capacity, grid_capacity)

            hall_students = []
            current_courses = []

            # Pick up to 3 available courses for this hall
            for course in courses_by_size:
                if remaining_courses.get(course):
                    current_courses.append(course)
                if len(current_courses) == 4:
                    break

            # Store courses for this hall
            self.hall_courses[hall_code] = set(current_courses)

            if not current_courses:
                continue  # no students left for this hall

            # Allocate in 2:1:1 ratio if exactly 3 courses (A:B:C)
            if len(current_courses) == 3:
                # Assuming current_courses[0] is A (largest), [1] B, [2] C
                allocation_order = [0, 0, 1, 2]  # 2:1:1 ratio
                idx = 0
                while len(hall_students) < usable_capacity and any(remaining_courses[c] for c in current_courses):
                    course_idx = allocation_order[idx % len(allocation_order)]
                    course = current_courses[course_idx]
                    if remaining_courses[course]:
                        hall_students.append(remaining_courses[course].pop(0))
                    idx += 1
            else:
                # Fallback to round-robin for other cases
                while len(hall_students) < usable_capacity and any(remaining_courses[c] for c in current_courses):
                    for course in current_courses:
                        if remaining_courses[course] and len(hall_students) < usable_capacity:
                            hall_students.append(remaining_courses[course].pop(0))

            # Sort hall_students by reg_no ascending (numeric when possible)
            try:
                import re

                def _reg_sort_key(student):
                    reg = student.get('reg_no') if isinstance(student, dict) else str(student)
                    reg = str(reg)
                    digits = re.sub(r"\D", "", reg)
                    if digits:
                        return (0, int(digits))
                    return (1, reg)

                hall_students.sort(key=_reg_sort_key)
            except Exception:
                pass

            # Row-wise seat assignment
            student_idx = 0
            for row in range(hall['rows']):
                for col in range(hall['cols']):
                    if student_idx >= len(hall_students):
                        break
                    student = hall_students[student_idx]
                    self.hall_seats[hall_code][row][col] = student['reg_no']
                    self.allocations[student['reg_no']] = {
                        'hall_code': hall_code,
                        'hall_name': hall['hall_name'],
                        'block': hall['block'],
                        'row': row + 1,
                        'col': col + 1,
                        'course_code': student['course_code'],
                        'course_title': student['course_title'],
                        'department': student['department'],
                        'date': student['date'],
                    }
                    student_idx += 1

        # Second pass: Allocate remaining students to halls with available space
        unallocated_students = []
        for course, studs in remaining_courses.items():
            unallocated_students.extend(studs)

        if unallocated_students:
            # Sort unallocated students by course and reg_no
            # numeric-aware reg_no sort for predictable physical ordering when placed
            try:
                import re

                def _reg_sort_key2(student):
                    course = student.get('course_code')
                    reg = str(student.get('reg_no'))
                    digits = re.sub(r"\D", "", reg)
                    if digits:
                        return (course, 0, int(digits))
                    return (course, 1, reg)

                unallocated_students.sort(key=_reg_sort_key2)
            except Exception:
                unallocated_students.sort(key=lambda x: (x['course_code'], x['reg_no']))

            for hall in self.halls:
                hall_code = hall['hall_code']
                grid_capacity = hall['rows'] * hall['cols']
                hall_capacity = hall['total_capacity']
                usable_capacity = min(hall_capacity, grid_capacity)

                # Find available seats in this hall
                allocated_count = sum(1 for row in self.hall_seats[hall_code] for seat in row if seat is not None)
                available_seats = usable_capacity - allocated_count

                if available_seats > 0 and unallocated_students:
                    # Allocate up to available_seats from unallocated_students
                    to_allocate = min(available_seats, len(unallocated_students))
                    for _ in range(to_allocate):
                        student = unallocated_students.pop(0)
                        # Find next available seat
                        for row in range(hall['rows']):
                            for col in range(hall['cols']):
                                if self.hall_seats[hall_code][row][col] is None:
                                    self.hall_seats[hall_code][row][col] = student['reg_no']
                                    self.allocations[student['reg_no']] = {
                                        'hall_code': hall_code,
                                        'hall_name': hall['hall_name'],
                                        'block': hall['block'],
                                        'row': row + 1,
                                        'col': col + 1,
                                        'course_code': student['course_code'],
                                        'course_title': student['course_title'],
                                        'department': student['department'],
                                        'date': student['date'],
                                    }
                                    break
                            else:
                                continue
                            break

        # Check for unallocated students
        unallocated = sum(len(studs) for studs in remaining_courses.values())
        if unallocated > 0:
            raise ValueError(f"Could not allocate all students. Remaining unallocated: {unallocated}")

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

        for idx, hall in enumerate(self.halls):
            hall_code = hall['hall_code']

            # Skip unallocated halls
            if not any(self.hall_seats[hall_code][r][c] for r in range(hall['rows']) for c in range(hall['cols'])):
                continue

            hall_name = hall['hall_name']
            block = hall['block']
            hall_courses = sorted(self.hall_courses[hall_code])

            # Collect summary per course for this hall
            summary = defaultdict(int)
            for reg_no, alloc in self.allocations.items():
                if alloc['hall_code'] == hall_code:
                    summary[alloc['course_code']] += 1

            # Use exam date from first student in this hall (if any)
            exam_date = None
            for reg_no, alloc in self.allocations.items():
                if alloc['hall_code'] == hall_code:
                    exam_date = alloc['date']
                    break

            # Header image
            header_img_path = os.path.join(os.path.dirname(__file__), "Header.png")
            if os.path.exists(header_img_path):
                elements.append(Image(header_img_path, width=550, height=119))

            # Header
            elements.append(
                Paragraph(
                    f"<b>Continuous Assessment Examination -2<br/>Seating Plan ({hall_code}, Block: {block})</b>", 
                    styles['Title']
                )
            )
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
            # Course Summary Table
            elements.append(Paragraph("<b>Courses in this Hall:</b>", styles['Heading3']))
            summary_data = [["Course Code", "Course Title", "No. of Students", "--------- Absentees ---------"]]
            for course in hall_courses:
                count = summary.get(course, 0)
                course_title = next((s['course_title'] for s in self.students if s['course_code'] == course), "")
                summary_data.append([course, course_title, str(count), ""])  # Blank Absentees column

            summary_table = Table(summary_data, hAlign='LEFT')
            summary_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
            ]))
            elements.append(summary_table)
            elements.append(Spacer(1, 12))

            # Seating Table
            data = [["Row/Col"] + [f"C{c+1}" for c in range(hall['cols'])]]
            for r in range(hall['rows']):
                row_data = [f"R{r+1}"]
                for c in range(hall['cols']):
                    reg_no = self.hall_seats[hall_code][r][c]
                    if reg_no:
                        course = self.allocations[reg_no]['course_code']
                        row_data.append(f"{reg_no}\n({course})")
                    else:
                        row_data.append("-")
                data.append(row_data)

            table = Table(data, repeatRows=1)
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

            # Add page break if not last hall
            if idx != len(self.halls) - 1:
                elements.append(PageBreak())

        # Build single PDF
        doc.build(elements)
        print(f"PDF Exported: {filepath}")




    def export_master_seating_plan(self, output_file="Master_Seating_Plan.pdf"):
            """Generate a master seating plan report with hall, department, course, regno ranges, totals."""
            styles = getSampleStyleSheet()
            doc = SimpleDocTemplate(output_file, pagesize=portrait(A4),
                                    leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
            elements = []

            # Header image
            header_img_path = "Header.png"
            if os.path.exists(header_img_path):
                elements.append(Image(header_img_path, width=550, height=119))

            # Header
            elements.append(
                Paragraph(
                    "<b>Continuous Assessment Examination -2<br/>Master Seating Plan</b>",
                    styles['Title']
                )
            )

            # Master date/session: prefer allocator attributes then any allocation date
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

            # Collect data
            master_data = [["Block", "Hall Name", "Department", "Course Code",
                            "Register No. From–To", "Total Students"]]

            hall_summary = defaultdict(int)
            grand_total = 0

            # Group allocations: hall -> dept -> course
            grouped = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
            for reg_no, alloc in self.allocations.items():
                grouped[alloc['hall_code']][alloc['department']][alloc['course_code']].append(reg_no)

            # Prepare rows
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
                        regno_range_str = ",\n".join(ranges)

                        count = len(regnos_sorted)
                        grand_total += count
                        hall_summary[hall_code] += count

                        master_data.append([
                            block,
                            hall_name,
                            dept,
                            course_code,
                            regno_range_str,
                            str(count)
                        ])

            # Table
            table = Table(master_data, repeatRows=1)
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
            elements.append(Spacer(1, 20))

            # Summary Section
            elements.append(Paragraph("<b>Summary</b>", styles['Heading2']))
            elements.append(Spacer(1, 6))

            # Department & course totals
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
            elements.append(Paragraph("<b>By Department & Course:</b>", styles['Heading3']))
            elements.append(dept_table)
            elements.append(Spacer(1, 12))

            # Hall totals
            hall_data = [["Hall", "Total Students"]]
            for hall in self.halls:
                hall_code = hall['hall_code']
                if hall_code in hall_summary:
                    hall_data.append([f"{hall['hall_name']} ({hall_code})", str(hall_summary[hall_code])])
            hall_table = Table(hall_data, hAlign="LEFT")
            hall_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ]))
            elements.append(Paragraph("<b>By Hall:</b>", styles['Heading3']))
            elements.append(hall_table)
            elements.append(Spacer(1, 12))

            # Grand total
            elements.append(Paragraph(f"<b>Grand Total Students: {grand_total}</b>", styles['Heading2']))

            doc.build(elements)
            print(f"Master PDF Exported: {output_file}")




# Example Usage
if __name__ == "__main__":
    try:
        allocator = ExamSeatAllocator(students_file='16students.xlsx', halls_file='16halls.xlsx')
        allocator.allocate_seats()
        allocator.print_seating_plan()
        allocator.export_pdf_seating_plan()
        allocator.export_master_seating_plan()

    except Exception as e:
        print(f"Error: {e}")
