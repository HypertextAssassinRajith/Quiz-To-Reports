# A small, friendly script that reads quiz responses and creates a student-friendly PDF report.
# It does three simple things:
# 1) load the CSV export from your quiz platform,
# 2) compute per-question stats (percent correct / wrong), and
# 3) render a neat PDF with question text, options and a small progress bar for each question.
# The goal is to make a report teachers or students can read quickly — no data wrangling required.

import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
from reportlab.lib.enums import TA_CENTER
from datetime import datetime
from typing import Any, cast, List


# Load the CSV containing the quiz responses. Change this filename if you want to use another quiz file. (This used moodle defult export format)
df = pd.read_csv("Eng 3 Months-Quiz 7-responses.csv")

# Detect question indices by presence of Response and Right answer columns
max_questions = 0
for col in df.columns:
    if col.startswith("Response "):
        try:
            idx = int(col.replace("Response ", ""))
            max_questions = max(max_questions, idx)
        except:
            pass

question_stats = []
for i in range(1, max_questions + 1):
    resp_col = f"Response {i}"
    right_col = f"Right answer {i}"
    text_col = f"Question {i}"
    if resp_col not in df.columns or right_col not in df.columns:
        continue

    # Extract a representative question text (first non-empty value)
    q_text = None
    if text_col in df.columns:
        non_empty = df[text_col].dropna().astype(str).str.strip()
        non_empty = non_empty[non_empty.ne('') & non_empty.ne('-')]
        if len(non_empty) > 0:
            q_text = non_empty.iloc[0]
    if not q_text:
        q_text = text_col

    # Try to parse prompt and options from the question text
    prompt = q_text
    options = []
    if ':' in q_text:
        parts = q_text.split(':', 1)
        prompt = parts[0].strip()
        opts_part = parts[1]
        # options are often separated by ';'
        options = [o.strip() for o in opts_part.split(';') if o.strip() and o.strip() != '-']

    # If no options found in the question text, gather unique answers from right/response columns
    if not options:
        vals = pd.concat([df[right_col].dropna().astype(str), df[resp_col].dropna().astype(str)], ignore_index=True)
        vals = vals.str.strip()
        vals = vals[vals.ne('') & vals.ne('-')]
        options = list(dict.fromkeys(vals.tolist()))[:6]  # keep order, limit to reasonable number

    # Representative correct answer (first non-empty from right_col)
    correct_answer = ''
    if right_col in df.columns:
        non_empty_right = df[right_col].dropna().astype(str).str.strip()
        non_empty_right = non_empty_right[non_empty_right.ne('') & non_empty_right.ne('-')]
        if len(non_empty_right) > 0:
            correct_answer = non_empty_right.iloc[0]

    # Consider only rows that actually answered this question
    answered_mask = df[resp_col].notna() & df[resp_col].astype(str).str.strip().ne('') & df[resp_col].astype(str).str.strip().ne('-')
    total_answered = int(answered_mask.sum())
    if total_answered == 0:
        percent = 0.0
        correct = 0
    else:
        correct = int(((df[resp_col] == df[right_col]) & answered_mask).sum())
        percent = (correct / total_answered) * 100

    question_stats.append({
        "Index": i,
        "Question": f"Q{i}",
        "Prompt": prompt,
        "Text": q_text,
        "Options": options,
        "CorrectAnswer": correct_answer,
        "Answered": total_answered,
        "Correct %": round(percent, 2),
        "Wrong %": round(100 - percent, 2)
    })

# Sort by lowest correct
hardest = sorted(question_stats, key=lambda x: x["Correct %"])[:3]

# Small  draw horizontal percentage bars 
class DrawingFlowable(Flowable):
    # DrawingFlowable wraps a ReportLab drawing so we can drop it into a table cell like any other element.
    # Think of it as turning a tiny graphic into a normal object the PDF builder understands.
    def __init__(self, drawing, width=None, height=None):
        super().__init__()
        self.drawing = drawing
        self.width = width if width is not None else getattr(drawing, 'width', 0)
        self.height = height if height is not None else getattr(drawing, 'height', 0)

    def wrap(self, availWidth, availHeight):
        # Tell ReportLab the size we'll take so layout works correctly.
        return (self.width, self.height)

    def draw(self):
        # When the PDF is being rendered, draw the graphic on the canvas.
        renderPDF.draw(self.drawing, self.canv, 0, 0)

# Helper to create a small horizontal progress bar showing percent correct.
# It returns something we can insert straight into a table cell.
def make_bar(percent: float, width: float = 120.0, height: float = 12.0) -> DrawingFlowable:
    # Make sure the bar has a sensible minimum width and compute how much of it to fill.
    width = float(width)
    width = max(40.0, width)
    d = Drawing(width, height)

    # Background track and a colored filled portion based on performance.
    d.add(Rect(0, 0, width, height, strokeColor=cast(Any, colors.black), fillColor=cast(Any, colors.lightgrey)))
    fill_w = max(0.0, min(width * (percent / 100.0), width))
    bar_color = colors.green if percent >= 70 else colors.orange if percent >= 40 else colors.red
    d.add(Rect(0, 0, fill_w, height, strokeColor=cast(Any, None), fillColor=cast(Any, bar_color)))

    # Draw the percent label inside the bar so it never spills outside the cell.
    padding = 4.0
    text_x = fill_w + padding if fill_w + 30.0 < width else max(width - 30.0, padding)
    text_x = min(max(text_x, padding), width - 4.0)
    d.add(String(text_x, height / 4, f"{percent}%", fontSize=8, fillColor=cast(Any, colors.black)))
    return DrawingFlowable(d, width=width, height=height)

# Generate PDF
styles = getSampleStyleSheet()
# custom title style
title_style = ParagraphStyle(
    'CenteredTitle',
    parent=styles['Title'],
    alignment=TA_CENTER,
    fontSize=20,
    spaceAfter=6
)
meta_style = ParagraphStyle('Meta', parent=styles['Normal'], fontSize=9, textColor=colors.grey)

# Add comfortable margins
doc = SimpleDocTemplate("Quiz_Report.pdf", pagesize=landscape(A4), leftMargin=6*mm, rightMargin=6*mm, topMargin=10*mm, bottomMargin=10*mm)
story = []

# Title block
title_tbl = Table(
    [[Paragraph("<b>Quiz Summary Report</b>", title_style)]],
    colWidths=[doc.width]
)
title_tbl.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#0b6e4f')),
    ('TEXTCOLOR', (0, 0), (-1, -1), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ('TOPPADDING', (0, 0), (-1, -1), 8),
]))
story.append(title_tbl)
story.append(Spacer(1, 8))

# Meta info
story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", meta_style))
story.append(Paragraph(f"Total participants: {len(df)}", styles['Normal']))
story.append(Spacer(1, 12))

# Average performance
avg_percent = round(sum(q['Correct %'] for q in question_stats) / len(question_stats), 2) if question_stats else 0
story.append(Paragraph(f"<b>Average score across all questions:</b> {avg_percent}%", styles['Normal']))
story.append(Spacer(1, 12))

# Hardest questions
story.append(Paragraph("<b>Hardest Questions</b>", styles['Heading2']))
for h in hardest:
    color = 'red' if h['Correct %'] < 40 else 'orange' if h['Correct %'] < 70 else 'green'
    story.append(Paragraph(f"{h['Question']} — <font color=\"{color}\">{h['Correct %']}% correct</font>", styles['Normal']))

story.append(Spacer(1, 12))
story.append(Paragraph("<b>Question Performance</b>", styles['Heading2']))

# Prepare table rows including full question text (wrapped)
question_text_style = ParagraphStyle('QuestionText', parent=styles['Normal'], fontSize=9)
header_style = ParagraphStyle('Header', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER, textColor=colors.whitesmoke)

# initialize as a generic list to avoid static type inference as list[str]
table_data = []
# header row as Paragraphs (removed "Correct answer" column)
table_data.append([
    Paragraph("Question", header_style),
    Paragraph("Question text & options", header_style),
    Paragraph("Correct %", header_style),
    Paragraph("Wrong %", header_style),
    Paragraph("Visual", header_style)
])

for q in question_stats:
    # Put the question label and the prompt+options in separate columns so text can wrap nicely
    q_label = Paragraph(f"<b>{q['Question']}</b>", styles['Normal'])

    # Build HTML for prompt and options, highlight the correct answer
    prompt_html = f"<b>{q['Prompt']}</b>"
    opts_html = []
    for opt in q.get('Options', []):
        if q.get('CorrectAnswer') and opt.strip() == q.get('CorrectAnswer').strip():
            opts_html.append(f"<font color=\"#0b6e4f\"><b>&#10003; {opt}</b></font>")
        else:
            opts_html.append(f"• {opt}")
    if opts_html:
        prompt_html += '<br/>' + '<br/>'.join(opts_html)

    q_text_para = Paragraph(prompt_html, question_text_style)
    pct = q["Correct %"]
    pct_para = Paragraph(f"{pct}%", styles['Normal'])
    wrong_para = Paragraph(f"{q['Wrong %']}%", styles['Normal'])
    # now append row without the separate correct-answer column
    table_data.append([q_label, q_text_para, pct_para, wrong_para, None])

# Use proportional column widths so they sum to the document width (removed one column)
col_widths = [doc.width * 0.08, doc.width * 0.60, doc.width * 0.10, doc.width * 0.10, doc.width * 0.12]

# Create visual bars sized to the visual column width
visual_col_width = col_widths[4]
for i in range(1, len(table_data)):
    cell = table_data[i][2]  # Correct % is now column index 2
    # extract numeric percent robustly
    if isinstance(cell, Paragraph):
        try:
            pct_text = cell.getPlainText()
        except Exception:
            pct_text = str(cell)
    else:
        pct_text = str(cell)
    try:
        pct_val = float(pct_text.replace('%', '').strip())
    except Exception:
        pct_val = 0.0
    # create bar slightly smaller than column to leave cell padding
    table_data[i][4] = make_bar(round(pct_val, 2), width=visual_col_width - 12)

# Build the Table
t = Table(table_data, colWidths=col_widths, hAlign="LEFT")

# Improved table styling: clear borders and center visual column (update indices)
t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2b7a78')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    # center numeric columns (Correct %, Wrong %) and the visual column
    ('ALIGN', (2, 1), (3, -1), 'CENTER'),
    ('ALIGN', (4, 1), (4, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('BOX', (0, 0), (-1, -1), 0.8, colors.grey),
    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f6f6f6')]),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ('TOPPADDING', (0, 0), (-1, 0), 8),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
]))

story.append(t)

# Add a small gap and then a student results table so teachers can see individual marks
story.append(Spacer(1, 12))
story.append(Paragraph("<b>Student Results</b>", styles['Heading2']))

# Calculate how many students completed the quiz out of the total enrolled (excluding system accounts)
skip_usernames = { 'admin', 'rajithsanjaya' }
# Count unique enrolled usernames (exclude system accounts)
usernames_series = df.get('Username', pd.Series(dtype=str)).dropna().astype(str).str.strip()
valid_usernames = [u for u in usernames_series if u and u.lower() not in skip_usernames]
total_enrolled = len(set([u.lower() for u in valid_usernames]))

# Prepare a DataFrame that contains only the highest-grade attempt per student
df_users = df.copy()
# Ensure we have a username series we can manipulate safely
if 'Username' in df_users.columns:
    username_series = df_users['Username'].astype(str).str.strip()
else:
    username_series = pd.Series([''] * len(df_users))

# normalized lowercase username for grouping
df_users['Username_lc'] = username_series.str.lower()

# safe numeric grade column; missing -> NaN then fill with -1 so missing attempts sort last
grade_col = 'Grade/100.00'
if grade_col in df_users.columns:
    grade_series = pd.to_numeric(df_users[grade_col], errors='coerce')
else:
    grade_series = pd.Series([float('nan')] * len(df_users))

df_users['grade_val'] = grade_series.fillna(-1)

# consider only rows with a username and not in skip list
df_sel = df_users[(df_users['Username_lc'] != '') & (~df_users['Username_lc'].isin(skip_usernames))]

# For each username pick the row with the highest grade_val
if not df_sel.empty:
    best_idx = df_sel.groupby('Username_lc')['grade_val'].idxmax()
    df_best = df_sel.loc[best_idx].copy()
else:
    df_best = df_sel.copy()

# Count how many unique students completed based on the best attempt per student
completed_usernames = set()
for _, row in df_best.iterrows():
    username = str(row.get('Username', '')).strip()
    status = str(row.get('Status', '')).strip().lower()
    grade_val = row.get('grade_val', -1)
    if status == 'finished' or (isinstance(grade_val, (int, float)) and grade_val >= 0):
        completed_usernames.add(username.lower())

completed = len(completed_usernames)
percent_done = round((completed / total_enrolled) * 100, 1) if total_enrolled > 0 else 0
story.append(Paragraph(f"Students completed: {completed} / {total_enrolled} ({percent_done}%)", styles['Normal']))
story.append(Spacer(1, 8))

# Build a list of students using only the best (highest-grade) attempt per username
students = []
for _, row in df_best.iterrows():
    last = str(row.get('Last name', '')).strip()
    first = str(row.get('First name', '')).strip()
    name = f"{first} {last}".strip() or str(row.get('Username', '')).strip()
    username = str(row.get('Username', '')).strip()
    email = str(row.get('Email address', '')).strip()
    status = str(row.get('Status', '')).strip()
    duration = str(row.get('Duration', '')).strip()
    grade_val = row.get('grade_val', None)
    try:
        grade = None if grade_val is None or (isinstance(grade_val, (int, float)) and grade_val < 0) else float(grade_val)
    except Exception:
        grade = None
    students.append({
        'Name': name,
        'Username': username,
        'Email': email,
        'Grade': grade,
        'Status': status,
        'Duration': duration
    })

# Sort students by grade (highest first), keep those without grade at bottom
students = sorted(students, key=lambda s: (s['Grade'] is not None, s['Grade'] if s['Grade'] is not None else -1), reverse=True)

# Prepare table data for students
student_header_style = ParagraphStyle('StudentHeader', parent=styles['Normal'], fontSize=9, alignment=TA_CENTER, textColor=colors.whitesmoke)
student_small = ParagraphStyle('StudentSmall', parent=styles['Normal'], fontSize=8)

student_table = []
student_table.append([
    Paragraph('Name', student_header_style),
    Paragraph('Username', student_header_style),
    Paragraph('Email', student_header_style),
    Paragraph('Grade', student_header_style),
    Paragraph('Status', student_header_style),
    Paragraph('Duration', student_header_style),
])

# Track which rows are 'not attempted' so we can highlight them
not_attempted_rows = []
for idx, s in enumerate(students):
    name_para = Paragraph(s['Name'], student_small)
    username_para = Paragraph(s['Username'], student_small)
    email_para = Paragraph(s['Email'], student_small)
    grade_text = f"{s['Grade']:.2f}" if s['Grade'] is not None else '-'
    grade_para = Paragraph(grade_text, student_small)
    status_text = s['Status'] or '-'
    status_para = Paragraph(status_text, student_small)
    duration_para = Paragraph(s['Duration'] or '-', student_small)
    student_table.append([name_para, username_para, email_para, grade_para, status_para, duration_para])

    # Consider a student 'not attempted' if they don't have a numeric grade or didn't finish
    status_norm = str(s['Status']).strip().lower() if s['Status'] else ''
    if s['Grade'] is None or status_norm != 'finished':
        # table header occupies row 0, so first student row is 1 -> add 1 to idx
        not_attempted_rows.append(1 + idx)

# Column widths that fit the page nicely
student_col_widths = [doc.width * 0.28, doc.width * 0.12, doc.width * 0.30, doc.width * 0.08, doc.width * 0.12, doc.width * 0.10]

stu_t = Table(student_table, colWidths=student_col_widths, hAlign='LEFT')

# Base table style
style_cmds = [
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2b7a78')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('ALIGN', (3, 1), (3, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('BOX', (0, 0), (-1, -1), 0.6, colors.grey),
    ('INNERGRID', (0, 0), (-1, -1), 0.3, colors.grey),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#fbfbfb')]),
    ('FONTSIZE', (0, 0), (-1, -1), 8),
    ('LEFTPADDING', (0, 0), (-1, -1), 4),
    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
]

# Highlight rows where students did not attempt the quiz (light yellow background and muted text)
for r in not_attempted_rows:
    style_cmds.append(('BACKGROUND', (0, r), (-1, r), colors.HexColor('#fff3cd')))
    style_cmds.append(('TEXTCOLOR', (0, r), (-1, r), colors.HexColor('#6c2d00')))

stu_t.setStyle(TableStyle(style_cmds))

story.append(Spacer(1, 8))
story.append(stu_t)

# Footer and build the document
story.append(Spacer(1, 10))
story.append(Paragraph("Generated by Quiz-To-Reports", meta_style))

# Build PDF
doc.build(story)
print("✅ Generated Quiz_Report.pdf successfully.")
