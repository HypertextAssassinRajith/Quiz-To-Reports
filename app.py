import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
from reportlab.lib.enums import TA_CENTER
from datetime import datetime
from typing import Any, cast, List


# Load the CSV (use the new file provided)
df = pd.read_csv("Eng 3 Months-Quiz 6-responses.csv")

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
    def __init__(self, drawing, width=None, height=None):
        super().__init__()
        self.drawing = drawing
        self.width = width if width is not None else getattr(drawing, 'width', 0)
        self.height = height if height is not None else getattr(drawing, 'height', 0)

    def wrap(self, availWidth, availHeight):
        return (self.width, self.height)

    def draw(self):
        # render the graphics drawing onto the current canvas
        renderPDF.draw(self.drawing, self.canv, 0, 0)
def make_bar(percent: float, width: float = 120.0, height: float = 12.0) -> DrawingFlowable:
    # ensure numeric floats for layout math (accept ints/floats)
    width = float(width)
    width = max(40.0, width)
    d = Drawing(width, height)
    # cast colors to Any so the static type checker accepts them
    d.add(Rect(0, 0, width, height, strokeColor=cast(Any, colors.black), fillColor=cast(Any, colors.lightgrey)))
    fill_w = max(0.0, min(width * (percent / 100.0), width))
    bar_color = colors.green if percent >= 70 else colors.orange if percent >= 40 else colors.red
    d.add(Rect(0, 0, fill_w, height, strokeColor=cast(Any, None), fillColor=cast(Any, bar_color)))
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
doc = SimpleDocTemplate("Quiz_Report.pdf", pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
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
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ('TOPPADDING', (0, 0), (-1, 0), 8),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
]))

story.append(t)

story.append(Spacer(1, 10))
story.append(Paragraph("Generated by Quiz-To-Reports", meta_style))

# Build PDF
doc.build(story)
print("✅ Generated Quiz_Report.pdf successfully.")
