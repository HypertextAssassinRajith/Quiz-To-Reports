import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib.enums import TA_CENTER
from datetime import datetime


df = pd.read_csv("Eng 3 Months-Quiz 5-responses.csv")

# Identify question columns
question_cols = [col for col in df.columns if "Response" in col]
answer_cols = [col for col in df.columns if "Right answer" in col]


question_stats = []
for i, q in enumerate(question_cols):
    right_col = answer_cols[i]
    total = len(df)
    correct = sum(df[q] == df[right_col])
    percent = (correct / total) * 100 if total > 0 else 0
    question_stats.append({
        "Question": q.replace("Response ", "Q"),
        "Correct %": round(percent, 2),
        "Wrong %": round(100 - percent, 2)
    })

# Sort by lowest correct
hardest = sorted(question_stats, key=lambda x: x["Correct %"])[:3]

# Small  draw horizontal percentage bars 
def make_bar(percent, width=120, height=12):
    width = max(40, float(width))
    d = Drawing(width, height)
    d.add(Rect(0, 0, width, height, strokeColor=colors.black, fillColor=colors.lightgrey))
    fill_w = max(0, min(width * (percent / 100.0), width))
    bar_color = colors.green if percent >= 70 else colors.orange if percent >= 40 else colors.red
    d.add(Rect(0, 0, fill_w, height, strokeColor=None, fillColor=bar_color))
    padding = 4
    text_x = fill_w + padding if fill_w + 30 < width else max(width - 30, padding)
    text_x = min(max(text_x, padding), width - 4)
    d.add(String(text_x, height / 4, f"{percent}%", fontSize=8, fillColor=colors.black))
    return d

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

# Table with a visual bar column
table_data = [["Question", "Correct %", "Wrong %", "Visual"]]
for q in question_stats:
    pct = q["Correct %"]

    table_data.append([Paragraph(q["Question"], styles['Normal']), Paragraph(f"{pct}%", styles['Normal']), Paragraph(f"{q['Wrong %']}%", styles['Normal']), None])


col_widths = [doc.width * 0.45, doc.width * 0.12, doc.width * 0.12, doc.width * 0.31]

visual_col_width = col_widths[3]
for i in range(1, len(table_data)):
    pct_text = table_data[i][1].getPlainText() if hasattr(table_data[i][1], 'getPlainText') else ''
    try:
        pct_val = float(pct_text.replace('%', ''))
    except:
        pct_val = 0.0
    table_data[i][3] = make_bar(round(pct_val, 2), width=visual_col_width - 8)

t = Table(table_data, colWidths=col_widths, hAlign="LEFT")

t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2b7a78')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('ALIGN', (1, 1), (2, -1), 'CENTER'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f6f6f6')]),
    ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
]))

story.append(t)

story.append(Spacer(1, 10))
story.append(Paragraph("Generated by Quiz-To-Reports", meta_style))

# Build PDF
doc.build(story)
print("✅ Generated Quiz_Report.pdf successfully.")
