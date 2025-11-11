import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors


df = pd.read_csv("Eng 3 Months-Quiz 5-responses.csv")

# Identify question columns
question_cols = [col for col in df.columns if "Response" in col]
answer_cols = [col for col in df.columns if "Right answer" in col]


question_stats = []
for i, q in enumerate(question_cols):
    right_col = answer_cols[i]
    total = len(df)
    correct = sum(df[q] == df[right_col])
    percent = (correct / total) * 100
    question_stats.append({
        "Question": q.replace("Response ", "Q"),
        "Correct %": round(percent, 2),
        "Wrong %": round(100 - percent, 2)
    })

# Sort by lowest correct
hardest = sorted(question_stats, key=lambda x: x["Correct %"])[:3]

# Generate PDF
styles = getSampleStyleSheet()
doc = SimpleDocTemplate("Quiz_Report.pdf", pagesize=A4)
story = []

story.append(Paragraph("<b>Quiz Summary Report</b>", styles['Title']))
story.append(Spacer(1, 12))

story.append(Paragraph(f"Total participants: {len(df)}", styles['Normal']))
story.append(Spacer(1, 12))

story.append(Paragraph("<b>Hardest Questions</b>", styles['Heading2']))
for h in hardest:
    story.append(Paragraph(f"{h['Question']} — {h['Correct %']}% correct", styles['Normal']))

story.append(Spacer(1, 12))
story.append(Paragraph("<b>Question Performance</b>", styles['Heading2']))

table_data = [["Question", "Correct %", "Wrong %"]]
for q in question_stats:
    table_data.append([q["Question"], q["Correct %"], q["Wrong %"]])

t = Table(table_data, hAlign="LEFT")
t.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
]))
story.append(t)

doc.build(story)
print("✅ Generated Quiz_Report.pdf successfully.")
