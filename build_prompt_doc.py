from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date

today = date.today().strftime('%B %-d, %Y')  # e.g. "June 28, 2026"

doc = Document()

def heading(text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    return p

def body(text):
    p = doc.add_paragraph(text)
    p.style.font.size = Pt(11)
    return p

def bullet(text):
    doc.add_paragraph(text, style='List Bullet')

# ── Title block ──────────────────────────────────────────────
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('ENDEAVOR HEALTH SYSTEM\nHEALTHCARE EXECUTIVE BRIEFING')
run.bold = True
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.add_run('Weekly Literature Scan  |  Past 7 Days  |  Covers all six Endeavor Health System hospitals\n'
            'Prompt Version 3.0  |  Updated June 2026  |  Works in Claude and ChatGPT').italic = True

doc.add_paragraph()

# ── Instructions ─────────────────────────────────────────────
heading('HOW TO USE THIS PROMPT', level=2)
body('Open a new Claude or ChatGPT conversation. Copy everything inside the box below and paste it '
     'as your first message. The AI will immediately begin researching and building the briefing — '
     'no additional input needed.')

doc.add_paragraph()
doc.add_paragraph('━' * 60)
heading('▼  BEGIN PROMPT — COPY FROM HERE', level=2)
doc.add_paragraph()

# ── THE PROMPT ───────────────────────────────────────────────
prompt_paragraphs = [

("ROLE & PURPOSE",
 "You are a senior healthcare intelligence analyst preparing a comprehensive executive briefing for "
 f"a meeting at Endeavor Health System or one of its hospitals. Today's date is {today}. "
 "Research only the past 7 days from that date. Synthesize everything clearly and produce a single "
 "polished, well-formatted document ready for executive review."),

("HOSPITAL FILTER — CRITICAL INSTRUCTION",
 "The six hospitals and system entities covered by this briefing are:\n"
 "  1. Endeavor Health (the consolidated system)\n"
 "  2. Evanston Hospital\n"
 "  3. Glenbrook Hospital\n"
 "  4. Skokie Hospital\n"
 "  5. Highland Park Hospital\n"
 "  6. Lake Forest Hospital\n\n"
 "FILTER RULE: In Section 2 (News Stories) and Section 3 (Endeavor Deep Dive), include a story ONLY "
 "if it explicitly mentions one or more of the six entities above. Do not include stories about other "
 "health systems unless they directly involve one of the six entities."),

("STEP 1 — CONDUCT LIVE WEB RESEARCH BEFORE WRITING ANYTHING",
 "Search the following sources for content published in the PAST 7 DAYS ONLY. Do not use training "
 "data for current events.\n\n"
 "BECKER'S HOSPITAL REVIEW (beckershospitalreview.com):\n"
 "  • AI & Health IT: /healthcare-information-technology/ai/\n"
 "  • Cybersecurity: /healthcare-information-technology/cybersecurity/\n"
 "  • Finance & Transactions: /finance/ and /hospital-transactions-and-valuation/\n"
 "  • Clinical & Quality: /quality/ and /pharmacy/\n"
 "  • Workforce & HR: /workforce/ and /hr/\n"
 "  • Hospital Management & Capital: /hospital-management-administration/ and /capital/\n"
 "  • Weekly Newsletter Archives: /newsletter-category/beckers-hospital-review-e-weekly/ (last 1–2 issues)\n\n"
 "ADDITIONAL SOURCES (past 7 days):\n"
 "  • Modern Healthcare (modernhealthcare.com)\n"
 "  • Fierce Healthcare (fiercehealthcare.com)\n"
 "  • Advisory Board (advisory.com)\n"
 "  • Kaufman Hall (kaufmanhall.com)\n"
 "  • American Hospital Association (aha.org)\n"
 "  • HIPAA Journal (hipaajournal.com)\n"
 "  • MobiHealthNews, AJMC, Applied Clinical Trials Online\n"
 "  • Medscape, IQVIA — drug pipeline and GLP-1 developments\n"
 "  • KFF / Georgetown CCF — Medicaid policy\n\n"
 "IMPORTANCE FILTER — keep only stories meeting at least TWO of these criteria:\n"
 "  • Affects hospital finances, strategy, or operations at material scale\n"
 "  • Represents a new or measurably accelerating trend\n"
 "  • Would need to be on a CEO, CFO, CMO, or board member's radar\n"
 "  • Published in the past 7 days"),

("STEP 2 — RESEARCH ALL SIX ENDEAVOR HEALTH SYSTEM HOSPITALS SPECIFICALLY",
 "Conduct dedicated searches for each entity for the past 7 days:\n\n"
 "SYSTEM-LEVEL:\n"
 "  • Endeavor Health (Illinois) — system overview, recent news, financial moves, strategic initiatives\n"
 "  • Endeavor Health clinical trials — clinicaltrials.endeavorhealth.org and clinicaltrials.gov\n"
 "  • Endeavor Health innovation and AI\n"
 "  • Endeavor Health community investment — CIF awards, health equity, community partnerships\n\n"
 "HOSPITAL-SPECIFIC (search each individually):\n"
 "  • Evanston Hospital\n"
 "  • Glenbrook Hospital\n"
 "  • Skokie Hospital\n"
 "  • Highland Park Hospital\n"
 "  • Lake Forest Hospital\n\n"
 "LEADERSHIP:\n"
 "  • Gabrielle Cummings-Johnson — current title, recent public statements, media appearances"),

("STEP 3 — TOPIC AREAS TO COVER (do not skip any)",
 "  1. AI in Hospitals & Health Systems\n"
 "  2. Cybersecurity & Data Breaches\n"
 "  3. Hospital M&A & Strategic Consolidation\n"
 "  4. Hospital Finances, Medicaid & Federal Policy\n"
 "  5. New Medications, Treatments & GLP-1s\n"
 "  6. Clinical Trials & FDA Regulatory Shifts\n"
 "  7. Health IT / EHR & Digital Infrastructure\n"
 "  8. Hospital Safety, Workforce Violence & Nursing Shortage\n"
 "  9. Hospital Administration & Organizational Structure\n"
 " 10. Community Relations & Health Equity\n"
 " 11. Endeavor Health System Deep Dive [Special Focus — all six hospitals]"),

("STEP 4 — PRODUCE THE BRIEFING IN THIS EXACT STRUCTURE",
 "Use the four sections below — in this order — with no other sections added.\n\n"
 "─────────────────────────────────────────\n"
 "SECTION 1 — TOPIC BRIEFING\n"
 "─────────────────────────────────────────\n"
 "One row per topic (1–10). Five columns:\n"
 "  (1) Topic Area\n"
 "  (2) What's Going On — 3–5 sentences with key data points and recent developments\n"
 "  (3) Why It Matters to Hospital Executives — 2–3 sentences\n"
 "  (4) Growing Trend? — one-line verdict\n"
 "  (5) Key Source — hyperlinked publication name\n\n"
 "─────────────────────────────────────────\n"
 "SECTION 2 — NEWS STORIES BY TOPIC (HOSPITAL-FILTERED)\n"
 "─────────────────────────────────────────\n"
 "For each of the 10 topic areas:\n"
 "  • Print the topic name as a bold heading\n"
 "  • List 4–6 stories that mention Endeavor Health, Evanston Hospital, Glenbrook Hospital,\n"
 "    Skokie Hospital, Highland Park Hospital, or Lake Forest Hospital\n"
 "  • Story format: [N]. Bold headline → two-sentence summary → Source: hyperlinked pub + date\n"
 "  • If no qualifying stories exist for a topic, write:\n"
 "    'No recent stories found that directly mention an Endeavor Health System hospital.'\n\n"
 "─────────────────────────────────────────\n"
 "SECTION 3 — ENDEAVOR HEALTH SYSTEM DEEP DIVE\n"
 "─────────────────────────────────────────\n"
 "Four columns: (1) Area | (2) Key Findings — 3–5 sentences | (3) Relevance to Your Meeting — 2–3 sentences | (4) Source\n"
 "Cover each of these areas:\n"
 "  • System Overview (Endeavor Health)\n"
 "  • Evanston Hospital\n"
 "  • Glenbrook Hospital\n"
 "  • Skokie Hospital\n"
 "  • Highland Park Hospital\n"
 "  • Lake Forest Hospital\n"
 "  • Leadership Profile (Gabrielle Cummings-Johnson)\n"
 "  • AI & Innovation\n"
 "  • Clinical Trials & Research\n"
 "  • Community Investment\n\n"
 "─────────────────────────────────────────\n"
 "SECTION 4 — INDUSTRY TRENDS WITH ENDEAVOR COVERAGE  ← THIS GOES AT THE END\n"
 "─────────────────────────────────────────\n"
 "Identify 6–8 cross-cutting industry trends found across all your research.\n"
 "For each trend:\n"
 "  • Bold trend name as a subheading\n"
 "  • 2–3 sentences describing what is changing and why it matters\n"
 "  • A sub-list titled 'Endeavor & Hospital Coverage' listing every article from your research\n"
 "    that relates to this trend AND mentions one of the six Endeavor hospitals.\n"
 "    Aim for 3–6 bullet points per trend.\n"
 "    Format each bullet: Hospital Name — Headline — one sentence — Source link\n"
 "  • If no Endeavor-specific articles exist for a trend, write:\n"
 "    'No Endeavor-specific coverage on this trend this week.'\n\n"
 "NOTE: Section 4 replaces the old 'Emerging Trends' table that appeared at the beginning.\n"
 "Do NOT add a standalone trends section anywhere before Section 4."),

("STEP 5 — QUALITY STANDARDS (CHECK BEFORE FINALIZING)",
 "  • Every story in Section 2 must have a real, working, specific source URL\n"
 "  • No topic area may be vague or contain placeholder text\n"
 "  • The Endeavor section must be specific and based on live research\n"
 "  • All four sections must be present and complete\n"
 "  • The date at the top of the document must be today's actual run date — not a hardcoded date\n"
 "  • Section 2 must have 4–6 stories per topic, not just one\n"
 "  • Section 4 must have 3–6 Endeavor/hospital bullets per trend, not just one\n"
 "  • Remove any Section 2 stories that do not name one of the six hospitals"),

("DELIVERY",
 "When the briefing is complete, email it to:\n"
 "  neilsnovich@gmail.com\n"
 "  nndngrp@ameritech.net\n"
 "using the Gmail account shadowext9@gmail.com.\n"
 "Subject: 'Endeavor Health Executive Briefing — [Today's Date]'"),
]

for title_text, content in prompt_paragraphs:
    p = doc.add_paragraph()
    p.add_run(title_text).bold = True
    doc.add_paragraph(content)
    doc.add_paragraph()

doc.add_paragraph('━' * 60)
heading('▲  END PROMPT — STOP COPYING HERE', level=2)

out = '/home/user/Default-Test/Endeavor_Health_Briefing_Prompt_v3.docx'
doc.save(out)
print('Saved:', out)
