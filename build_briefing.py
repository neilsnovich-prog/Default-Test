"""
Healthcare Executive Briefing Generator
Produces a landscape .docx with 4 sections:
  1. Topic Briefing (10 topics)
  2. Emerging Trends (8 trends)
  3. News Stories by Topic (≤5 stories each)
  4. Endeavor Health Deep Dive
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy
from lxml import etree

# ──────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────

NAVY  = RGBColor(0x1F, 0x4E, 0x79)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BLUE = RGBColor(0xD6, 0xE4, 0xF0)
BLACK = RGBColor(0x00, 0x00, 0x00)


def set_cell_bg(cell, hex_color):
    """Fill a table cell background with a hex color string like '1F4E79'."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_font(cell, bold=False, color=None, size_pt=9, font_name='Arial'):
    for para in cell.paragraphs:
        for run in para.runs:
            run.font.name = font_name
            run.font.size = Pt(size_pt)
            run.font.bold = bold
            if color:
                run.font.color.rgb = color


def add_hyperlink(paragraph, url, text, bold=False, font_size=9):
    """Add a hyperlink to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    # Style: blue underline
    color_el = OxmlElement('w:color')
    color_el.set(qn('w:val'), '0563C1')
    u_el = OxmlElement('w:u')
    u_el.set(qn('w:val'), 'single')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size * 2)))
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(font_size * 2)))
    if bold:
        b = OxmlElement('w:b')
        rPr.append(b)
    rPr.append(rFonts)
    rPr.append(color_el)
    rPr.append(u_el)
    rPr.append(sz)
    rPr.append(szCs)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def cell_para(cell, text='', bold=False, color=None, size=9, align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
    """Return first paragraph of cell after setting its text and formatting."""
    para = cell.paragraphs[0]
    para.alignment = align
    if text:
        run = para.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        if color:
            run.font.color.rgb = color
    return para


def header_row(table, *headers, col_widths=None):
    """Set table's first row as navy header."""
    row = table.rows[0]
    for i, (cell, hdr) in enumerate(zip(row.cells, headers)):
        set_cell_bg(cell, '1F4E79')
        cell_para(cell, hdr, bold=True, color=WHITE, size=9)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return row


def body_row(table, values, row_idx, hyperlinks=None):
    """Add a body row to a table with alternating shading."""
    row = table.add_row()
    fill = 'FFFFFF' if row_idx % 2 == 0 else 'D6E4F0'
    for i, (cell, val) in enumerate(zip(row.cells, values)):
        set_cell_bg(cell, fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        if hyperlinks and i == len(values) - 1 and hyperlinks[row_idx - 1]:
            url, link_text = hyperlinks[row_idx - 1]
            add_hyperlink(cell.paragraphs[0], url, link_text, font_size=9)
        else:
            cell_para(cell, val, size=9)
    return row


def section_heading(doc, text):
    """Add a bold navy section heading paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = NAVY
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    return p


def make_table(doc, col_count, col_widths_in):
    """Create a table with specified column count and widths (inches list)."""
    tbl = doc.add_table(rows=1, cols=col_count)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, w in enumerate(col_widths_in):
        for cell in tbl.columns[i].cells:
            cell.width = Inches(w)
    return tbl


def add_footer(doc, text):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(7)
    run.font.italic = True
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


def set_landscape(doc):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)


# ──────────────────────────────────────────────────────
# RESEARCH DATA
# ──────────────────────────────────────────────────────

TOPIC_BRIEFING = [
    {
        "topic": "1. AI in Hospitals & Health Systems",
        "whats_going_on": (
            "AI has moved from pilot programs to enterprise-scale infrastructure across U.S. health systems. "
            "Epic reports 300+ health systems using at least one AI tool in production as of early 2026, "
            "covering clinical documentation, sepsis detection, radiology triage, and capacity management. "
            "Abridge's ambient documentation platform is now licensed at 1,000+ clinician seats at Endeavor Health alone. "
            "CIOs are consolidating point solutions into integrated AI platforms, reducing vendor sprawl. "
            "Advanced clinical decision support is evolving into real-time assistants aware of the patient chart, payer rules, and the live clinical conversation."
        ),
        "why_matters": (
            "AI governance has become a C-suite and board-level priority demanding clear ROI frameworks and enterprise policies. "
            "Vendor consolidation reduces complexity but increases strategic dependency risk. "
            "Revenue cycle AI and ambient documentation can materially improve margins and reduce physician burnout."
        ),
        "trend": "Rapidly accelerating — AI is now clinical infrastructure, not a project",
        "source_text": "Becker's Hospital Review",
        "source_url": "https://www.beckershospitalreview.com/healthcare-information-technology/ai/",
    },
    {
        "topic": "2. Cybersecurity & Data Breaches",
        "whats_going_on": (
            "252 large healthcare data breaches were reported to OCR from January–April 2026, on pace with record years. "
            "The largest incident: NYC Health + Hospitals suffered a breach affecting 1.8 million individuals — exposing medical records, "
            "government IDs, geolocation data, and biometric fingerprint/palm data via a compromised third-party vendor. "
            "Healthcare data breach cost averages $7.42 million per incident — highest of any industry for the 14th consecutive year. "
            "Over 80% of stolen patient records come from third-party vendors, not hospitals directly. "
            "Ransomware remains the primary attack vector, exploiting complex hybrid IT environments."
        ),
        "why_matters": (
            "Financial and reputational exposure from breaches now rivals major capital investments. "
            "Third-party vendor risk management is a board-level governance requirement. "
            "Regulators are increasing enforcement, and class-action litigation following breaches is growing."
        ),
        "trend": "Persistent and growing — breach volume and cost both increasing year over year",
        "source_text": "HIPAA Journal / TechCrunch",
        "source_url": "https://www.hipaajournal.com/healthcare-data-breach-statistics/",
    },
    {
        "topic": "3. Hospital M&A & Strategic Consolidation",
        "whats_going_on": (
            "Q1 2026 recorded 22 announced transactions totaling $14.5 billion in transacted revenue — "
            "the strongest Q1 M&A result in at least five years, per Kaufman Hall. "
            "Three 'mega mergers' were announced: Sutter Health (CA) with Allina Health (MN); "
            "UPMC with CommonSpirit Health (expected close fall 2026); and Atrium Health with WakeMed, "
            "paired with a $2 billion investment in Wake County facilities. "
            "Systems are repositioning by shedding underperforming markets and acquiring new capabilities — "
            "AI platforms, genomics, ambulatory scale. Private equity M&A is also projected to rise through 2026."
        ),
        "why_matters": (
            "Scale is increasingly necessary to compete on AI capability, capital access, and payer negotiating leverage. "
            "Systems not actively growing face margin erosion and closure vulnerability, particularly in Medicaid-heavy markets. "
            "M&A creates integration risk but also opens partnerships in research, technology, and clinical service lines."
        ),
        "trend": "Accelerating — 2026 on track for strongest annual M&A volume in years",
        "source_text": "Kaufman Hall Q1 2026 M&A Report",
        "source_url": "https://www.kaufmanhall.com/insights/research-report/ma-quarterly-activity-report-q1-2026",
    },
    {
        "topic": "4. Hospital Finances, Medicaid & Federal Policy",
        "whats_going_on": (
            "The One Big Beautiful Bill Act (OBBBA), signed July 4, 2025, cuts $911 billion in federal Medicaid spending over 10 years. "
            "Since passage, 800+ hospitals, nursing homes, and health facilities have closed, cut services, or are at risk. "
            "CMS closed a Medicaid MCO tax financing loophole effective April 3, 2026, saving $24+ billion annually. "
            "New Medicaid work requirements are now in effect, reducing enrollment particularly among low-income adults. "
            "A $50 billion Rural Health Fund was created, but analysts estimate it covers only 37% of projected rural Medicaid losses. "
            "The 2026 Medicare payment final rule delayed a major hospital payment cut, but it remains on the horizon."
        ),
        "why_matters": (
            "Hospitals with high Medicaid payer mix — including safety-net and academic medical centers — face direct revenue threats. "
            "For Illinois-based systems, state-level Medicaid responses to federal cuts will determine the financial floor. "
            "CEO and CFO teams must model multi-year scenarios now given the 10-year policy horizon."
        ),
        "trend": "Worsening — near-crisis conditions for safety-net, rural, and high-Medicaid-volume hospitals",
        "source_text": "American Medical Association / KFF",
        "source_url": "https://www.ama-assn.org/health-care-advocacy/federal-advocacy/changes-medicaid-aca-and-other-key-provisions-one-big",
    },
    {
        "topic": "5. New Medications, Treatments & GLP-1s",
        "whats_going_on": (
            "2026 is the year of oral GLP-1 therapy: FDA approved oral semaglutide (Wegovy 25mg) — the first GLP-1 pill for weight loss — "
            "and Eli Lilly's orforglipron (Foundayo), the first small-molecule GLP-1 receptor agonist. "
            "Novo Nordisk launched Wegovy HD (7.2mg injectable) in the U.S. in April 2026. "
            "Medicare's GLP-1 Bridge Program begins July 2026, covering weight-loss medications for qualifying Medicare patients. "
            "Tirzepatide gained FDA approval for sleep apnea; retatrutide shows 28.7% average weight loss in knee osteoarthritis trials. "
            "Semaglutide (Wegovy) received FDA approval for MASH (metabolic dysfunction-associated steatohepatitis) in August 2025."
        ),
        "why_matters": (
            "GLP-1 expansion will reduce bariatric surgical volumes long-term while opening new clinical service lines in cardiology, hepatology, and pulmonology. "
            "Formulary decisions and patient access programs are becoming strategically important for pharmacy and managed care teams. "
            "Medicare GLP-1 coverage represents both a cost and a care quality opportunity for Medicare Advantage plans and ACOs."
        ),
        "trend": "Rapidly accelerating — oral formulations and new indications are reshaping the entire obesity treatment paradigm",
        "source_text": "AJMC / IQVIA",
        "source_url": "https://www.ajmc.com/view/fda-approves-oral-semaglutide-as-first-glp-1-pill-for-weight-loss",
    },
    {
        "topic": "6. Clinical Trials & FDA Regulatory Shifts",
        "whats_going_on": (
            "The FDA launched its Real-Time Clinical Trials (RTCT) initiative in April 2026, with two proof-of-concept trials: "
            "AstraZeneca's mantle cell lymphoma study and Amgen's small cell lung cancer trial. "
            "The RTCT model uses Paradigm Health's technology to stream safety and efficacy data to FDA scientists in real time. "
            "A separate AI-enabled pilot program for early-phase clinical trial optimization was announced via Federal Register (April 2026). "
            "On January 12, 2026, FDA released draft guidance formally accepting Bayesian methodology in pivotal confirmatory trials. "
            "FDA reminded 2,200+ clinical trial sponsors and researchers to disclose results to ClinicalTrials.gov, increasing compliance pressure."
        ),
        "why_matters": (
            "Real-time FDA data review could compress drug development timelines by months or years, creating competitive advantage for research-active health systems. "
            "Bayesian trial methodology allows more adaptive, efficient trial designs — benefiting systems conducting complex oncology or rare disease studies. "
            "AI-assisted trial design reduces cost barriers to academic and community hospital research participation."
        ),
        "trend": "Emerging — major regulatory modernization underway that will reshape research timelines and competitive positioning",
        "source_text": "FDA / HHS",
        "source_url": "https://www.fda.gov/news-events/press-announcements/fda-announces-major-steps-implement-real-time-clinical-trials",
    },
    {
        "topic": "7. Health IT / EHR & Digital Infrastructure",
        "whats_going_on": (
            "Epic now holds 43.7% of acute care hospital market share (56.9% of beds), gaining 77 hospitals in 2025 alone while Oracle Health shed 56. "
            "300+ health systems are using Epic AI tools in production, including ambient documentation with Nuance DAX Copilot and Abridge. "
            "Oracle Health's new AI-enabled EHR is entering production in 2026 — a make-or-break moment: 30% of Millennium customers are not in long-term plans, "
            "35% are considered 'vulnerable,' and Oracle ranked lowest in KLAS Best in KLAS 2026 across all acute care EHR size tiers. "
            "Door County Medical Center and OhioHealth Morrow County Hospital both launched Epic in early 2026, as Epic's community hospital market penetration grows."
        ),
        "why_matters": (
            "EHR platform selection is now a 10–20-year strategic commitment with major AI capability implications. "
            "Systems on Oracle/Cerner face growing platform risk and must evaluate migration costs and timelines now. "
            "Epic's AI integration advantage is widening — creating a compounding competitive gap for late adopters."
        ),
        "trend": "Consolidating rapidly around Epic; Oracle Health in strategic uncertainty with 2026 as pivotal year",
        "source_text": "HealthSystemCIO / KLAS",
        "source_url": "https://healthsystemcio.com/2026/05/14/acute-care-ehr-market-share-2026/",
    },
    {
        "topic": "8. Hospital Safety, Workforce Violence & Nursing Shortage",
        "whats_going_on": (
            "138,000 nurses have left the workforce since 2022; 45% of remaining nurses indicate they are likely to leave their jobs within 12 months. "
            "Workplace violence is a primary driver: 54% of healthcare workers have felt threatened by patients or visitors on the job. "
            "Healthcare workers account for nearly three-quarters of all nonfatal workplace violence injuries requiring days away from work. "
            "Federal legislation protecting healthcare workers from violence has passed the House twice but remains stalled in the Senate. "
            "In Illinois, nurses' unions sued Prime Healthcare and Ascension over alleged understaffing (June 2026). "
            "Violence, staffing shortages, and stress create a vicious cycle accelerating departure from the profession."
        ),
        "why_matters": (
            "Nursing shortages directly drive quality risk, patient safety incidents, longer lengths of stay, and elevated agency/contract labor costs. "
            "Violence exposure is accelerating the departure cycle — making safety investments not just ethical but financially necessary. "
            "Legislative inaction increases pressure on health systems to implement independent safety programs."
        ),
        "trend": "Worsening — a compounding workforce crisis with no near-term policy solution visible",
        "source_text": "Nurse.org / Chief Healthcare Executive",
        "source_url": "https://nurse.org/articles/state-of-nursing-survey-2026/",
    },
    {
        "topic": "9. Hospital Administration & Organizational Structure",
        "whats_going_on": (
            "Health systems are undergoing structural transformation driven by two parallel forces: cost pressure and AI enablement. "
            "The 'Great Flattening' — elimination of administrative layers — is accelerating: Care New England cut 30+ leadership roles to close a $20M budget gap. "
            "Simultaneously, technology leadership is being elevated: Sanford Health created a new EVP/Chief Transformation Officer role (Jan 2026); "
            "NYC Health + Hospitals named a VP/Chief Data and AI Officer. "
            "Ambulatory care is reshaping org charts — BayCare elevated ambulatory leadership to its executive cabinet alongside system strategy. "
            "66 large health systems were ranked by long-term debt in June 2026, signaling growing focus on balance sheet discipline."
        ),
        "why_matters": (
            "Admin cost reduction is near-term necessity, but poorly managed restructuring damages retention and morale among high-performing leaders. "
            "New digital leadership roles reflect AI's shift from optional to infrastructure — systems without clear AI governance structures are falling behind. "
            "Ambulatory integration at the executive level signals where future growth and margin will be generated."
        ),
        "trend": "Accelerating — dual structural shifts (flattening + tech elevation) underway simultaneously",
        "source_text": "AHA / Becker's Hospital Review",
        "source_url": "https://www.aha.org/aha-center-health-innovation-market-scan/2026-02-03-c-suite-rewrites-org-chart-how-health-systems-are-elevating-technology",
    },
    {
        "topic": "10. Community Relations & Health Equity",
        "whats_going_on": (
            "Health equity has moved from a compliance function to a strategic imperative, driven by IRS community benefit requirements, "
            "new U.S. News hospital rankings criteria, and CMS's Framework for Health Equity. "
            "The 2026 SDOH & Health Equity Forum convened in Chicago on June 25–26. "
            "OBBBA Medicaid work requirements and coverage cuts disproportionately threaten low-income and minority communities — "
            "placing hospitals at the intersection of policy advocacy and community mission. "
            "Community benefit investment is increasingly tied to specific SDOH targets: food security, housing, workforce development, and behavioral health."
        ),
        "why_matters": (
            "Community investment is now evaluated by external ratings and IRS auditors — inadequate documentation creates legal and reputational exposure. "
            "Systems that proactively demonstrate health equity outcomes gain preferred status in CMS value-based contracts and state Medicaid agreements. "
            "Medicaid enrollment declines from OBBBA will increase uncompensated care burdens that community benefit programs must offset."
        ),
        "trend": "Growing strategic importance — simultaneously threatened by Medicaid cuts reducing the communities hospitals serve",
        "source_text": "CMS / KFF",
        "source_url": "https://www.kff.org/medicaid/medicaid-what-to-watch-in-2026/",
    },
]

EMERGING_TRENDS = [
    {
        "name": "AI as Core Clinical Infrastructure",
        "what_changing": (
            "Health systems have crossed the inflection point from AI pilots to enterprise deployment. "
            "Ambient documentation, predictive sepsis/deterioration alerts, radiology AI, and revenue cycle automation are now running as infrastructure — "
            "not projects. CIOs are consolidating vendors into platform relationships."
        ),
        "implication": (
            "Systems without enterprise AI governance frameworks risk chaotic adoption and missed ROI. "
            "Vendor consolidation will intensify — fewer, larger AI relationships will define competitive positioning through 2028. "
            "AI fluency is rapidly becoming a required competency for clinical and administrative leadership."
        ),
        "timeframe": "Now–2027",
        "source_text": "Becker's Hospital Review / Epic",
        "source_url": "https://www.beckershospitalreview.com/healthcare-information-technology/ai/",
    },
    {
        "name": "Medicaid Coverage Erosion & Safety Net Stress",
        "what_changing": (
            "The One Big Beautiful Bill Act's $911B in Medicaid cuts, work requirements, and provider tax restrictions "
            "are reducing Medicaid enrollment and shrinking hospital reimbursement — with 800+ facilities already impacted. "
            "States face budget pressure to cut provider rates or limit benefits."
        ),
        "implication": (
            "Safety-net hospitals, rural facilities, and high-Medicaid-mix urban systems face 10-year revenue compression. "
            "Care delivery model redesign is urgent — ambulatory expansion and virtual care can partially offset inpatient revenue losses. "
            "Community advocacy and state-level Medicaid waiver strategy are becoming CFO-level priorities."
        ),
        "timeframe": "2026–2035 (10-year horizon per OBBBA)",
        "source_text": "KFF / AMA / RAND",
        "source_url": "https://www.kff.org/medicaid/medicaid-what-to-watch-in-2026/",
    },
    {
        "name": "GLP-1 Revolution Reshaping Surgical & Specialty Volumes",
        "what_changing": (
            "Oral GLP-1 medications, expanded Medicare coverage, and new indications (sleep apnea, MASH, MACE reduction) "
            "are creating a new treatment paradigm for obesity and cardiometabolic disease. "
            "2026 is the year oral formulations reach mass market scale."
        ),
        "implication": (
            "Bariatric surgical volumes will decline materially over 3–7 years. "
            "New service lines in metabolic medicine, GLP-1 management clinics, and MASH hepatology will be growth vectors. "
            "Formulary decisions, patient stratification tools, and side-effect management programs are becoming strategic differentiators."
        ),
        "timeframe": "Now–2030",
        "source_text": "IQVIA / AJMC / Medicare Rights Center",
        "source_url": "https://www.iqvia.com/locations/emea/blogs/2026/01/outlook-for-obesity-in-2026",
    },
    {
        "name": "EHR Platform Bifurcation: Epic Dominant, Oracle in Crisis",
        "what_changing": (
            "Epic has gained a net 568 hospitals over five years while Oracle Health has lost 173. "
            "2026 is Oracle's critical test year as its AI-enabled EHR enters production — "
            "but 30% of Millennium customers are already not planning a long-term relationship."
        ),
        "implication": (
            "Health systems on Oracle/Cerner must evaluate platform migration costs now vs. 3 years from now — delay compounds the decision. "
            "Epic's AI integration advantage will widen the competitive gap further by 2028. "
            "EHR selection is a 15-year decision with material AI capability implications that executives must own, not delegate."
        ),
        "timeframe": "2026–2030",
        "source_text": "HealthSystemCIO / EHR Source / KLAS",
        "source_url": "https://healthsystemcio.com/2026/05/14/acute-care-ehr-market-share-2026/",
    },
    {
        "name": "Real-Time FDA Review & AI-Accelerated Clinical Research",
        "what_changing": (
            "FDA's Real-Time Clinical Trials initiative streams safety and efficacy data to regulators during the trial — "
            "not after. Combined with AI-enabled early-phase optimization and Bayesian trial design guidance, "
            "the full drug development model is being restructured."
        ),
        "implication": (
            "Research-active health systems with strong informatics infrastructure can become preferred trial sites for RTCT studies, "
            "attracting pharmaceutical partner investment and advanced therapy access. "
            "Community hospital research programs can now access AI-assisted trial design tools that previously required academic center resources."
        ),
        "timeframe": "2026–2030",
        "source_text": "FDA / HHS",
        "source_url": "https://www.fda.gov/news-events/press-announcements/fda-announces-major-steps-implement-real-time-clinical-trials",
    },
    {
        "name": "Cybersecurity Governance Moves to the Board",
        "what_changing": (
            "Third-party vendor breaches now account for 80%+ of stolen healthcare records; "
            "average breach cost of $7.42M and growing class-action litigation are forcing board-level risk committees. "
            "Biometric data theft (NYC Health+Hospitals) has raised the severity classification of breaches."
        ),
        "implication": (
            "Boards require quarterly cybersecurity reporting; vendor contracts must include breach liability provisions. "
            "Health systems must implement zero-trust architectures and annual third-party vendor security audits. "
            "Cyber insurance costs will rise further — self-insured systems face growing reserve requirements."
        ),
        "timeframe": "Now–2027",
        "source_text": "HIPAA Journal / SecurityWeek",
        "source_url": "https://www.hipaajournal.com/healthcare-data-breach-statistics/",
    },
    {
        "name": "Workforce Redesign Around AI Augmentation",
        "what_changing": (
            "Ambient documentation, AI scheduling, and automated clinical decision support are partially offsetting nursing shortages — "
            "but 138,000 nurses have already left since 2022 and 45% of current nurses may follow. "
            "Workplace violence remains unaddressed by federal legislation, accelerating departure."
        ),
        "implication": (
            "Systems that invest in AI augmentation for remaining nurses reduce per-nurse workload and improve retention economics. "
            "Violence prevention investment (de-escalation programs, visitor management, panic systems) is now a retention tool, not just a safety tool. "
            "Workforce planning must account for a structurally smaller nursing pipeline through at least 2030."
        ),
        "timeframe": "2026–2030",
        "source_text": "Nurse.org / AHA",
        "source_url": "https://nurse.org/articles/state-of-nursing-survey-2026/",
    },
    {
        "name": "The Great Flattening of Hospital Administration",
        "what_changing": (
            "Health systems are simultaneously eliminating administrative layers (Great Flattening) while creating new senior AI/digital roles. "
            "Ambulatory care is being elevated to executive cabinet level as outpatient volumes drive system growth and financial sustainability."
        ),
        "implication": (
            "Mid-level administrative talent faces displacement — health systems must manage morale and institutional knowledge loss carefully. "
            "New digital leadership roles require hiring or developing people with both clinical operations expertise and technology fluency. "
            "Ambulatory care leadership will increasingly drive system-level strategy, requiring new org chart designs."
        ),
        "timeframe": "Now–2027",
        "source_text": "AHA / HealthLeaders / Becker's Hospital Review",
        "source_url": "https://www.aha.org/aha-center-health-innovation-market-scan/2026-02-03-c-suite-rewrites-org-chart-how-health-systems-are-elevating-technology",
    },
]

NEWS_STORIES = {
    "1. AI in Hospitals & Health Systems": [
        {
            "headline": "Endeavor Health Scales Abridge Ambient Documentation to 1,000 Clinician Licenses",
            "summary": "Endeavor Health purchased 1,000 Abridge ambient documentation licenses with 250+ clinicians already actively using the tool, aiming to reduce documentation burden and cognitive load. The program reflects a deliberate strategy to drive AI adoption among frontline clinicians through tool access and peer-led change management.",
            "source_text": "Becker's Hospital Review",
            "source_url": "https://www.beckershospitalreview.com/healthcare-information-technology/ai/how-endeavor-health-is-driving-ai-adoption-among-clinicians/",
        },
        {
            "headline": "Epic Reports 300+ Health Systems Using AI Tools in Production",
            "summary": "Epic announced that over 300 health systems are using at least one of its AI-powered tools in production environments, spanning clinical documentation, sepsis prediction, patient deterioration, hospital readmission risk, and no-show forecasting. The milestone marks AI's transition from optional add-on to core platform capability.",
            "source_text": "Becker's Hospital Review, 2026",
            "source_url": "https://www.beckershospitalreview.com/healthcare-information-technology/ai/",
        },
        {
            "headline": "Sanford Health Creates Inaugural Chief Transformation Officer Role for AI/Digital Strategy",
            "summary": "Sanford Health established the first EVP and Chief Transformation Officer position in January 2026, tasked with overseeing enterprise AI, data analytics, innovation, and digital strategy. The move reflects a nationwide pattern of health systems elevating technology leadership to the executive cabinet.",
            "source_text": "AHA Center for Health Innovation, Feb 2026",
            "source_url": "https://www.aha.org/aha-center-health-innovation-market-scan/2026-02-03-c-suite-rewrites-org-chart-how-health-systems-are-elevating-technology",
        },
        {
            "headline": "FDA Launches AI-Enabled Early-Phase Clinical Trial Optimization Pilot",
            "summary": "The FDA published a Federal Register Request for Information in April 2026 for a pilot program exploring how AI and data science can improve clinical trial efficiency, safety monitoring, and dose-selection decisions. The initiative signals regulatory openness to AI-accelerated research designs that were previously unsupported by formal FDA guidance.",
            "source_text": "Federal Register, Apr 29, 2026",
            "source_url": "https://www.federalregister.gov/documents/2026/04/29/2026-08281/ai-enabled-optimization-of-early-phase-clinical-trials-pilot-program-request-for-information",
        },
        {
            "headline": "Health Systems Consolidating AI Vendors: From Point Solutions to Platforms",
            "summary": "A 2026 survey of health system digital leaders found that the priority for AI in 2026 is consolidation — fewer vendors, deeper platform integration, and enterprise-scale value. The era of one-off AI tools is ending as CIOs prioritize reliability, scalability, and tools that clinicians encounter embedded in their existing workflows.",
            "source_text": "Becker's Hospital Review, 2026",
            "source_url": "https://www.beckershospitalreview.com/healthcare-information-technology/ai/all-things-ai-what-health-system-digital-leaders-are-watching-in-26/",
        },
    ],
    "2. Cybersecurity & Data Breaches": [
        {
            "headline": "NYC Health + Hospitals Breach Exposes 1.8M Records Including Fingerprints and Biometrics",
            "summary": "Hackers accessed NYC Health + Hospitals' network from November 2025 through February 2026 via a third-party vendor, stealing medical records, government IDs, geolocation data, and biometric fingerprint and palm-print data for at least 1.8 million individuals. The incident is among the largest healthcare-specific biometric data breaches on record.",
            "source_text": "TechCrunch, May 18, 2026",
            "source_url": "https://techcrunch.com/2026/05/18/nyc-health-and-hospitals-says-hackers-stole-medical-data-and-fingerprints-during-breach-affecting-at-least-1-8-million-people/",
        },
        {
            "headline": "252 Large Healthcare Breaches Reported to OCR in First Four Months of 2026",
            "summary": "HIPAA Journal's tracking shows 252 large healthcare data breaches (500+ individuals) reported to OCR from January–April 2026, a 9.5% decline from 2025's pace but still historically elevated. The full year 2026 is already showing 772 total breach filings, maintaining healthcare's status as the highest-breach industry.",
            "source_text": "HIPAA Journal, 2026",
            "source_url": "https://www.hipaajournal.com/healthcare-data-breach-statistics/",
        },
        {
            "headline": "Healthcare Data Breach Cost Hits $7.42M Average — Highest of Any Industry, 14th Consecutive Year",
            "summary": "Healthcare remains the most costly industry for data breaches, with average incident costs reaching $7.42 million — driven by breach notification, regulatory fines, litigation, and recovery operations. The figure reflects both the increasing sensitivity of healthcare data and the growing sophistication of attackers targeting the sector.",
            "source_text": "SecurityWeek / DeepStrike, 2026",
            "source_url": "https://www.securityweek.com/millions-impacted-across-several-us-healthcare-data-breaches/",
        },
        {
            "headline": "Erie Family Health Centers Breach Affects 570,000 Individuals",
            "summary": "Erie Family Health Centers reported a significant data breach affecting 570,000 individuals, part of a wave of February 2026 healthcare breaches that also included TriZetto Provider Solutions and QualDerm Partners. Third-party business associate vulnerabilities were identified as the primary entry point in multiple simultaneous incidents.",
            "source_text": "HIPAA Journal, Feb 2026",
            "source_url": "https://www.hipaajournal.com/february-2026-healthcare-data-breach-report/",
        },
        {
            "headline": "Over 80% of Stolen Healthcare Records Now Coming from Third-Party Vendors",
            "summary": "Analysis of recent breaches confirms that more than 80% of all stolen patient records originate from third-party vendors, business associates, and non-hospital providers — not directly from health system infrastructure. The pattern is driving health systems to implement third-party risk management programs and enhanced vendor contract security provisions.",
            "source_text": "HIPAA Journal / Swif AI, 2026",
            "source_url": "https://www.swif.ai/blog/healthcare-cybersecurity-statistics",
        },
    ],
    "3. Hospital M&A & Strategic Consolidation": [
        {
            "headline": "Q1 2026: 22 Hospital M&A Transactions, $14.5B in Revenue — Strongest Q1 in 5+ Years",
            "summary": "Kaufman Hall's Q1 2026 M&A report found 22 announced transactions representing $14.5 billion in transacted revenue — the highest Q1 figure in at least five years. Three mega mergers (revenue >$1B for smaller party) were included, reflecting a new appetite for transformational consolidation.",
            "source_text": "Kaufman Hall, Apr 2026",
            "source_url": "https://www.kaufmanhall.com/insights/research-report/ma-quarterly-activity-report-q1-2026",
        },
        {
            "headline": "UPMC and CommonSpirit Health Sign Definitive Merger Agreement, Expected Close Fall 2026",
            "summary": "UPMC and CommonSpirit Health announced a definitive merger agreement creating one of the nation's largest health systems, with the transaction expected to close in fall 2026. The combination will create a coast-to-coast integrated delivery network spanning dozens of markets.",
            "source_text": "Kaufman Hall / Modern Healthcare, 2026",
            "source_url": "https://www.modernhealthcare.com/mergers-acquisitions/mh-healthcare-deals-live-updates/",
        },
        {
            "headline": "Sutter Health (CA) and Allina Health (MN) Propose Cross-Regional Merger",
            "summary": "California-based Sutter Health and Minnesota-based Allina Health announced a proposed merger creating a large cross-regional health system — one of the most geographically ambitious hospital mergers in recent years. Regulatory review is underway as the system prepares integration planning.",
            "source_text": "Kaufman Hall Q1 2026 Report",
            "source_url": "https://www.kaufmanhall.com/insights/research-report/ma-quarterly-activity-report-q1-2026",
        },
        {
            "headline": "Atrium Health and WakeMed Propose Merger with $2B Investment in Wake County",
            "summary": "Atrium Health and WakeMed announced a proposed combination backed by a $2 billion commitment to expand Wake County hospital facilities, grow the clinical workforce, and broaden care access across North Carolina. The deal positions the combined entity as the dominant provider in the greater Raleigh market.",
            "source_text": "Healthcare Brew, Apr 1, 2026",
            "source_url": "https://www.healthcare-brew.com/stories/2026/04/01/march-2026-hospital-mergers-acquisitions",
        },
        {
            "headline": "21 Large Health Systems Growing in Size Through M&A and Expansion (June 2026 Analysis)",
            "summary": "Becker's June 2026 analysis identified 21 large health systems actively growing through mergers, acquisitions, or affiliate expansions — signaling that scale-seeking behavior is not limited to large academic or for-profit systems but is now a strategic imperative across health system types.",
            "source_text": "Becker's Hospital Review, Jun 2026",
            "source_url": "https://www.beckershospitalreview.com/",
        },
    ],
    "4. Hospital Finances, Medicaid & Federal Policy": [
        {
            "headline": "One Big Beautiful Bill Act Cuts $911B in Medicaid; 800+ Facilities Closed or At Risk",
            "summary": "Since the OBBBA was signed July 4, 2025, the Congressional Budget Office projects $911 billion in federal Medicaid spending cuts over 10 years. More than 800 hospitals, nursing homes, maternity wards, and psychiatric centers have already closed, cut services, or reported immediate risk — driven by enrollment declines and provider tax restrictions.",
            "source_text": "AMA / Center for American Progress, 2026",
            "source_url": "https://www.ama-assn.org/health-care-advocacy/federal-advocacy/changes-medicaid-aca-and-other-key-provisions-one-big",
        },
        {
            "headline": "CMS Closes Medicaid MCO Tax Loophole Effective April 3, 2026 — Saving $24B+ Annually",
            "summary": "CMS finalized a rule closing a longstanding Medicaid financing gimmick in which states imposed taxes on MCOs to draw down inflated federal matching funds. The rule — effective April 3, 2026 — eliminates schemes generating $24+ billion annually in excess federal matching dollars and forces states to restructure their Medicaid financing.",
            "source_text": "CMS, 2026",
            "source_url": "https://www.cms.gov/newsroom/press-releases/cms-shuts-down-massive-medicaid-tax-loophole-saving-billions-federal-taxpayers-restoring-federal",
        },
        {
            "headline": "720 Hospitals at Risk of Closure by State — June 2026 Analysis",
            "summary": "A June 2026 analysis identified 720 hospitals at risk of closure nationally, broken down by state — with rural and high-Medicaid-mix facilities most exposed. The analysis cites OBBBA enrollment effects, CMS payment reductions, and labor cost inflation as the compounding drivers.",
            "source_text": "Becker's Hospital Review, Jun 2026",
            "source_url": "https://www.beckershospitalreview.com/",
        },
        {
            "headline": "2026 Medicare Payment Final Rule Delays Major Hospital Payment Cut",
            "summary": "The 2026 Medicare payment final rule postponed a significant accelerated payment reduction for hospitals — but the cut remains on the legislative horizon and is expected to converge with OBBBA Medicaid rollbacks, creating a compound revenue pressure event for health systems.",
            "source_text": "HFMA, 2026",
            "source_url": "https://www.hfma.org/payment-reimbursement-and-managed-care/2026-medicare-final-rule-postpones-a-significant-payment-cut-for-hospitals/",
        },
        {
            "headline": "$50B Rural Health Fund Created — But Analysts Say It Covers Only 37% of Projected Rural Medicaid Losses",
            "summary": "Congress established a $50 billion Rural Health Fund within OBBBA to stabilize rural hospitals — but RAND analysis and Congressional Budget Office projections show the fund covers only about 37% of the projected loss of federal Medicaid funding in rural areas. More than 300 rural hospitals are at immediate risk of closure despite the fund.",
            "source_text": "RAND / KFF, 2026",
            "source_url": "https://www.rand.org/pubs/research_reports/RRA4098-1.html",
        },
    ],
    "5. New Medications, Treatments & GLP-1s": [
        {
            "headline": "FDA Approves Oral Semaglutide (Wegovy 25mg) — First GLP-1 Pill for Weight Loss",
            "summary": "The FDA approved once-daily oral semaglutide 25mg (Wegovy) as the first GLP-1 receptor agonist pill for weight loss, representing a major shift in obesity treatment delivery from injectable to oral administration. Novo Nordisk launched the product in the U.S. in early January 2026, dramatically expanding the eligible patient population.",
            "source_text": "AJMC / Nature, 2026",
            "source_url": "https://www.ajmc.com/view/fda-approves-oral-semaglutide-as-first-glp-1-pill-for-weight-loss",
        },
        {
            "headline": "Eli Lilly's Orforglipron (Foundayo) Approved: First Small-Molecule GLP-1 for Obesity",
            "summary": "The FDA approved Eli Lilly's orforglipron (brand: Foundayo) as the first small-molecule (non-peptide) GLP-1 receptor agonist for obesity treatment — a chemically distinct class from semaglutide that may have different side-effect profiles and manufacturing cost advantages.",
            "source_text": "CNBC, Jan 2026",
            "source_url": "https://www.cnbc.com/2026/01/10/2026-is-the-year-of-obesity-pills-from-novo-nordisk-eli-lilly-.html",
        },
        {
            "headline": "Medicare GLP-1 Bridge Program Begins July 2026 — Weight-Loss Drug Coverage for Medicare Patients",
            "summary": "Starting July 2026, qualifying Medicare patients became eligible for coverage of certain GLP-1 weight-loss medications through the new Medicare GLP-1 Bridge Program, dramatically expanding access for older Americans who were previously excluded from obesity drug coverage under Medicare.",
            "source_text": "Medicare Rights Center, Jun 4, 2026",
            "source_url": "https://www.medicarerights.org/medicare-watch/2026/06/04/glp-1-weight-loss-drug-demonstration-begins-july-2026",
        },
        {
            "headline": "Tirzepatide Gains FDA Approval for Sleep Apnea; Expands GLP-1 Indications Beyond Obesity",
            "summary": "Tirzepatide (Mounjaro/Zepbound) received FDA approval for sleep apnea, adding to its obesity and type 2 diabetes indications and expanding GLP-1 drugs into pulmonology and sleep medicine. The approval further validates GLP-1 agents as multi-system therapeutic tools relevant to multiple hospital service lines.",
            "source_text": "IQVIA, 2026",
            "source_url": "https://www.iqvia.com/locations/emea/blogs/2026/01/outlook-for-obesity-in-2026",
        },
        {
            "headline": "Retatrutide Shows 28.7% Average Weight Loss in Knee Osteoarthritis Trial — Regulatory Filing Expected",
            "summary": "Eli Lilly's retatrutide demonstrated 28.7% average weight loss and up to 75.8% pain reduction in adults with obesity and knee osteoarthritis, with a regulatory submission expected in 2026. The result may open GLP-1 therapy as a pre-surgical or surgery-avoidance option in orthopedic care.",
            "source_text": "IQVIA / FRQ Tech, 2026",
            "source_url": "https://www.frqtech.ai/blog/top-7-fda-approved-obesity-drugs-in-2026",
        },
    ],
    "6. Clinical Trials & FDA Regulatory Shifts": [
        {
            "headline": "FDA Launches Real-Time Clinical Trials (RTCT) Initiative — AstraZeneca and Amgen as First Participants",
            "summary": "FDA announced the successful initiation of two proof-of-concept Real-Time Clinical Trials: AstraZeneca's mantle cell lymphoma study and Amgen's small cell lung cancer trial. Under RTCT, trial data is streamed to FDA scientists in real time via Paradigm Health's Study Conduct platform, potentially compressing review timelines dramatically.",
            "source_text": "FDA / HHS, Apr 29, 2026",
            "source_url": "https://www.fda.gov/news-events/press-announcements/fda-announces-major-steps-implement-real-time-clinical-trials",
        },
        {
            "headline": "FDA Formally Accepts Bayesian Methodology in Pivotal Clinical Trials via January 2026 Draft Guidance",
            "summary": "FDA's January 12, 2026 draft guidance on 'Use of Bayesian Methodology in Clinical Trials of Drug and Biological Products' marks the most significant formal regulatory shift toward adaptive trial designs in decades. Systems conducting oncology, rare disease, or complex chronic disease trials can now design Bayesian primary analyses for pivotal studies.",
            "source_text": "FDA / arXiv, Jan 2026",
            "source_url": "https://www.fda.gov/",
        },
        {
            "headline": "FDA Reminds 2,200+ Sponsors to Disclose Clinical Trial Results on ClinicalTrials.gov",
            "summary": "FDA issued formal reminders to more than 2,200 clinical trial sponsors and researchers to submit results to ClinicalTrials.gov as required by law, signaling increased enforcement attention. Health systems conducting research must ensure compliance to avoid FDA warning letters and reputational risk.",
            "source_text": "FDA, 2026",
            "source_url": "https://www.fda.gov/news-events/press-announcements/fda-reminds-more-2200-sponsors-and-researchers-disclose-trial-results",
        },
        {
            "headline": "FDA Finalizes Guidance on Enhancing Participation in Clinical Trials — New Enrollment Design Requirements",
            "summary": "Finalized December 15, 2025 and taking effect in 2026, FDA's guidance on 'Enhancing Participation in Clinical Trials' formally updates expectations for inclusive enrollment design, requiring explicit plans for historically underrepresented populations in trial protocols.",
            "source_text": "ClinicalLeader, 2026",
            "source_url": "https://www.clinicalleader.com/doc/fda-issues-final-guidance-on-clinical-trial-participation-what-you-need-to-do-now-0001",
        },
        {
            "headline": "Paradigm Health–FDA Collaboration Transforms Regulatory Review Using Real-Time Data Streaming",
            "summary": "Paradigm Health announced a collaboration with FDA to automate data collection, analysis, and reporting of safety and efficacy signals during clinical trials — enabling the FDA's RTCT initiative. The model could become the standard for research-active health systems seeking accelerated regulatory engagement.",
            "source_text": "PR Newswire, 2026",
            "source_url": "https://www.prnewswire.com/news-releases/paradigm-health-announces-collaboration-with-the-fda-to-transform-regulatory-review-of-clinical-trial-data-302756035.html",
        },
    ],
    "7. Health IT / EHR & Digital Infrastructure": [
        {
            "headline": "Epic Holds 43.7% of Acute Care Market Share; Oracle Health Sheds 56 Hospitals in One Year",
            "summary": "May 2026 KLAS analysis found Epic serving 43.7% of acute care hospitals (56.9% of beds), gaining 77 hospitals in 2025 alone. Oracle Health lost 56 hospitals in the same period and has shed a net 173 since 2021 — creating a widening competitive gap that is forcing migration decisions at hundreds of Cerner-legacy institutions.",
            "source_text": "HealthSystemCIO / KLAS, May 14, 2026",
            "source_url": "https://healthsystemcio.com/2026/05/14/acute-care-ehr-market-share-2026/",
        },
        {
            "headline": "Oracle Health: 30% of Millennium Customers Not in Long-Term Plans; Ranked Lowest in KLAS 2026",
            "summary": "Oracle Health's Millennium EHR ranked as the lowest-scoring acute care EHR in the 2026 Best in KLAS report across all organization sizes. Sampled customers show 30% are not including Millennium in long-term IT plans and 35% are considered 'vulnerable' — raising migration urgency for Oracle-dependent health systems.",
            "source_text": "EHR Source / KLAS, 2026",
            "source_url": "https://www.ehrsource.com/vendors/oracle-health/",
        },
        {
            "headline": "OhioHealth Morrow County Hospital Launches Epic in $6M IT Investment",
            "summary": "OhioHealth Morrow County Hospital went live with Epic EHR on February 22 as part of a $6 million IT investment — reflecting the ongoing trend of smaller and critical access hospitals upgrading to Epic as the platform's small-hospital market share grows. Epic's KLAS gains among small health systems are accelerating.",
            "source_text": "Becker's Hospital Review, Feb 2026",
            "source_url": "https://www.beckershospitalreview.com/healthcare-information-technology/ehrs/",
        },
        {
            "headline": "300+ Health Systems Using Epic AI Tools in Production as of Early 2026",
            "summary": "Epic confirmed that over 300 health systems are actively using AI-powered tools in production — spanning predictive deterioration, sepsis, readmission risk, no-show prediction, and ambient documentation via partnerships with Nuance DAX Copilot and Abridge. The milestone cements Epic's position as the dominant AI-enabled EHR platform.",
            "source_text": "Becker's Hospital Review, 2026",
            "source_url": "https://www.beckershospitalreview.com/healthcare-information-technology/ai/",
        },
        {
            "headline": "Oracle Health's AI-Enabled EHR Enters Production in 2026 — Strategic Make-or-Break Year",
            "summary": "Oracle Health's new AI-enabled EHR platform moved into production environments in 2026, representing the company's primary strategy to stabilize its shrinking customer base. Reception among early adopters will determine whether Oracle can reverse a 5-year market share decline or accelerate customer departures.",
            "source_text": "EHR Source, 2026",
            "source_url": "https://www.ehrsource.com/compare/epic-vs-oracle-health/",
        },
    ],
    "8. Hospital Safety, Workforce Violence & Nursing Shortage": [
        {
            "headline": "Illinois Nurses Union Sues Prime Healthcare and Ascension Over Alleged Understaffing",
            "summary": "In June 2026, Illinois nurses unions filed suit against Prime Healthcare and Ascension Health alleging chronic understaffing that endangers patient safety and violates nurse-patient ratio commitments. The action reflects a national pattern of unions using litigation to enforce staffing standards as legislative solutions stall.",
            "source_text": "Becker's Hospital Review, Jun 2026",
            "source_url": "https://www.beckershospitalreview.com/workforce/",
        },
        {
            "headline": "138,000 Nurses Have Left the Workforce Since 2022; 45% Likely to Leave in Next 12 Months",
            "summary": "Nurse.org's 2026 State of Nursing Survey found that 138,000 nurses have departed the profession since 2022 and 45% of current nurses indicate they are likely to leave their jobs within the next 12 months. Stress, workplace violence, administrative burden, and inadequate compensation are the top cited drivers.",
            "source_text": "Nurse.org State of Nursing Survey, 2026",
            "source_url": "https://nurse.org/articles/state-of-nursing-survey-2026/",
        },
        {
            "headline": "54% of Healthcare Workers Have Felt Threatened by Patients or Visitors on the Job",
            "summary": "A 2026 Verkada/Harris Poll survey found that 54% of healthcare workers have felt threatened by patients or visitors, with younger workers experiencing higher rates of witnessed or experienced violence. The data reinforces that workplace violence is not an outlier event but a pervasive daily reality for clinical staff.",
            "source_text": "ASIS International / Verkada/Harris Poll, Feb 2026",
            "source_url": "https://www.asisonline.org/security-management-magazine/articles/2026/02/visitor-management-in-healthcare/workplace-violence-views/",
        },
        {
            "headline": "One in Three Rural Hospitals at Risk of Closure; Rural Leaders Betting on Care Redesign",
            "summary": "Becker's May 2026 issue highlighted that one in three U.S. rural hospitals is at risk of closure — a figure worsened by OBBBA Medicaid cuts, nurse shortages, and inflation in supply and labor costs. Rural health leaders are responding with care redesign initiatives including telehealth expansion and ambulatory care investment.",
            "source_text": "Becker's Hospital Review, May 2026",
            "source_url": "https://www.beckershospitalreview.com/print-issues/may-2026-issue-of-beckers-hospital-review/",
        },
        {
            "headline": "Federal Healthcare Worker Violence Bill Passes House — Remains Stalled in Senate for Third Year",
            "summary": "Legislation to provide federal protections for healthcare workers facing workplace violence passed the House for the third consecutive congressional session but has again stalled in the Senate, leaving hospitals without federal mandate support for safety programs. The bipartisan bill has broad support but faces procedural and legislative calendar barriers.",
            "source_text": "AllNurses, 2026",
            "source_url": "https://allnurses.com/news/workplace-violence-nursing-legislation-r65/",
        },
    ],
    "9. Hospital Administration & Organizational Structure": [
        {
            "headline": "NYC Health + Hospitals Establishes VP and Chief Data and AI Officer Role",
            "summary": "NYC Health + Hospitals, the nation's largest municipal health system, established a VP and Chief Data and AI Officer position — signaling that even publicly operated safety-net systems are formalizing AI governance at the C-suite level. The role oversees enterprise data strategy, AI deployment, and algorithmic accountability.",
            "source_text": "AHA Center for Health Innovation, Feb 2026",
            "source_url": "https://www.aha.org/aha-center-health-innovation-market-scan/2026-02-03-c-suite-rewrites-org-chart-how-health-systems-are-elevating-technology",
        },
        {
            "headline": "BayCare Elevates Ambulatory Leadership to Executive Cabinet, Merging with System Strategy",
            "summary": "BayCare Health System restructured its executive team to combine ambulatory services leadership directly with system strategy — reflecting the shift of health system growth from inpatient to outpatient care. The structural change positions ambulatory as a primary growth engine rather than a secondary support function.",
            "source_text": "Becker's Hospital Review, 2026",
            "source_url": "https://www.beckershospitalreview.com/hospital-management-administration/how-ambulatory-care-shifts-are-reshaping-health-system-leadership-structures/",
        },
        {
            "headline": "Care New England Cuts 30+ Leadership Positions to Close $20M Budget Gap",
            "summary": "Care New England eliminated more than 30 leadership and nonclinical positions as part of a 2026 cost restructuring plan targeting a $20 million budget deficit. The action reflects 'The Great Flattening' trend — health systems reducing administrative overhead while protecting frontline care capacity.",
            "source_text": "HealthLeaders, 2026",
            "source_url": "https://www.healthleadersmedia.com/ceo/why-hospitals-are-targeting-leadership-administrative-jobs-latest-layoffs",
        },
        {
            "headline": "Becker's June 2026: 66 Health Systems Ranked by Long-Term Debt; 64 by Total Assets",
            "summary": "Becker's Hospital Review published comprehensive rankings of 66 and 64 large health systems by long-term debt and total assets respectively, reflecting intensifying scrutiny of health system balance sheets. The rankings are being used by analysts, bond raters, and M&A advisors to identify financially vulnerable and strategically attractive institutions.",
            "source_text": "Becker's Hospital Review, Jun 2026",
            "source_url": "https://www.beckershospitalreview.com/finance/",
        },
        {
            "headline": "AdventHealth Shifts from Mandatory Training to Continuous Learning Culture Across 100,000 Team Members",
            "summary": "AdventHealth announced a system-wide transition from episodic, mandatory training programs to a culture of continuous, competency-based learning across more than 100,000 employees in 2026. The initiative reflects a broader hospital sector trend toward learning infrastructure as a workforce retention and safety quality tool.",
            "source_text": "Becker's Hospital Review, May 2026",
            "source_url": "https://www.beckershospitalreview.com/print-issues/may-2026-issue-of-beckers-hospital-review/",
        },
    ],
    "10. Community Relations & Health Equity": [
        {
            "headline": "Endeavor Health Awards $8.1M in 2026 Community Investment Fund Grants",
            "summary": "Endeavor Health announced $8.1 million in 2026 Community Investment Fund (CIF) awards, including nine Capacity Builder grants ($100K–$500K each) for food access, housing, and telehealth programs, and a $5 million, five-year Impact Award to Loaves & Fishes for regional food access infrastructure. Since CIF's 2022 launch, $38.5M+ has been invested supporting 1 million+ lives across Chicagoland.",
            "source_text": "Evanston RoundTable, Jan 14, 2026",
            "source_url": "https://evanstonroundtable.com/2026/01/14/endeavor-health-announces-2026-recipients-of-community-investment-funds/",
        },
        {
            "headline": "2026 SDOH & Health Equity Forum Convenes in Chicago — June 25-26",
            "summary": "The 2026 Social Determinants of Health and Health Equity Forum brought together hospital and health system executives in Chicago on June 25–26 to address strategies for managing health equity programs and SDOH integration at the community level. The Chicago venue underscores the Midwest's role as a health equity innovation hub.",
            "source_text": "WC Forum, 2026",
            "source_url": "https://www.wcforum.com/conferences/sdoh",
        },
        {
            "headline": "OBBBA Medicaid Work Requirements Expected to Disproportionately Impact Low-Income, Minority Communities",
            "summary": "Urban Institute analysis found that the OBBBA's Medicaid work requirements leave 3 in 10 young adults vulnerable to losing healthcare access — with disproportionate impact on communities of color, rural residents, and those with intermittent employment. Hospital equity programs must adapt to serve a growing uninsured population.",
            "source_text": "Urban Institute, 2026",
            "source_url": "https://www.urban.org/urban-wire/medicaid-cuts-one-big-beautiful-bill-act-leave-3-10-young-adults-vulnerable-losing",
        },
        {
            "headline": "U.S. News Hospital Rankings Now Include Social Representation and Racial Disparity Metrics",
            "summary": "U.S. News and World Report's 2024–2025 hospital rankings — now shaping 2026 strategy — incorporate social representation and racial disparities in length of stay as evaluation criteria. Health systems are responding by embedding equity metrics into quality dashboards and community health needs assessments.",
            "source_text": "NEJM Catalyst, 2026",
            "source_url": "https://catalyst.nejm.org/doi/full/10.1056/CAT.22.0329",
        },
        {
            "headline": "CMS Framework for Health Equity Now Governing Program Design Across All CMS Initiatives",
            "summary": "CMS released its Framework for Health Equity, requiring all CMS programs and policies to address health disparities through design, implementation, and operationalization. The framework affects how Medicaid, Medicare Advantage, and ACO programs are evaluated, with equity metrics increasingly tied to reimbursement incentives.",
            "source_text": "CMS, 2026",
            "source_url": "https://www.cms.gov/priorities/program/minority-health/equity-programs",
        },
    ],
}

ENDEAVOR_SECTIONS = [
    {
        "area": "System Overview",
        "findings": (
            "Endeavor Health is Illinois' third-largest health system, comprising nine award-winning hospitals, "
            "300+ clinic locations, and more than 27,000 team members across Chicagoland. "
            "The system was formed by the merger of NorthShore University HealthSystem, Swedish Hospital, "
            "Northwest Community Healthcare, and Edward-Elmhurst Health. "
            "Its flagship campus is Evanston Hospital (Evanston, IL), which anchors the NorthShore Hospitals group "
            "including Glenbrook, Highland Park, and Skokie Hospitals. "
            "Endeavor operates an active Research Institute and maintains 300+ clinic locations across northern Illinois and metropolitan Chicago."
        ),
        "relevance": (
            "Understanding Endeavor's scale and mission context is essential for any meeting focused on partnership, "
            "research collaboration, innovation investment, or clinical service development. "
            "As IL's 3rd-largest system, Endeavor operates at sufficient scale to be a consequential partner or competitor "
            "in any regional healthcare initiative."
        ),
        "source_text": "Endeavor Health",
        "source_url": "https://www.endeavorhealth.org/",
    },
    {
        "area": "Leadership Profile: Gabrielle Cummings-Johnson",
        "findings": (
            "Gabrielle Cummings-Johnson serves as President of Endeavor Health Evanston Hospital and as the senior operations executive "
            "for Endeavor's NorthShore Hospitals group (Evanston, Glenbrook, Highland Park, Skokie). "
            "She also leads the Endeavor Health Cancer Service Line system-wide. "
            "Cummings-Johnson joined Endeavor as an Administrative Fellow in 2002 and has progressively held senior roles at Evanston, "
            "Glenbrook, and Highland Park Hospitals before being named President of Evanston Hospital effective January 1, 2024. "
            "She is a champion of Endeavor's health equity programs and has actively guided the Community Investment Fund strategy. "
            "Her profile is featured on Becker's Hospital Review podcast and she is active on LinkedIn."
        ),
        "relevance": (
            "Cummings-Johnson's trajectory from administrative fellow to hospital president reflects deep organizational loyalty and operational breadth — "
            "she understands Endeavor's culture from the inside. Her dual role across hospital operations and the Cancer Service Line "
            "signals both operational and clinical priorities. Health equity and community investment are established leadership themes to engage around."
        ),
        "source_text": "Endeavor Health / Becker's Hospital Review",
        "source_url": "https://www.endeavorhealth.org/leadership/gabrielle-cummings-johnson",
    },
    {
        "area": "AI & Innovation Programs",
        "findings": (
            "Endeavor Health has scaled Abridge ambient documentation to 1,000 clinician licenses with 250+ actively using the platform — "
            "one of the largest deployments of ambient documentation in Illinois. "
            "The system operates a decade-long precision medicine program: 225,000+ patients screened for hereditary conditions; "
            "35,000+ completed clinical genetic testing; 100+ primary care providers trained in genomic medicine. "
            "Endeavor partnered with Northwestern to develop an AI-enabled lung disorder detection program. "
            "The Endeavor Health Precision Medicine & AI Fellowship at the University of Chicago Department of Family Medicine "
            "trains next-generation physician innovators. "
            "A peer-reviewed publication in PMC documents Endeavor's experience operationalizing personalized medicine in a community health system."
        ),
        "relevance": (
            "Endeavor's AI and precision medicine programs are operationally mature — not aspirational. "
            "The scale of genomic screening (225,000+ patients in a community hospital setting) is unusual and nationally notable. "
            "Meetings touching on AI governance, precision medicine partnerships, or clinical trial recruitment should reference Endeavor's existing infrastructure."
        ),
        "source_text": "Becker's Hospital Review / UChicago Medicine",
        "source_url": "https://familymedicine.uchicago.edu/education/endeavor-health-precision-medicine-ai-fellowship",
    },
    {
        "area": "Clinical Trials & Research Activity",
        "findings": (
            "Endeavor Health conducts clinical research through the NorthShore University HealthSystem Research Institute, "
            "with an active portal at clinicaltrials.endeavorhealth.org. "
            "Active and recent trials cover conditions including Acute Myeloid Leukemia and ARDS (Acute Respiratory Distress Syndrome) as of June 2026. "
            "The Neuroscience Institute maintains a dedicated Research & Clinical Trials program. "
            "ClinicalTrials.gov lists Endeavor/NorthShore Research Institute as an active trial site across multiple therapeutic areas. "
            "Research participants receive study-related medical care at no cost; the Institute partners with both academic and industry sponsors."
        ),
        "relevance": (
            "Endeavor's research infrastructure positions the system to participate in emerging FDA real-time clinical trial models (RTCT) "
            "and AI-assisted trial design programs. "
            "For pharma or medtech companies, Endeavor is a credible community-system trial site with established recruitment infrastructure — "
            "a potential partnership angle in any research-focused meeting."
        ),
        "source_text": "Endeavor Health Clinical Trials / ClinicalTrials.gov",
        "source_url": "https://clinicaltrials.endeavorhealth.org/",
    },
    {
        "area": "Recent Strategic Moves",
        "findings": (
            "Endeavor Health's primary 2024–2026 strategic move has been the full brand and operational integration of four legacy systems "
            "into the unified Endeavor Health identity — a complex multi-system merger requiring clinical, IT, and cultural alignment. "
            "Kristen Murtos serves as Chief Innovation and Transformation Officer, overseeing system strategy, innovation, community impact, "
            "and government relations — the executive counterpart to Cummings-Johnson's operational portfolio. "
            "Endeavor was featured on Becker's 2026 list of '48 Hospital and Health System Chief Innovation Officers to Know,' "
            "signaling external recognition of its innovation leadership. "
            "The system continues ambulatory care expansion to support financial sustainability and access."
        ),
        "relevance": (
            "Endeavor is mid-integration — a period of both opportunity and organizational stress. "
            "Decisions about research partnerships, technology investments, and clinical program development are being made now "
            "as the integrated system sets its post-merger strategic agenda. "
            "Any meeting is occurring at a strategically receptive moment."
        ),
        "source_text": "Becker's Hospital Review 2026 / Endeavor Health",
        "source_url": "https://www.beckershospitalreview.com/hospital-management-administration/48-hospital-and-health-system-chief-innovation-officers-to-know-2026/",
    },
    {
        "area": "Community Investment & Health Equity",
        "findings": (
            "Endeavor Health's Community Investment Fund (CIF) is one of the most active hospital-led community investment programs in Illinois. "
            "In 2026, Endeavor awarded $8.1 million through nine Capacity Builder awards and one Impact Award ($5M over 5 years to Loaves & Fishes). "
            "Since CIF's 2022 launch: $38.5M+ invested; 730+ jobs created or supported; 1 million+ lives impacted across Chicagoland. "
            "A $101,626 Capacity Builder Award went to North Central College for telehealth education expansion. "
            "2027 CIF Impact Award applications are now open (announced May 2026), signaling the program's ongoing commitment. "
            "Endeavor also partnered with Oakton College in 2024 to expand healthcare education access in Evanston and surrounding areas."
        ),
        "relevance": (
            "Cummings-Johnson personally champions the CIF — making it a relevant relationship-building topic in any meeting. "
            "The program's SDOH investment focus (food, housing, telehealth) aligns with CMS community benefit expectations "
            "and positions Endeavor favorably for Medicaid managed care and value-based contract evaluations. "
            "The sheer scale ($38.5M since 2022) demonstrates institutional commitment, not pilot-level activity."
        ),
        "source_text": "Evanston RoundTable / Endeavor Health, Jan 2026",
        "source_url": "https://evanstonroundtable.com/2026/01/14/endeavor-health-announces-2026-recipients-of-community-investment-funds/",
    },
]

# ──────────────────────────────────────────────────────
# DOCUMENT BUILDER
# ──────────────────────────────────────────────────────

def build_document():
    doc = Document()
    set_landscape(doc)
    add_footer(doc, (
        "Methodology: This briefing was generated on June 22, 2026 via live web research across Becker's Hospital Review, "
        "Modern Healthcare, Kaufman Hall, HIPAA Journal, CMS, KFF, AJMC, FDA, IQVIA, AHA, Fierce Healthcare, and Endeavor Health. "
        "All sources are hyperlinked. Stories selected using a multi-criteria importance filter: material financial/strategic impact, "
        "measurable trend, executive relevance, and recency (preference for past 30–60 days). "
        "Endeavor Health section based on dedicated searches of Endeavor's public web properties, ClinicalTrials.gov, Becker's, and Evanston RoundTable."
    ))

    # ── COVER ───────────────────────────────────────────
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title_p.add_run("HEALTHCARE EXECUTIVE BRIEFING")
    r.font.name = 'Arial'
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = NAVY

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub_p.add_run("Comprehensive Intelligence Report for Hospital Executive Meeting Preparation")
    r2.font.name = 'Arial'
    r2.font.size = Pt(13)
    r2.font.bold = False
    r2.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = date_p.add_run("Prepared: June 22, 2026  |  Special Focus: Endeavor Health / Evanston Hospital / Gabrielle Cummings-Johnson")
    r3.font.name = 'Arial'
    r3.font.size = Pt(10)
    r3.font.italic = True
    r3.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph()

    # Source line
    src_p = doc.add_paragraph()
    src_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = src_p.add_run(
        "KEY SOURCES: Becker's Hospital Review  |  Kaufman Hall  |  HIPAA Journal  |  CMS  |  KFF  |  "
        "Modern Healthcare  |  AJMC  |  FDA  |  AHA  |  IQVIA  |  Fierce Healthcare  |  "
        "Evanston RoundTable  |  Endeavor Health"
    )
    r4.font.name = 'Arial'
    r4.font.size = Pt(8)
    r4.font.bold = True
    r4.font.color.rgb = NAVY

    doc.add_paragraph()

    # ── SECTION 1 ──────────────────────────────────────
    section_heading(doc, "SECTION 1: TOPIC BRIEFING — Current Healthcare Landscape (10 Topic Areas)")

    # col widths sum to ~10" (landscape 11" - 1" margins)
    col_w = [1.5, 3.5, 2.5, 1.2, 1.3]
    tbl = make_table(doc, 5, col_w)
    header_row(tbl, "Topic Area", "What's Going On", "Why It Matters to Hospital Executives", "Growing Trend?", "Key Source")

    for i, item in enumerate(TOPIC_BRIEFING, start=1):
        row = tbl.add_row()
        fill = 'FFFFFF' if i % 2 == 0 else 'D6E4F0'
        cells = row.cells
        for c in cells:
            set_cell_bg(c, fill)
            c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        cell_para(cells[0], item["topic"], bold=True, size=9)
        cell_para(cells[1], item["whats_going_on"], size=8.5)
        cell_para(cells[2], item["why_matters"], size=8.5)
        cell_para(cells[3], item["trend"], size=8.5, italic=True)
        add_hyperlink(cells[4].paragraphs[0], item["source_url"], item["source_text"], font_size=8.5)

    doc.add_paragraph()

    # ── SECTION 2 ──────────────────────────────────────
    section_heading(doc, "SECTION 2: EMERGING TRENDS — Cross-Cutting Industry Themes")

    col_w2 = [1.5, 2.8, 2.8, 1.1, 1.8]
    tbl2 = make_table(doc, 5, col_w2)
    header_row(tbl2, "Trend Name", "What Is Changing", "Future Implication for Health Systems", "Timeframe", "Source(s)")

    for i, item in enumerate(EMERGING_TRENDS, start=1):
        row = tbl2.add_row()
        fill = 'FFFFFF' if i % 2 == 0 else 'D6E4F0'
        cells = row.cells
        for c in cells:
            set_cell_bg(c, fill)
            c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        cell_para(cells[0], item["name"], bold=True, size=9)
        cell_para(cells[1], item["what_changing"], size=8.5)
        cell_para(cells[2], item["implication"], size=8.5)
        cell_para(cells[3], item["timeframe"], size=8.5, italic=True)
        add_hyperlink(cells[4].paragraphs[0], item["source_url"], item["source_text"], font_size=8.5)

    doc.add_paragraph()

    # ── SECTION 3 ──────────────────────────────────────
    section_heading(doc, "SECTION 3: NEWS STORIES BY TOPIC — Up to 5 Specific, Sourced Stories per Topic")

    col_w3 = [1.5, 8.5]
    tbl3 = make_table(doc, 2, col_w3)
    header_row(tbl3, "Topic", "Recent News Stories (Headline → Summary → Source)")

    story_idx = 0
    for topic_name, stories in NEWS_STORIES.items():
        row = tbl3.add_row()
        fill = 'FFFFFF' if story_idx % 2 == 0 else 'D6E4F0'
        for c in row.cells:
            set_cell_bg(c, fill)
            c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        cell_para(row.cells[0], topic_name, bold=True, size=9)

        # Build stories in the second cell
        right_cell = row.cells[1]
        first = True
        for j, story in enumerate(stories):
            if first:
                para = right_cell.paragraphs[0]
                first = False
            else:
                para = right_cell.add_paragraph()
                para.paragraph_format.space_before = Pt(4)

            # Number + bold headline
            num_run = para.add_run(f"{j+1}. ")
            num_run.font.name = 'Arial'
            num_run.font.size = Pt(8.5)
            num_run.font.bold = True

            hl_run = para.add_run(story["headline"] + " → ")
            hl_run.font.name = 'Arial'
            hl_run.font.size = Pt(8.5)
            hl_run.font.bold = True

            sum_run = para.add_run(story["summary"] + "  ")
            sum_run.font.name = 'Arial'
            sum_run.font.size = Pt(8.5)
            sum_run.font.bold = False

            src_label = para.add_run("Source: ")
            src_label.font.name = 'Arial'
            src_label.font.size = Pt(8.5)
            src_label.font.bold = False
            src_label.font.italic = True

            add_hyperlink(para, story["source_url"], story["source_text"], font_size=8.5)

        story_idx += 1

    doc.add_paragraph()

    # ── SECTION 4 ──────────────────────────────────────
    section_heading(doc, "SECTION 4: ENDEAVOR HEALTH DEEP DIVE — Evanston Hospital & Gabrielle Cummings-Johnson")

    col_w4 = [1.7, 4.1, 2.5, 1.7]
    tbl4 = make_table(doc, 4, col_w4)
    header_row(tbl4, "Area", "Key Findings", "Relevance to Your Meeting", "Source")

    for i, item in enumerate(ENDEAVOR_SECTIONS, start=1):
        row = tbl4.add_row()
        fill = 'FFFFFF' if i % 2 == 0 else 'D6E4F0'
        cells = row.cells
        for c in cells:
            set_cell_bg(c, fill)
            c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        cell_para(cells[0], item["area"], bold=True, size=9)
        cell_para(cells[1], item["findings"], size=8.5)
        cell_para(cells[2], item["relevance"], size=8.5)
        add_hyperlink(cells[3].paragraphs[0], item["source_url"], item["source_text"], font_size=8.5)

    # ── SAVE ───────────────────────────────────────────
    path = "/home/user/Default-Test/Healthcare_Executive_Briefing_June2026.docx"
    doc.save(path)
    print(f"Document saved: {path}")
    return path


if __name__ == '__main__':
    build_document()
