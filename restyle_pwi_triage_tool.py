#!/usr/bin/env python3
"""Upgrade the slicer-based Principal Wellbeing Triage Tool workbook.

Takes the original draft 'Triage Tool.xlsx' (which contains working
slicers - not authorable by openpyxl) and performs XML-level surgery to:

* remove the "all eight slicers must be answered" requirement - unanswered
  questions are simply left out of the draft;
* build a complete email (greeting dropdown + recipient name + topic,
  thank-you line, body, contact details, sign-off) instead of just a body;
* update all wording to the June 2026 Ikon guidance and the team's email
  voice, and expand the common-response overrides from 2 to 8 recurring
  query types drawn from real replies;
* restyle the workbook in the Ocean Jade palette (blue 3C8BAC, jade 51B3B4,
  orange E5803A, yellow F4C55B), including theme accents so the slicers and
  tables pick up the scheme, plus a colour-coded triage result banner.

Usage: python3 restyle_pwi_triage_tool.py <source Triage Tool.xlsx> <output.xlsx>
"""

import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape

# ------------------------------------------------------------- palette ----
BLUE = "FF3C8BAC"
BLUE_DARK = "FF2E6E88"
JADE = "FF51B3B4"
JADE_PALE = "FFE6F4F3"
ORANGE = "FFE5803A"
YELLOW = "FFF4C55B"
SAND_PALE = "FFFCEBC8"

CONTACT = ("If you have any further questions, please reach out again. You can contact "
           "Principal wellbeing expenditure support by phone on 9264 4201 or by email at "
           "whsw.staffwellbeing@education.wa.edu.au.")

# --------------------------------------------------- shared string edits ----
Q3 = "Would this expenditure uphold Department values and standards, as per the Integrity framework and Code of Conduct?"
Q5 = "Does this service meet a need that cannot be provided for by current Department programs or services?"
Q6 = "Have all supplier-related conflicts of interest been identified, reported and appropriately managed (or do none apply)?"

SI_REPLACEMENTS = {
    5: "Q3. " + Q3,
    21: "Q3: " + Q3,
    104: Q3,
    7: "Q5. " + Q5,
    23: "Q5: " + Q5,
    112: Q5,
    8: "Q6. " + Q6,
    24: "Q6: " + Q6,
    116: Q6,
    57: "Greeting & recipient name",
    59: "Pick a greeting and add a name, e.g. 'Good morning Damien,'",
    60: ("Select Yes, No or Maybe in the question slicers below - you do not need to answer "
         "every question; unanswered ones are simply left out of the draft. For recurring "
         "queries, pick a common response override instead. Then add a greeting, name and "
         "topic under 'Draft email reply' and copy the finished draft into Outlook. "
         "To clear a slicer, click the funnel-with-x button in its top right corner."),
    61: ("Edit the modular wording on the Response library tab and the recurring full responses "
         "on the Preset responses tab. The draft remains a starting point: add case-specific "
         "detail and confirm the final advice against current Ikon guidance before sending."),
    12: "Tip: to clear a slicer, select the small red 'X' funnel button in its top right corner.",
    # ---- Response library paragraphs (June 2026 guidance, email voice) ----
    97: "The proposed expenditure meets the first test of the decision-making framework: it is intended to directly benefit your health, wellbeing or both as a principal.",
    98: "The proposed expenditure does not appear to directly benefit your health or wellbeing as a principal. The targeted funding is to address your needs as a principal - it is not intended for staff wellbeing or general operating costs - so we recommend reconsidering this expenditure.",
    99: "It is not clear that this would directly benefit your health or wellbeing as a principal. We would be happy to talk this through with you before you proceed.",
    101: "It aligns with the priorities in the health and wellbeing strategies and the panel service categories (executive and leadership coaching, executive health assessment, and mental health support).",
    102: "It does not align with the priorities in the health and wellbeing strategies or the panel service categories, so we recommend reconsidering this expenditure.",
    103: "The alignment with the health and wellbeing strategies or the panel service categories is not clear - please check with us before proceeding.",
    105: "The expenditure would uphold Department values and standards, in line with our Integrity framework and Code of Conduct (including Standard 8 - use public resources responsibly).",
    106: "The expenditure does not appear to uphold Department values and standards under the Integrity framework and Code of Conduct, so we recommend reconsidering it.",
    107: "It is not clear that the expenditure would uphold Department values and standards under the Integrity framework and Code of Conduct - please contact us before proceeding.",
    109: "The provider is reputable, appropriately insured and qualified to deliver the service.",
    110: "The provider's qualifications, insurance or reputation cannot be confirmed, so please do not proceed until this is resolved. Information about managing provider risks is available in the Develop a staff health and wellbeing plan Ikon service.",
    111: "The provider's qualifications, insurance or reputation have not yet been confirmed. Where relevant, check the Australian Health Practitioner Regulation Agency (AHPRA) register and confirm the provider's insurance before proceeding.",
    113: "The service meets a need that current Department programs or services do not already provide for.",
    114: "Current Department programs or services already meet this need - for example, the Employee Assistance Program now provides up to 12 confidential counselling sessions per year. We recommend using the existing service.",
    115: "It is not clear whether current Department programs or services already meet this need. Please check the available services - including the Employee Assistance Program (up to 12 sessions per year) - before proceeding.",
    117: "All supplier-related conflicts of interest have been identified, reported and appropriately managed (or none apply).",
    118: "A supplier-related conflict of interest has not yet been reported or managed. Please report the conflict of interest before proceeding.",
    119: "A possible supplier-related conflict of interest has not been confirmed as managed. Please report and resolve this before proceeding.",
    120: "This service falls within the panel of wellbeing service providers' categories. Once the panel is live, we encourage engaging a provider through the panel; in the meantime, the principal wellbeing consultant can help you assess and access suitable services.",
    121: "This service sits outside the panel of wellbeing service providers. You can engage the provider through local procurement, following the rules for buying goods and services. For purchases under $50,000, the Buy goods and services under $50,000 Ikon service applies, and purchases of less than $5,000 are low-value procurements provided value for money has been considered.",
    122: "It is not yet clear whether this service aligns with the panel of wellbeing service providers. The panel is non-mandatory and local procurement remains available - please check with us if you are unsure.",
    123: "As the query involves travel or accommodation: the funding may be used for travel and accommodation where it is required to support access to an eligible principal wellbeing activity, in line with the Official Travel Policy and the associated Official Air Travel and Travel Accommodation, Meal and Other Expenses procedures. Travellers cannot approve their own travel - please seek approval from an Assistant Director of Education or above prior to booking. The funding is not intended to be used solely for travel or accommodation.",
    124: "",  # Travel|No adds nothing to the email
    11: "Does this involve travel or accommodation?",
    128: "Edit the response names or the yellow response text below. These options appear as buttons on the common response override slicer on the Interactive triage tab.",
}

# --------------------------------------------------------- preset bodies ----
NL = "\n\n"
PRESETS = [
    ("Treatments / physiotherapy",
     "To answer briefly: no - personal health care, treatments and therapies (including physiotherapy and chiropractic care) are not within the scope of the targeted wellbeing funding."
     + NL + "To answer more thoroughly:"
     + NL + "The Workers' Compensation and Injury Management Act 2023 (WA) prohibits employers from diverting workers away from making a claim, so wellbeing funds cannot be used for injury-related assessment or treatment services. Treatment for work-related pain or injury may be available through the workers' compensation process, and may include payment for expenses, loss of earnings or both. Personal medical expenses are considered a fringe benefit."
     + NL + "While this service falls outside the scope of the funding, there are a range of other services you may wish to consider, including professional learning and resilience training, mental health support or pastoral care, coaching or mentoring (including executive and wellbeing coaching), and executive health and wellbeing assessments."),
    ("Memberships, gym or devices",
     "To answer briefly: no - memberships (such as gym or Pilates memberships), goods and personal devices are for the personal benefit of the employee, are considered a fringe benefit, and are not within the scope of the targeted wellbeing funding."
     + NL + "You may instead wish to consider services such as professional learning and resilience training, mental health support or pastoral care, coaching or mentoring (including executive and wellbeing coaching), or an executive health and wellbeing assessment."),
    ("Travel & accommodation",
     "To answer briefly: yes - the wellbeing funding can be used toward travel and accommodation needed to support access to an eligible wellbeing activity (such as professional learning, a professional conference or similar)."
     + NL + "To answer more thoroughly:"
     + NL + "The funding may be used for travel and accommodation where the travel is required to support the wellbeing activity. It is not intended to be used solely for travel or accommodation."
     + NL + "Travel needs to align with the Official Travel Policy and the associated Official Air Travel and Travel Accommodation, Meal and Other Expenses procedures. Travellers cannot approve their own travel: per the Travel Approval Schedule, please seek approval from an Assistant Director of Education or above prior to booking. For domestic air travel, the domestic air travel application form must be completed and approved before booking."),
    ("Conference / professional learning",
     "To answer briefly: yes - conferences and professional learning are considered appropriate expenditure where the majority of the program focuses on enhancing leadership capability and includes wellbeing topics for executives."
     + NL + "To answer more thoroughly:"
     + NL + "Suitable themes include work-related stress reduction, managing workload and role overload, resilience skills for leaders, decision-making under stress, and high-performance leadership without burnout."
     + NL + "The targeted funding may complement your school's budget allocation for the principal's professional learning, and relevant travel costs, but it should not replace it. If travel is required to attend, it needs to align with the Official Travel Policy and associated procedures, with approval from an Assistant Director of Education or above prior to booking."),
    ("Pooling funding with schools",
     "To answer briefly: yes - you may choose to pool your wellbeing funding with other schools, such as those in your school network, to access higher cost support services (for example, leadership or wellbeing coaching)."
     + NL + "To answer more thoroughly:"
     + NL + "The decision-making framework applies to pooled funding in the same way as individual funding. You can choose an appropriate budget account or accounts to record the transfer of funds when pooling resources."
     + NL + "For purchases under $50,000, follow the Buy goods and services under $50,000 Ikon service - the Verbal quotation template or Very simple purchase template may be useful. Purchases of less than $5,000 are considered low-value procurements and a panel is not required, provided value for money has been considered."),
    ("Recording purchases / acquittal",
     "To answer briefly: expenditure only needs to be lodged through Oracle, and you can choose an appropriate budget account or accounts to record purchases against."
     + NL + "We do not monitor principal wellbeing expenditure, and you do not need to provide formal acquittal reports. We do encourage you to keep a record of your decisions to demonstrate accountable and ethical decision-making - the Manage school accounting records Ikon service provides guidance on record keeping."),
    ("Ergonomic equipment",
     "To answer briefly: ergonomic assessments and equipment are provided for in existing school funding, so they sit outside the targeted wellbeing funding."
     + NL + "For guidance, refer to the Manage workstation ergonomics and Work comfortably at your computer Ikon services."),
    ("Catering",
     "To answer briefly: the funding may be used for catering where the expenditure complies with the Expenditure on Hospitality procedures and is related to a principal wellbeing activity."
     + NL + "Please note that catering alone is not a principal wellbeing activity - it needs to be connected to an eligible wellbeing activity."),
]

TRAVEL_MAYBE = ("It is not yet clear whether travel or accommodation will be involved. If it is, "
                "it may be funded where it supports the eligible wellbeing activity, in line with "
                "the Official Travel Policy - please check the approval requirements with us "
                "before booking.")

NEW_LINKS = [
    ("Official Air Travel procedures", "https://www.education.wa.edu.au/web/policies/-/official-air-travel-procedures"),
    ("Travel Accommodation, Meal and Other Expenses procedures", "https://www.education.wa.edu.au/web/policies/-/travel-accommodation-meal-and-other-expenses-procedures"),
    ("Tenders WA (panel tender information for providers)", "https://www.tenders.wa.gov.au/"),
]


def esc(s):
    return escape(s)


def xml_text(s):
    """Escape and encode newlines for a <t> element."""
    return escape(s).replace("\n", "&#10;")


def f_esc(formula):
    """Escape a formula for a <f> element."""
    return escape(formula)


def sub_once(pattern, repl, text, part):
    new, n = re.subn(pattern, repl, text, count=1, flags=re.S)
    if n != 1:
        raise RuntimeError(f"pattern not found in {part}: {pattern[:80]}")
    return new


def main(src, out):
    work = tempfile.mkdtemp()
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        z.extractall(work)
    p = lambda rel: os.path.join(work, rel)

    def read(rel):
        with open(p(rel), encoding="utf-8") as f:
            return f.read()

    def write(rel, text):
        with open(p(rel), "w", encoding="utf-8") as f:
            f.write(text)

    # ---------------------------------------------------- shared strings --
    ss = read("xl/sharedStrings.xml")
    items = re.split(r"(<si>.*?</si>)", ss, flags=re.S)
    si_i = -1
    for j, chunk in enumerate(items):
        if chunk.startswith("<si>"):
            si_i += 1
            if si_i in SI_REPLACEMENTS:
                items[j] = f'<si><t xml:space="preserve">{xml_text(SI_REPLACEMENTS[si_i])}</t></si>'
    write("xl/sharedStrings.xml", "".join(items))

    # ------------------------------------------- sheet1: Interactive triage
    s1 = read("xl/worksheets/sheet1.xml")

    # Row 33: split D33:F33 into greeting (D33) + name (E33:F33)
    s1 = sub_once(
        r'<c r="D33" s="14" t="s"><v>58</v></c>',
        '<c r="D33" s="14" t="inlineStr"><is><t>Good morning</t></is></c>',
        s1, "sheet1 D33")

    # Insert row 34 (topic of the query) after row 33
    row34 = ('<row r="34" spans="2:9" ht="21.95" customHeight="1" x14ac:dyDescent="0.25">'
             '<c r="B34" s="13" t="inlineStr"><is><t>Topic of the query (optional)</t></is></c>'
             '<c r="C34" s="13"/><c r="D34" s="14"/><c r="E34" s="14"/><c r="F34" s="14"/>'
             '<c r="G34" s="11" t="inlineStr"><is><t>Completes &quot;Thank you for reaching out '
             'regarding ...&quot;, e.g. &quot;using the funding for physiotherapy&quot;</t></is></c>'
             '<c r="H34" s="11"/><c r="I34" s="11"/></row>')
    m = re.search(r'<row r="33".*?</row>', s1, re.S)
    s1 = s1[:m.end()] + row34 + s1[m.end():]

    # Merges: D33:F33 -> D33 alone + E33:F33; add row 34 merges
    s1 = sub_once(r'<mergeCell ref="D33:F33"/>',
                  '<mergeCell ref="E33:F33"/><mergeCell ref="B34:C34"/>'
                  '<mergeCell ref="D34:F34"/><mergeCell ref="G34:I34"/>',
                  s1, "sheet1 merge D33")
    s1 = sub_once(r'<mergeCells count="11">', '<mergeCells count="14">', s1, "sheet1 merge count")

    # B28: triage result, partial-friendly
    b28 = ('IF(\'Preset responses\'!$F$2<>"","Common response override selected: "'
           "&'Preset responses'!$E$2,"
           'IF(COUNTIF(\'Pivot source data\'!$J$40:$Q$40,"?*")=0,'
           '"Waiting for answers - select Yes, No or Maybe in any question slicer, or pick a common response override.",'
           'IF(COUNTIF(\'Pivot source data\'!$J$40:$O$40,"No")>0,'
           '"Likely NOT appropriate - recommend reconsidering or contacting the team.",'
           'IF(COUNTIF(\'Pivot source data\'!$J$40:$O$40,"Yes")=6,'
           '"Likely appropriate to proceed.",'
           '"Based on the answers so far, more information is needed - unanswered questions are left out of the draft."))))')
    s1 = sub_once(r'(<c r="B28" s="18" t="str"><f>).*?(</f>)<v>.*?</v>',
                  lambda mm: mm.group(1) + f_esc(b28) + mm.group(2), s1, "sheet1 B28")

    # B35: full draft email
    body = ('IF(\'Preset responses\'!$F$2<>"",\'Preset responses\'!$F$2,'
            'IF(COUNTIF(\'Pivot source data\'!$J$40:$Q$40,"?*")=0,'
            '"[Select an answer in one or more question slicers above, or pick a common response override, to build the body of the email.]",'
            'IF(\'Pivot source data\'!$AA$40="",'
            '"[The answers selected so far do not add any wording to the email - answer more of the question slicers above.]",'
            "'Pivot source data'!$AA$40)))")
    b35 = ('TRIM(IF($D$33="","Hi",$D$33)&" "&TRIM($E$33))&","&CHAR(10)&CHAR(10)&'
           'IF(TRIM($D$34)="","Thank you for reaching out regarding the principal wellbeing targeted funding.",'
           '"Thank you for reaching out regarding "&TRIM($D$34)&".")&CHAR(10)&CHAR(10)&'
           + body +
           '&CHAR(10)&CHAR(10)&"' + CONTACT + '"&CHAR(10)&CHAR(10)&"Kind regards,"')
    s1 = sub_once(r'(<c r="B35" s="15" t="str"><f>).*?(</f>)<v>.*?</v>',
                  lambda mm: mm.group(1) + f_esc(b35) + mm.group(2), s1, "sheet1 B35")

    # Taller rows 23-26 so the 9-button common-response slicer fits (3 x 3)
    for r in (23, 24, 25, 26):
        s1 = re.sub(rf'<row r="{r}" spans="2:9" ht="18"',
                    f'<row r="{r}" spans="2:9" ht="26.25"', s1, count=1)

    # Conditional formatting for the result band + greeting dropdown
    cf = ('<conditionalFormatting sqref="B28:I30">'
          '<cfRule type="expression" dxfId="0" priority="1"><formula>ISNUMBER(SEARCH("Likely appropriate",$B$28))</formula></cfRule>'
          '<cfRule type="expression" dxfId="1" priority="2"><formula>ISNUMBER(SEARCH("NOT appropriate",$B$28))</formula></cfRule>'
          '<cfRule type="expression" dxfId="2" priority="3"><formula>ISNUMBER(SEARCH("more information",$B$28))</formula></cfRule>'
          '<cfRule type="expression" dxfId="3" priority="4"><formula>ISNUMBER(SEARCH("Common response",$B$28))</formula></cfRule>'
          '</conditionalFormatting>'
          '<dataValidations count="1">'
          '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D33">'
          '<formula1>"Good morning,Good afternoon,Good evening,Hi"</formula1></dataValidation>'
          '</dataValidations>')
    s1 = sub_once(r'</mergeCells>', '</mergeCells>' + cf, s1, "sheet1 cf insert")
    write("xl/worksheets/sheet1.xml", s1)

    # --------------------------------------------- sheet4: Pivot source data
    s4 = read("xl/worksheets/sheet4.xml")
    z40 = ('IF(COUNTIF($J$40:$O$40,"?*")=0,"",'
           'IF(COUNTIF($J$40:$O$40,"No")>0,'
           '"To answer briefly: based on the decision-making framework, this expenditure does not appear to be appropriate. Please see the detail below.",'
           'IF(COUNTIF($J$40:$O$40,"Yes")=6,'
           '"To answer briefly: based on the decision-making framework, this expenditure is likely to be appropriate. Please see the detail below.",'
           '"To answer briefly: we need a little more information before this expenditure can be confirmed as appropriate. Please see the detail below.")))')
    s4 = sub_once(r'(<c r="Z40" t="str"><f>).*?(</f>)<v/>',
                  lambda mm: mm.group(1) + f_esc(z40) + mm.group(2), s4, "sheet4 Z40")
    aa40 = ('IF(COUNTIF($J$40:$Q$40,"?*")=0,"",'
            '_xlfn.TEXTJOIN(CHAR(10)&CHAR(10),TRUE(),$Z$40,'
            'IF(COUNTIF($J$40:$O$40,"?*")=0,"","To answer more thoroughly:"),$R$40:$Y$40))')
    s4 = sub_once(r'(<c r="AA40" t="str"><f>).*?(</f>)<v/>',
                  lambda mm: mm.group(1) + f_esc(aa40) + mm.group(2), s4, "sheet4 AA40")
    # Guard Travel|No empty paragraph (INDEX would return 0); range covers the
    # Travel|Maybe row added to the Response library below
    y40 = ('IF($Q$40="","",IFERROR(T(INDEX(\'Response library\'!$D$2:$D$25,'
           'MATCH("Travel|"&$Q$40,\'Response library\'!$E$2:$E$25,0))),""))')
    s4 = sub_once(r'(<c r="Y40" t="str"><f>).*?(</f>)<v/>',
                  lambda mm: mm.group(1) + f_esc(y40) + mm.group(2), s4, "sheet4 Y40")
    # Give the Travel slicer table a third option (Maybe) so it is structurally
    # identical to the seven other question tables
    s4 = sub_once(r'(<row r="39"[^>]*>)(<c r="J39")',
                  r'\1<c r="A39" t="s"><v>67</v></c><c r="B39"><f>SUBTOTAL(103,A39)</f></c>\2',
                  s4, "sheet4 A39/B39")
    s4 = sub_once(
        r'<f>IF\(SUM\(\$B\$37:\$B\$38\)=1,INDEX\(\$A\$37:\$A\$38,MATCH\(1,\$B\$37:\$B\$38,0\)\),""\)</f>',
        '<f>IF(SUM($B$37:$B$39)=1,INDEX($A$37:$A$39,MATCH(1,$B$37:$B$39,0)),"")</f>',
        s4, "sheet4 Q40 range")
    # drop stale cached results so nothing shows pre-upgrade values before recalc
    s4 = re.sub(r'(</f>)<v(?: [^>]*)?>[^<]*</v>', r"\1", s4).replace("</f><v/>", "</f>")
    write("xl/worksheets/sheet4.xml", s4)

    # --------------------------------------------- sheet7: Preset responses
    s7 = read("xl/worksheets/sheet7.xml")
    s7 = s7.replace('<dimension ref="A1:F8"/>', '<dimension ref="A1:F14"/>')
    # drop stale cached selection results
    s7 = re.sub(r'(<c r="E2" s="5" t="str"><f>.*?</f>)<v>.*?</v>', r'\1', s7, flags=re.S)
    s7 = re.sub(r'(<c r="F2" s="5" t="str"><f>.*?</f>)<v>.*?</v>', r'\1', s7, flags=re.S)
    # unhide row 6 and rebuild preset rows 7..14
    s7 = sub_once(r'<row r="6" spans="1:6" ht="36" hidden="1"',
                  '<row r="6" spans="1:6" ht="36"', s7, "sheet7 row6")
    rows = []
    for i, (name, bodytext) in enumerate(PRESETS):
        r = 7 + i
        lines = sum(len(pp) // 100 + 1 for pp in bodytext.split("\n\n")) + bodytext.count("\n\n")
        ht = max(40, 13 * lines + 8)
        rows.append(
            f'<row r="{r}" spans="1:6" ht="{ht}" customHeight="1" x14ac:dyDescent="0.25">'
            f'<c r="A{r}" s="5" t="inlineStr"><is><t xml:space="preserve">{xml_text(name)}</t></is></c>'
            f'<c r="B{r}" s="10" t="inlineStr"><is><t xml:space="preserve">{xml_text(bodytext)}</t></is></c>'
            f'<c r="C{r}" s="5"><f>SUBTOTAL(103,A{r})</f></c></row>')
    s7 = sub_once(r'<row r="7" spans="1:6".*?</row><row r="8" spans="1:6".*?</row>',
                  "".join(rows), s7, "sheet7 preset rows")
    # Excel repairs the workbook if a table header cell differs from the
    # tableColumn name ("Draft email body") - a defect inherited from the
    # source draft. The editing hint stays in the instruction row above.
    s7 = sub_once(r'<c r="B5" s="4" t="s"><v>131</v></c>',
                  '<c r="B5" s="4" t="inlineStr"><is><t>Draft email body</t></is></c>',
                  s7, "sheet7 B5 header")
    write("xl/worksheets/sheet7.xml", s7)

    # table9: extend to the 8 presets and clear the saved slicer filter
    t9 = read("xl/tables/table9.xml")
    t9 = t9.replace('ref="A5:C8"', 'ref="A5:C14"')
    t9 = re.sub(r'<filterColumn.*?</filterColumn>', '', t9, flags=re.S)
    write("xl/tables/table9.xml", t9)

    # table8: include the new Maybe row
    t8 = read("xl/tables/table8.xml")
    t8 = t8.replace('ref="A36:B38"', 'ref="A36:B39"')
    write("xl/tables/table8.xml", t8)

    # ------------------------------------------ sheet6: Response library
    # Add the Travel|Maybe row so every question offers Yes/No/Maybe
    s6 = read("xl/worksheets/sheet6.xml")
    s6 = s6.replace('<dimension ref="A1:E24"/>', '<dimension ref="A1:E25"/>')
    s6 = s6.replace('<autoFilter ref="A1:E24"', '<autoFilter ref="A1:E25"')
    row25 = ('<row r="25" spans="1:5" ht="48" customHeight="1" x14ac:dyDescent="0.25">'
             '<c r="A25" s="5" t="s"><v>74</v></c><c r="B25" s="5" t="s"><v>11</v></c>'
             '<c r="C25" s="6" t="s"><v>67</v></c>'
             f'<c r="D25" s="10" t="inlineStr"><is><t xml:space="preserve">{xml_text(TRAVEL_MAYBE)}</t></is></c>'
             '<c r="E25" s="5"><f>A25&amp;"|"&amp;C25</f></c></row>')
    m6 = re.search(r'<row r="24".*?</row>', s6, re.S)
    s6 = s6[:m6.end()] + row25 + s6[m6.end():]
    write("xl/worksheets/sheet6.xml", s6)

    # --------------------------------------------- sheet5: Reference links
    s5 = read("xl/worksheets/sheet5.xml")
    s5 = s5.replace('<dimension ref="B2:C16"/>', '<dimension ref="B2:C19"/>')
    row16 = re.search(r'<row r="16".*?</row>', s5, re.S).group(0)
    sb = re.search(r'<c r="B16" s="(\d+)"', row16).group(1)
    sc = re.search(r'<c r="C16" s="(\d+)"', row16).group(1)
    newrows, hlinks = [], []
    rels = read("xl/worksheets/_rels/sheet5.xml.rels")
    next_rid = max(int(n) for n in re.findall(r'Id="rId(\d+)"', rels)) + 1
    for i, (topic, url) in enumerate(NEW_LINKS):
        r, rid = 17 + i, next_rid + i
        newrows.append(
            f'<row r="{r}" spans="2:3" x14ac:dyDescent="0.25">'
            f'<c r="B{r}" s="{sb}" t="inlineStr"><is><t xml:space="preserve">{xml_text(topic)}</t></is></c>'
            f'<c r="C{r}" s="{sc}" t="inlineStr"><is><t xml:space="preserve">{xml_text(url)}</t></is></c></row>')
        hlinks.append(f'<hyperlink ref="C{r}" r:id="rId{rid}"/>')
        rels = rels.replace('</Relationships>',
                            f'<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/'
                            f'officeDocument/2006/relationships/hyperlink" Target="{esc(url)}" '
                            f'TargetMode="External"/></Relationships>')
    s5 = sub_once(r'</sheetData>', "".join(newrows) + '</sheetData>', s5, "sheet5 rows")
    s5 = sub_once(r'</hyperlinks>', "".join(hlinks) + '</hyperlinks>', s5, "sheet5 hyperlinks")
    write("xl/worksheets/sheet5.xml", s5)
    write("xl/worksheets/_rels/sheet5.xml.rels", rels)

    # --------------------------------------------------------- workbook ----
    wbx = read("xl/workbook.xml")
    wbx = wbx.replace('<sheet name="Select answers" sheetId="1" r:id="rId2"/>',
                      '<sheet name="Select answers" sheetId="1" state="hidden" r:id="rId2"/>')
    wbx = wbx.replace('<sheet name="Draft reply" sheetId="2" r:id="rId3"/>',
                      '<sheet name="Draft reply" sheetId="2" state="hidden" r:id="rId3"/>')
    if "<calcPr" in wbx:
        wbx = re.sub(r'<calcPr([^>]*?)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wbx, count=1)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    write("xl/workbook.xml", wbx)

    # drop the stale calcChain (Excel rebuilds it)
    if os.path.exists(p("xl/calcChain.xml")):
        os.remove(p("xl/calcChain.xml"))
        ct = read("[Content_Types].xml")
        ct = ct.replace('<Override PartName="/xl/calcChain.xml" ContentType="application/'
                        'vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>', '')
        write("[Content_Types].xml", ct)
        wrels = read("xl/_rels/workbook.xml.rels")
        wrels = re.sub(r'<Relationship Id="[^"]+" Type="[^"]*calcChain[^"]*" Target="calcChain.xml"/>',
                       '', wrels)
        write("xl/_rels/workbook.xml.rels", wrels)

    # ------------------------------------------------------ Ocean Jade ----
    st = read("xl/styles.xml")
    # fills: old dark-teal scheme -> ocean jade
    fill_map = {"FF007A78": BLUE[2:], "FF008080": BLUE[2:], "FF00524F": JADE[2:],
                "FFE7F3F2": JADE_PALE[2:], "FFFFF2CC": SAND_PALE[2:]}

    def remap_fills(mm):
        block = mm.group(0)
        for old, new in fill_map.items():
            block = block.replace(old, "FF" + new)
        return block
    st = re.sub(r'<fill>.*?</fill>', remap_fills, st, flags=re.S)
    # accent text colour: dark teal -> dark ocean blue
    st = st.replace('<color rgb="FF00524F"/>', f'<color rgb="{BLUE_DARK}"/>')
    # dxfs for the result band conditional formats
    dxfs = ('<dxfs count="4">'
            f'<dxf><font><color rgb="FF1F6B6C"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFDCF0EF"/></patternFill></fill></dxf>'
            f'<dxf><font><color rgb="FF9C4511"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFAE3D2"/></patternFill></fill></dxf>'
            f'<dxf><font><color rgb="FF8A6116"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFFCF0D2"/></patternFill></fill></dxf>'
            f'<dxf><font><color rgb="{BLUE_DARK}"/></font><fill><patternFill patternType="solid"><bgColor rgb="FFE1EFF6"/></patternFill></fill></dxf>'
            '</dxfs>')
    if "<dxfs" in st:
        st = re.sub(r'<dxfs count="0"/>', dxfs, st, count=1)
    else:
        st = st.replace("<tableStyles", dxfs + "<tableStyles")
    write("xl/styles.xml", st)

    # theme accents drive the slicer / table styling
    th = read("xl/theme/theme1.xml")
    for old, new in [("4F81BD", BLUE[2:]), ("C0504D", JADE[2:]),
                     ("9BBB59", ORANGE[2:]), ("8064A2", YELLOW[2:])]:
        th = th.replace(f'val="{old}"', f'val="{new}"')
    write("xl/theme/theme1.xml", th)

    # tab colours
    tabs = {"sheet1": BLUE, "sheet5": JADE}
    for sheet, colour in tabs.items():
        rel = f"xl/worksheets/{sheet}.xml"
        x = read(rel)
        x = re.sub(r'<tabColor rgb="[0-9A-F]+"/>', f'<tabColor rgb="{colour}"/>', x, count=1)
        write(rel, x)
    for sheet, colour in {"sheet6": ORANGE, "sheet7": YELLOW}.items():
        rel = f"xl/worksheets/{sheet}.xml"
        x = read(rel)
        if "<sheetPr>" not in x:
            x = re.sub(r'(<worksheet[^>]*>)', rf'\1<sheetPr><tabColor rgb="{colour}"/></sheetPr>',
                       x, count=1)
        write(rel, x)

    # ------------------------------------------------------------ rezip ----
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:
            full = p(name)
            if os.path.exists(full):
                z.write(full, name)
    shutil.rmtree(work)
    print(f"Saved {out}")


if __name__ == "__main__":
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "Principal wellbeing email triage tool.xlsx"
    main(src, out)
