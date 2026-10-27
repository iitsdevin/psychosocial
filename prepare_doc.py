#!/usr/bin/env python3
"""
Prepare the Level 1 PHIR Response for the PEEPO macro.

- Clears the answer cell under "Outline any relevant work or non-work
  factors" (two empty "Choose an item." dropdowns and an empty "Output Text"
  placeholder), leaving one blank paragraph for the macro to write into.
- Adds a blue "double-click here" line to the label cell that runs
  PEEPO_ShowChecklist once PEEPOModule.bas has been imported in Word.

Usage: python3 prepare_doc.py
"""

import copy
import os

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE_DOCX = os.path.join(HERE, "source", "Level1PHIRResponse_original.docx")
OUTPUT_DOCX = os.path.join(HERE, "Level1_PHIR_Response_ReadyForMacro.docx")

TARGET_LABEL = "Outline any relevant work or non-work factors"
BUTTON_CODE = " MACROBUTTON PEEPO_ShowChecklist [Double-click here to choose work factors from the PEEPO list] "


def find_label_row(doc):
    for table in doc.tables:
        rows = table._tbl.tr_lst
        for i, tr in enumerate(rows):
            text = "".join(t.text or "" for t in tr.iter(qn("w:t")))
            if TARGET_LABEL in text:
                return rows[i], rows[i + 1]
    raise SystemExit(f"Could not find the '{TARGET_LABEL}' row")


def button_run_props():
    rPr = OxmlElement("w:rPr")
    bold = OxmlElement("w:b")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "18")
    rPr.extend([bold, color, size])
    return rPr


def add_button(label_cell):
    label_para = label_cell.findall(qn("w:p"))[-1]
    para = OxmlElement("w:p")
    pPr = label_para.find(qn("w:pPr"))
    if pPr is not None:
        para.append(copy.deepcopy(pPr))

    for kind in ("begin", "instr", "end"):
        run = OxmlElement("w:r")
        run.append(button_run_props())
        if kind == "instr":
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = BUTTON_CODE
            run.append(instr)
        else:
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), kind)
            run.append(fld)
        para.append(run)
    label_para.addnext(para)


def clear_answer_cell(answer_cell):
    paras = answer_cell.findall(qn("w:p"))
    keep_pPr = paras[0].find(qn("w:pPr"))
    for child in paras + answer_cell.findall(qn("w:sdt")):
        answer_cell.remove(child)
    para = OxmlElement("w:p")
    if keep_pPr is not None:
        para.append(copy.deepcopy(keep_pPr))
    answer_cell.append(para)


def main():
    doc = docx.Document(SOURCE_DOCX)
    label_row, answer_row = find_label_row(doc)
    add_button(label_row.tc_lst[0])
    clear_answer_cell(answer_row.tc_lst[0])
    doc.save(OUTPUT_DOCX)
    print(f"Saved {OUTPUT_DOCX}")


if __name__ == "__main__":
    main()
