# psychosocial

## PEEPO work factor picker (Level 1 PHIR Response)

`Level1_PHIR_Response_ReadyForMacro.docx` has a blue "Double-click here to
choose work factors from the PEEPO list" line under "Outline any relevant work
or non-work factors". `PEEPOModule.bas` is the macro behind it.

### One-off setup (done once, by whoever owns the template)

1. Open `Level1_PHIR_Response_ReadyForMacro.docx` in Word.
2. Press **Alt+F11** to open the VBA editor.
3. In the Project pane, click the project for this document
   (`Project (Level1_PHIR_Response_ReadyForMacro)`).
4. **File > Import File...** and choose `PEEPOModule.bas`.
5. Close the VBA editor.
6. **File > Save As** and choose **Word Macro-Enabled Document (*.docm)**.

Share that `.docm`. Nobody else needs to do the setup.

### Using it

1. Double-click the blue "choose work factors" line. A checklist of the five
   PEEPO categories opens in the cell below.
2. Tick the factors that apply.
3. Double-click the blue "insert the ticked factors" line at the end of the
   list. The checklist is replaced with the ticked factors as text, grouped
   by category.

Repeat as often as needed; each run adds to the cell. The macros can also be
run from **Alt+F8**: `PEEPO_ShowChecklist` and `PEEPO_InsertTicked`.

If Word shows a red "Microsoft has blocked macros" bar, the file came from
email or the internet. Close it, right-click the file > **Properties**, tick
**Unblock**, then reopen.

### Rebuilding the prepared document

`python3 prepare_doc.py` rebuilds `Level1_PHIR_Response_ReadyForMacro.docx`
from `source/Level1PHIRResponse_original.docx`.
