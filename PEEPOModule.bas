Attribute VB_Name = "PEEPOModule"
Option Explicit

' PEEPO work factor picker for the Level 1 PHIR Response.
'
' Double-click the blue "choose work factors" text in the Nature of Concern
' table (or run PEEPO_ShowChecklist from Alt+F8). A tick-box checklist of the
' PEEPO categories opens in the "Outline any relevant work or non-work
' factors" cell. Tick the factors that apply, then double-click the blue
' "insert the ticked factors" text (or run PEEPO_InsertTicked). The checklist
' is replaced with the ticked factors as plain text, grouped by category.
' This can be repeated as many times as needed; each run adds to the cell.

Private Const TARGET_LABEL As String = "Outline any relevant work or non-work factors"
Private Const PICKER_BOOKMARK As String = "PEEPO_Checklist"
Private Const TAG_PREFIX As String = "PEEPO|"

Private Function CategoryNames() As Variant
    CategoryNames = Array( _
        "People (Human Factors)", _
        "Equipment (Resources & equipment)", _
        "Environment (Physical and workplace environment)", _
        "Processes (Tasks & procedures)", _
        "Organisation (System and policy)")
End Function

Private Function CategoryFactors(ByVal category As Long) As Variant
    Dim s As String
    Select Case category
    Case 0
        s = "Inadequate communication between staff"
        s = s & "|Unreasonable behaviour by worker(s)"
        s = s & "|Unreasonable behaviour by manager(s) or supervisor(s)"
        s = s & "|Interpersonal conflict (unresolved or escalating)"
        s = s & "|Fatigue or burnout"
        s = s & "|Mental stress or psychological distress"
        s = s & "|Exposure to traumatic event or material"
        s = s & "|Personal issues affecting work capacity"
        s = s & "|Incivility or inappropriate behaviour tolerated by peers"
    Case 1
        s = "Insufficient staffing levels or resources to meet work demands"
        s = s & "|Inadequate or unavailable tools and resources to do the job"
        s = s & "|Insufficient training on psychosocial hazard identification or management"
        s = s & "|Insufficient training for job demands or new tasks"
        s = s & "|Lack of constructive feedback or performance support"
        s = s & "|Inadequate access to employee assistance or support services"
        s = s & "|Lack of information about reporting processes or support pathways"
        s = s & "|Insufficient or inappropriate supervision"
        s = s & "|Inadequate de-escalation tools or conflict resolution resources"
    Case 2
        s = "Adverse environmental conditions (noise, temperature, air quality)"
        s = s & "|Poorly maintained or unclean facilities"
        s = s & "|Workplace layout limiting supervision or support access"
        s = s & "|Remote or isolated work location"
        s = s & "|Natural events restricting travel or creating uncertainty"
        s = s & "|Limited access to communication technology"
        s = s & "|Blurring of boundaries between work and home life"
        s = s & "|Workplace culture tolerating unreasonable behaviour"
        s = s & "|Lack of diversity, respect, or inclusion in the workplace"
        s = s & "|Limited opportunities for social interaction during work"
    Case 3
        s = "No or inadequate policies for managing psychosocial hazards"
        s = s & "|Policies or procedures not adhered to"
        s = s & "|Procedures that cannot be applied as written or lack flexibility"
        s = s & "|No or inadequate reporting and complaints processes"
        s = s & "|Investigation or appeals process not affording natural justice"
        s = s & "|No or inadequate incident response or investigation procedures"
        s = s & "|Procedures that discriminate or lack fairness"
        s = s & "|Inadequate rostering, hours of work, or fatigue management procedures"
        s = s & "|Inadequate change management procedures"
        s = s & "|No mechanism for impartially addressing behaviour by senior management"
    Case 4
        s = "Poor leadership practices or management style"
        s = s & "|Limited management accountability for psychosocial hazards"
        s = s & "|Unclear or conflicting roles, responsibilities, or expectations"
        s = s & "|Excessive work demands (cognitive, emotional, or physical)"
        s = s & "|Workplace culture"
        s = s & "|Multiple interacting psychosocial hazards not considered together"
        s = s & "|Low job control or autonomy"
        s = s & "|Lack of recognition or reward"
        s = s & "|Organisational injustice (unfairness, bias, inconsistency)"
        s = s & "|Inadequate consultation with workers on matters affecting them"
        s = s & "|Poorly managed organisational change"
        s = s & "|Insecure work arrangements"
        s = s & "|Inadequate monitoring and review of existing controls"
    End Select
    CategoryFactors = Split(s, "|")
End Function

Public Sub PEEPO_ShowChecklist()
    Dim doc As Document
    Set doc = ActiveDocument

    If doc.Bookmarks.Exists(PICKER_BOOKMARK) Then
        MsgBox "The PEEPO checklist is already open in the table." & vbCr & vbCr & _
               "Tick the factors that apply, then double-click the blue " & _
               "'insert the ticked factors' text at the end of the list.", _
               vbInformation, "PEEPO"
        doc.Bookmarks(PICKER_BOOKMARK).Range.Characters(1).Select
        Exit Sub
    End If

    Dim answerCell As Cell
    Set answerCell = FindAnswerCell(doc)
    If answerCell Is Nothing Then
        MsgBox "Could not find the '" & TARGET_LABEL & "' row in the " & _
               "Nature of Concern table.", vbExclamation, "PEEPO"
        Exit Sub
    End If

    On Error GoTo Failed
    Application.ScreenUpdating = False

    ' Start the checklist on a new paragraph after anything already in the cell
    Dim rng As Range
    Set rng = answerCell.Range
    rng.MoveEnd wdCharacter, -1
    If Not IsBlank(rng.Text) Then
        rng.Collapse wdCollapseEnd
        rng.InsertParagraphAfter
    End If
    rng.Collapse wdCollapseEnd

    ' One paragraph per line; kinds records what each paragraph is for
    Dim cats As Variant, factors As Variant
    Dim kinds As New Collection
    Dim body As String
    Dim c As Long, f As Long
    cats = CategoryNames()
    body = "Tick the factors that apply, then double-click the blue text at the end of this list."
    kinds.Add "N"
    For c = 0 To UBound(cats)
        body = body & vbCr & cats(c)
        kinds.Add "H"
        factors = CategoryFactors(c)
        For f = 0 To UBound(factors)
            body = body & vbCr & " " & factors(f)
            kinds.Add "F" & c & "|" & f
        Next f
    Next c
    body = body & vbCr
    kinds.Add "B"
    rng.InsertAfter body

    Dim region As Range
    Set region = doc.Range(rng.Start, answerCell.Range.End - 1)
    region.Font.Reset
    region.Font.Size = 9
    region.ParagraphFormat.SpaceBefore = 0
    region.ParagraphFormat.SpaceAfter = 0

    Dim i As Long
    Dim kind As String
    Dim p As Range
    Dim cc As ContentControl
    For i = 1 To kinds.Count
        kind = kinds(i)
        Set p = region.Paragraphs(i).Range
        Select Case Left$(kind, 1)
        Case "N"
            p.Font.Italic = True
            p.Font.Color = wdColorGray50
        Case "H"
            p.Font.Bold = True
            p.ParagraphFormat.SpaceBefore = 6
        Case "F"
            p.Collapse wdCollapseStart
            Set cc = doc.ContentControls.Add(wdContentControlCheckBox, p)
            cc.Tag = TAG_PREFIX & Mid$(kind, 2)
        Case "B"
            p.ParagraphFormat.SpaceBefore = 6
            p.Collapse wdCollapseStart
            doc.Fields.Add p, wdFieldEmpty, _
                "MACROBUTTON PEEPO_InsertTicked [Double-click here to insert the ticked factors]", False
            Set p = region.Paragraphs(i).Range
            p.Font.Color = wdColorBlue
            p.Font.Bold = True
        End Select
    Next i

    doc.Bookmarks.Add PICKER_BOOKMARK, doc.Range(region.Start, answerCell.Range.End - 1)
    Application.ScreenUpdating = True
    doc.Range(region.Start, region.Start).Select
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not open the PEEPO checklist: " & Err.Description, vbExclamation, "PEEPO"
End Sub

Public Sub PEEPO_InsertTicked()
    Dim doc As Document
    Set doc = ActiveDocument

    If Not doc.Bookmarks.Exists(PICKER_BOOKMARK) Then
        MsgBox "The PEEPO checklist is not open." & vbCr & vbCr & _
               "Double-click the blue 'choose work factors' text in the " & _
               "Nature of Concern table first.", vbInformation, "PEEPO"
        Exit Sub
    End If

    Dim region As Range
    Set region = doc.Bookmarks(PICKER_BOOKMARK).Range

    ' Collect the ticked factors in checklist order, grouped by category
    Dim cats As Variant, factors As Variant, parts As Variant
    Dim cc As ContentControl
    Dim summary As String
    Dim lastCat As Long, ticked As Long
    cats = CategoryNames()
    lastCat = -1
    For Each cc In region.ContentControls
        If cc.Type = wdContentControlCheckBox And Left$(cc.Tag, Len(TAG_PREFIX)) = TAG_PREFIX Then
            If cc.Checked Then
                parts = Split(Mid$(cc.Tag, Len(TAG_PREFIX) + 1), "|")
                If CLng(parts(0)) <> lastCat Then
                    lastCat = CLng(parts(0))
                    If Len(summary) > 0 Then summary = summary & vbCr
                    summary = summary & cats(lastCat) & ":"
                End If
                factors = CategoryFactors(lastCat)
                summary = summary & vbCr & ChrW(8226) & " " & factors(CLng(parts(1)))
                ticked = ticked + 1
            End If
        End If
    Next cc

    If ticked = 0 Then
        If MsgBox("No factors are ticked." & vbCr & vbCr & _
                  "Close the checklist without adding anything?", _
                  vbYesNo + vbQuestion, "PEEPO") = vbNo Then Exit Sub
    End If

    On Error GoTo Failed
    Application.ScreenUpdating = False

    Dim startPos As Long, i As Long
    startPos = region.Start
    doc.Bookmarks(PICKER_BOOKMARK).Delete
    For i = region.ContentControls.Count To 1 Step -1
        region.ContentControls(i).Delete True
    Next i
    region.Delete

    Dim outRng As Range
    Set outRng = doc.Range(startPos, startPos)
    If ticked > 0 Then
        outRng.InsertAfter summary
        outRng.Font.Reset
        Dim para As Paragraph
        For Each para In outRng.Paragraphs
            para.Range.Font.Bold = (Left$(para.Range.Text, 1) <> ChrW(8226))
        Next para
    ElseIf startPos > outRng.Cells(1).Range.Start Then
        ' Remove the empty paragraph the checklist was added on
        doc.Range(startPos - 1, startPos).Delete
    End If

    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not insert the ticked factors: " & Err.Description, vbExclamation, "PEEPO"
End Sub

Private Function FindAnswerCell(ByVal doc As Document) As Cell
    Dim tbl As Table
    Dim c As Cell
    For Each tbl In doc.Tables
        For Each c In tbl.Range.Cells
            If InStr(1, c.Range.Text, TARGET_LABEL, vbTextCompare) > 0 Then
                ' The answer cell is the cell straight after the label
                Set FindAnswerCell = c.Next
                Exit Function
            End If
        Next c
    Next tbl
End Function

Private Function IsBlank(ByVal s As String) As Boolean
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(11), "")
    s = Replace(s, vbTab, "")
    IsBlank = (Len(Trim$(s)) = 0)
End Function
