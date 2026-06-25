Attribute VB_Name = "PEEPOModule"
Option Explicit

' ---------------------------------------------------------------
'  PEEPO Work-Factor Picker
'  Run this macro (Alt+F8 > SelectPEEPOFactors) to select
'  psychosocial factors and insert them into the
'  "Outline any relevant work or non-work factors" cell.
' ---------------------------------------------------------------

Public Sub SelectPEEPOFactors()
    Dim cats(1 To 5) As String
    cats(1) = "People (Human Factors)"
    cats(2) = "Equipment (Resources & Equipment)"
    cats(3) = "Environment (Physical & Workplace)"
    cats(4) = "Processes (Tasks & Procedures)"
    cats(5) = "Organisation (System & Policy)"

    Dim menuText As String
    menuText = "SELECT A CATEGORY" & vbCrLf & vbCrLf
    menuText = menuText & "1 - People (Human Factors)" & vbCrLf
    menuText = menuText & "2 - Equipment (Resources & Equipment)" & vbCrLf
    menuText = menuText & "3 - Environment (Physical & Workplace)" & vbCrLf
    menuText = menuText & "4 - Processes (Tasks & Procedures)" & vbCrLf
    menuText = menuText & "5 - Organisation (System & Policy)" & vbCrLf
    menuText = menuText & vbCrLf & "Enter number (or 0 when finished):"

    Dim allSelected As String
    allSelected = ""

    Do
        Dim catChoice As String
        catChoice = InputBox(menuText, "PEEPO Factor Selection")

        If catChoice = "" Or catChoice = "0" Then Exit Do

        Dim catNum As Long
        catNum = Val(catChoice)

        If catNum < 1 Or catNum > 5 Then
            MsgBox "Please enter a number between 1 and 5.", vbExclamation
        Else
            Dim factorList As String
            factorList = GetFactorsForCategory(catNum)

            Dim selected As String
            selected = InputBox(factorList, cats(catNum))

            If selected <> "" Then
                Dim result As String
                result = ParseSelection(catNum, selected)
                If result <> "" Then
                    If allSelected <> "" Then
                        allSelected = allSelected & vbCrLf
                    End If
                    allSelected = allSelected & cats(catNum) & ":" & vbCrLf & result
                End If
            End If
        End If
    Loop

    If allSelected <> "" Then
        InsertFactorsIntoTable allSelected
    End If
End Sub

Private Function GetFactorsForCategory(catNum As Long) As String
    Dim s As String
    s = "Enter the numbers of applicable factors," & vbCrLf
    s = s & "separated by commas (e.g. 1,3,5):" & vbCrLf & vbCrLf

    Select Case catNum
    Case 1
        s = s & "1 - Inadequate communication between staff" & vbCrLf
        s = s & "2 - Unreasonable behaviour by worker(s)" & vbCrLf
        s = s & "3 - Unreasonable behaviour by manager(s) or supervisor(s)" & vbCrLf
        s = s & "4 - Interpersonal conflict (unresolved or escalating)" & vbCrLf
        s = s & "5 - Fatigue or burnout" & vbCrLf
        s = s & "6 - Mental stress or psychological distress" & vbCrLf
        s = s & "7 - Exposure to traumatic event or material" & vbCrLf
        s = s & "8 - Personal issues affecting work capacity" & vbCrLf
        s = s & "9 - Incivility or inappropriate behaviour tolerated by peers"
    Case 2
        s = s & "1 - Insufficient staffing levels or resources" & vbCrLf
        s = s & "2 - Inadequate or unavailable tools and resources" & vbCrLf
        s = s & "3 - Insufficient training on psychosocial hazard ID or mgmt" & vbCrLf
        s = s & "4 - Insufficient training for job demands or new tasks" & vbCrLf
        s = s & "5 - Lack of constructive feedback or performance support" & vbCrLf
        s = s & "6 - Inadequate access to employee assistance or support" & vbCrLf
        s = s & "7 - Lack of info about reporting processes or support" & vbCrLf
        s = s & "8 - Insufficient or inappropriate supervision" & vbCrLf
        s = s & "9 - Inadequate de-escalation or conflict resolution resources"
    Case 3
        s = s & "1 - Adverse environmental conditions (noise, temp, air)" & vbCrLf
        s = s & "2 - Poorly maintained or unclean facilities" & vbCrLf
        s = s & "3 - Workplace layout limiting supervision or support" & vbCrLf
        s = s & "4 - Remote or isolated work location" & vbCrLf
        s = s & "5 - Natural events restricting travel or creating uncertainty" & vbCrLf
        s = s & "6 - Limited access to communication technology" & vbCrLf
        s = s & "7 - Blurring of boundaries between work and home life" & vbCrLf
        s = s & "8 - Workplace culture tolerating unreasonable behaviour" & vbCrLf
        s = s & "9 - Lack of diversity, respect, or inclusion" & vbCrLf
        s = s & "10 - Limited opportunities for social interaction"
    Case 4
        s = s & "1 - No/inadequate policies for psychosocial hazards" & vbCrLf
        s = s & "2 - Policies or procedures not adhered to" & vbCrLf
        s = s & "3 - Procedures that lack flexibility or cannot be applied" & vbCrLf
        s = s & "4 - No/inadequate reporting and complaints processes" & vbCrLf
        s = s & "5 - Investigation/appeals not affording natural justice" & vbCrLf
        s = s & "6 - No/inadequate incident response or investigation" & vbCrLf
        s = s & "7 - Procedures that discriminate or lack fairness" & vbCrLf
        s = s & "8 - Inadequate rostering, hours, or fatigue management" & vbCrLf
        s = s & "9 - Inadequate change management procedures" & vbCrLf
        s = s & "10 - No mechanism for addressing senior mgmt behaviour"
    Case 5
        s = s & "1 - Poor leadership practices or management style" & vbCrLf
        s = s & "2 - Limited management accountability for psych hazards" & vbCrLf
        s = s & "3 - Unclear or conflicting roles, responsibilities" & vbCrLf
        s = s & "4 - Excessive work demands (cognitive, emotional, physical)" & vbCrLf
        s = s & "5 - Workplace culture" & vbCrLf
        s = s & "6 - Multiple interacting hazards not considered together" & vbCrLf
        s = s & "7 - Low job control or autonomy" & vbCrLf
        s = s & "8 - Lack of recognition or reward" & vbCrLf
        s = s & "9 - Organisational injustice (unfairness, bias)" & vbCrLf
        s = s & "10 - Inadequate consultation with workers" & vbCrLf
        s = s & "11 - Poorly managed organisational change" & vbCrLf
        s = s & "12 - Insecure work arrangements" & vbCrLf
        s = s & "13 - Inadequate monitoring and review of controls"
    End Select

    GetFactorsForCategory = s
End Function

Private Function GetFactorText(catNum As Long, factorNum As Long) As String
    Dim f As String
    f = ""
    Select Case catNum
    Case 1
        Select Case factorNum
        Case 1: f = "Inadequate communication between staff"
        Case 2: f = "Unreasonable behaviour by worker(s)"
        Case 3: f = "Unreasonable behaviour by manager(s) or supervisor(s)"
        Case 4: f = "Interpersonal conflict (unresolved or escalating)"
        Case 5: f = "Fatigue or burnout"
        Case 6: f = "Mental stress or psychological distress"
        Case 7: f = "Exposure to traumatic event or material"
        Case 8: f = "Personal issues affecting work capacity"
        Case 9: f = "Incivility or inappropriate behaviour tolerated by peers"
        End Select
    Case 2
        Select Case factorNum
        Case 1: f = "Insufficient staffing levels or resources to meet work demands"
        Case 2: f = "Inadequate or unavailable tools and resources to do the job"
        Case 3: f = "Insufficient training on psychosocial hazard identification or management"
        Case 4: f = "Insufficient training for job demands or new tasks"
        Case 5: f = "Lack of constructive feedback or performance support"
        Case 6: f = "Inadequate access to employee assistance or support services"
        Case 7: f = "Lack of information about reporting processes or support pathways"
        Case 8: f = "Insufficient or inappropriate supervision"
        Case 9: f = "Inadequate de-escalation tools or conflict resolution resources"
        End Select
    Case 3
        Select Case factorNum
        Case 1: f = "Adverse environmental conditions (noise, temperature, air quality)"
        Case 2: f = "Poorly maintained or unclean facilities"
        Case 3: f = "Workplace layout limiting supervision or support access"
        Case 4: f = "Remote or isolated work location"
        Case 5: f = "Natural events restricting travel or creating uncertainty"
        Case 6: f = "Limited access to communication technology"
        Case 7: f = "Blurring of boundaries between work and home life"
        Case 8: f = "Workplace culture tolerating unreasonable behaviour"
        Case 9: f = "Lack of diversity, respect, or inclusion in the workplace"
        Case 10: f = "Limited opportunities for social interaction during work"
        End Select
    Case 4
        Select Case factorNum
        Case 1: f = "No or inadequate policies for managing psychosocial hazards"
        Case 2: f = "Policies or procedures not adhered to"
        Case 3: f = "Procedures that cannot be applied as written or lack flexibility"
        Case 4: f = "No or inadequate reporting and complaints processes"
        Case 5: f = "Investigation or appeals process not affording natural justice"
        Case 6: f = "No or inadequate incident response or investigation procedures"
        Case 7: f = "Procedures that discriminate or lack fairness"
        Case 8: f = "Inadequate rostering, hours of work, or fatigue management procedures"
        Case 9: f = "Inadequate change management procedures"
        Case 10: f = "No mechanism for impartially addressing behaviour by senior management"
        End Select
    Case 5
        Select Case factorNum
        Case 1: f = "Poor leadership practices or management style"
        Case 2: f = "Limited management accountability for psychosocial hazards"
        Case 3: f = "Unclear or conflicting roles, responsibilities, or expectations"
        Case 4: f = "Excessive work demands (cognitive, emotional, or physical)"
        Case 5: f = "Workplace culture"
        Case 6: f = "Multiple interacting psychosocial hazards not considered together"
        Case 7: f = "Low job control or autonomy"
        Case 8: f = "Lack of recognition or reward"
        Case 9: f = "Organisational injustice (unfairness, bias, inconsistency)"
        Case 10: f = "Inadequate consultation with workers on matters affecting them"
        Case 11: f = "Poorly managed organisational change"
        Case 12: f = "Insecure work arrangements"
        Case 13: f = "Inadequate monitoring and review of existing controls"
        End Select
    End Select
    GetFactorText = f
End Function

Private Function ParseSelection(catNum As Long, selection As String) As String
    Dim parts() As String
    parts = Split(selection, ",")
    Dim result As String
    result = ""
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim num As Long
        num = Val(Trim(parts(i)))
        If num > 0 Then
            Dim factorText As String
            factorText = GetFactorText(catNum, num)
            If factorText <> "" Then
                result = result & "  - " & factorText & vbCrLf
            End If
        End If
    Next i
    ParseSelection = result
End Function

Public Sub InsertFactorsIntoTable(ByVal factorText As String)
    Dim tbl As Table
    Dim targetCell As Cell
    Dim found As Boolean
    found = False
    Dim r As Long
    Dim c As Long

    For Each tbl In ActiveDocument.Tables
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                On Error Resume Next
                Dim cellText As String
                cellText = ""
                cellText = tbl.Cell(r, c).Range.Text
                On Error GoTo 0
                If InStr(1, cellText, "Outline any relevant work or non-work factors", vbTextCompare) > 0 Then
                    On Error Resume Next
                    Set targetCell = tbl.Cell(r + 1, 1)
                    If targetCell Is Nothing Then
                        Set targetCell = tbl.Cell(r + 1, c)
                    End If
                    On Error GoTo 0
                    If Not targetCell Is Nothing Then
                        found = True
                        GoTo FoundIt
                    End If
                End If
            Next c
        Next r
    Next tbl

FoundIt:
    If found Then
        Dim rng As Range
        Set rng = targetCell.Range
        rng.End = rng.End - 1

        Dim existingText As String
        existingText = Trim(rng.Text)

        If existingText <> "" Then
            rng.Collapse Direction:=0
            rng.InsertAfter vbCr & vbCr & factorText
        Else
            rng.Text = factorText
        End If

        MsgBox "Factors inserted successfully into the table.", vbInformation, "PEEPO"
    Else
        MsgBox "Could not find the target cell." & vbCrLf & vbCrLf & _
               "Please ensure this document contains the Nature of Concerns " & _
               "table with the row 'Outline any relevant work or non-work factors'.", _
               vbExclamation, "PEEPO"
    End If
End Sub
