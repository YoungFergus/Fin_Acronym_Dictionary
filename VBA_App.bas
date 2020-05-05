Attribute VB_Name = "App"
Option Explicit

Sub AppRun()

'---Purpose: This sub will create an array from the text box, see if it contains matches in the dictionary
'and then will pull the infromation from the dictionary onto the Users' clipboard

Application.ScreenUpdating = False

'   ---VARIABLES---   '

Dim text, Matcher, TestCheck, output, CatchString As String
Dim SplitArray, chkChar, clean As Variant ' clean for string cleaning
Dim i, x, c, lrow, dictrown As Long
Dim dictrow, PrevResults, rng As Range
Dim Answer As VbMsgBoxResult


'   ---TEXT CLEANING---   '

text = Dashboard.Range("C12").Value

If text = "" Then
    MsgBox "Please add text to the box."
    Exit Sub
End If

    For i = 1 To Len(text)
        clean = Mid(text, i, 1)
        If (clean >= "a" And clean <= "z") Or (clean >= "0" And clean <= "9") Or (clean >= "A" And clean <= "Z") Or (clean = "-") Or (clean = ",") Or (clean = ".") Or (clean = "(") Or (clean = ")") Or (clean = "&") Or (clean = "/") Then 'This is where the bug is
            output = output & clean 'add the character to out'
        Else
            output = output & " "
        End If
    Next

output = Application.WorksheetFunction.Trim(output)
text = output

'   ---ASSIGNMENTS---   '

SplitArray = Split(text)
i = UBound(SplitArray) + 1
x = LBound(SplitArray)
c = 1 ' for adding matching definitions to Hidden
lrow = Dict.Cells(Rows.Count, 2).End(xlUp).Row


'   ---MAIN PROGRAM---    '

Hidden.Cells.ClearContents

Do Until x >= i

'   A: Assign ARRAY & CLEAN
    'NOTE: CatchString is tactical fix

    CatchString = "!!!!@@@@####$$$$%%%%^^^^&&&&****(((())))----____++++===={{{{}}}}[[[[]]]]||||\\\\::::;;;;<<<<>>>>,,,,....????////"
    Matcher = SplitArray(x)
    If InStr(1, CatchString, Matcher) <> 0 Then
        Matcher = "x"
    End If

        chkChar = Right(Matcher, 1)
        Do Until chkChar = ""
            chkChar = Right(Matcher, 1)
            If (chkChar >= "a" And chkChar <= "z") Or (chkChar >= "0" And chkChar <= "9") Or (chkChar >= "A" And chkChar <= "Z") Then
                Exit Do
            Else
                Matcher = Left(Matcher, Len(Matcher) - 1)
            End If
        Loop

        Do Until chkChar = ""
            chkChar = Left(Matcher, 1)
            If (chkChar >= "a" And chkChar <= "z") Or (chkChar >= "0" And chkChar <= "9") Or (chkChar >= "A" And chkChar <= "Z") Then
                Exit Do
            Else
                Matcher = Right(Matcher, Len(Matcher) - 1)
            End If
        Loop

'   B: ACRONYM PRECHECK

    TestCheck = UCase(Matcher)
        If Matcher = TestCheck Then

            Set PrevResults = Hidden.Range("A:A").Find(What:=Matcher, LookIn:=xlValues, LookAt:=xlWhole)
            If PrevResults Is Nothing Or IsEmpty(Hidden.Range("A1")) Then 'ADD ON

                'C: PROGRAM FUNCTION

                Set dictrow = Dict.Range("B2", "B" & lrow).Find(What:=Matcher, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                If Not dictrow Is Nothing Then
                    If Dashboard.Range("L15") = False Then 'CHECKING OPTIONS BOX
                        dictrown = dictrow.Row
                        If dictrown < lrow Then
                            Dict.Range("B" & dictrown, "C" & dictrown).Copy Hidden.Range("A" & c)
                        End If
                    Else
                        dictrown = dictrow.Row
                        If dictrown < lrow Then
                            Dict.Range("B" & dictrown, "D" & dictrown).Copy Hidden.Range("A" & c)
                        End If
                    End If
                End If

                c = c + 1
            End If
        End If

        '   D: LOOP

        x = x + 1

        Loop

'   ---CLEAN UP & PRESENTATION---   '

    i = 1
    'On Error GoTo Handle:

    lrow = Hidden.Cells(1048576, 1).End(xlUp).Row

    On Error GoTo 1004:
    Set rng = Hidden.Range("A1", "A" & lrow).SpecialCells(xlCellTypeBlanks)
    rng.EntireRow.Delete

1004: ' No blanks

    If IsEmpty(Hidden.Range("A1")) Then
        VBA.MsgBox "No acronyms found."

    ElseIf Dashboard.Range("L18") = True Then
        NiceFormat
        Formatting.Range("A1").CurrentRegion.Copy
        Formatting.Activate
        MsgBox "Results copied."
    Else
        Answer = MsgBox("Would you like to copy the result to your clipboard?", vbYesNo)
            If Answer = vbYes Then
                Hidden.Range("A1").CurrentRegion.Copy
            End If
        Hidden.Activate
    End If

Application.ScreenUpdating = True
Exit Sub

Handle:
    MsgBox "No results found!"
End Sub


Public Sub NiceFormat()

Dim rowc As Integer

Formatting.Cells.ClearContents
Formatting.Cells.ClearFormats

Hidden.Range("A1").CurrentRegion.Copy Formatting.Range("A3")

Formatting.Range("A1").Value = "Glossary"
With Formatting.Range("A1").Font
    .Name = "Arial Black"
    .Size = 16
End With

'   ---GLOSSARY TITLE---   '
If Dashboard.Range("L15") = True Then
    Formatting.Range("A1:D1").Merge
    Formatting.Range("A1:D1").HorizontalAlignment = xlCenter
    Formatting.Range("A1:D1").Interior.ColorIndex = 15
    Formatting.Range("A1:D1").BorderAround ColorIndex:=1, Weight:=xlThick
End If

'   ---COLUMN TITLES---   '
Formatting.Range("A2").Value = "Acronym"
    With Formatting.Range("A2").Font
        .Name = "Calibri Light"
        .Size = 11
        .Bold = True
    End With

Formatting.Range("B2").Value = "Full Name"
    With Formatting.Range("B2").Font
        .Name = "Calibri Light"
        .Size = 11
        .Bold = True
    End With

If Dashboard.Range("L15") = True Then
    Formatting.Range("C2").Value = "Defintion"
        With Formatting.Range("C2").Font
            .Name = "Calibri Light"
            .Size = 11
        End With

        Formatting.Range("A2:D2").BorderAround ColorIndex:=1, Weight:=xlThick
        Formatting.Columns("A:D").AutoFit
Else
    Formatting.Range("A2:B2").BorderAround ColorIndex:=1, Weight:=xlThick
    Formatting.Columns("A:B").AutoFit
End If

rowc = 3
Do Until rowc = Formatting.Cells(Rows.Count, 1).End(xlUp).Row + 1

    If rowc Mod 2 = 1 Then
        Formatting.Cells.Rows(rowc).Interior.ColorIndex = 15
    Else
        Formatting.Cells.Rows(rowc).Interior.ColorIndex = 2
    End If

rowc = rowc + 1
Loop


End Sub
