Attribute VB_Name = "find_duplicates"
Option Explicit

Sub Button1_Click()
Call sbFindDuplicatesInColumn
End Sub
Sub sbFindDuplicatesInColumn()
Dim lastRow As Long
Dim matchFoundIndex As Long, i As Long, countif As Long, cell As Long
Dim iCntr As Long
Dim rngX As Range
Dim rng1 As Range
Dim C As Range
Dim objDic
Dim strMsg As String

'Just change the "B" icon to the column you want the code to detect the last row
With ActiveSheet
    lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
End With


'change the column number (3) to the column in which all the duplicate text
'are entered this is to assure that your previous checking of duplicate will not conflict on the new one

Columns(3).ClearContents

For iCntr = 1 To lastRow
'change the number 2 based on the column number in which checking of duplicates will be executed. in our example it is B which is equal to 2
If Cells(iCntr, 2) <> "" Then
'change 2 and column B1:B to the column numner and column letter you want duplictes to be checked
matchFoundIndex = WorksheetFunction.Match(Cells(iCntr, 2), Range("B1:B" & lastRow), 0)
If iCntr <> matchFoundIndex Then

'change  3 to the column in which duplicate text will be printed
Cells(iCntr, 3) = "Duplicate"
End If
End If
Next

'change b1 and b to the column in which duplicate text will be checked
    Set objDic = CreateObject("scripting.dictionary")
    Set rng1 = Range([b1], Cells(Rows.Count, "B").End(xlUp))
    For Each C In rng1
        If Len(C.Value) > 0 Then
            If Not objDic.exists(C.Value) Then
                objDic.Add C.Value, 1
            Else
                strMsg = strMsg & C.Value & " found in cell " & C.Address(0, 0) & vbNewLine
            End If
        End If
    Next
    If Len(strMsg) > 0 Then MsgBox strMsg

cell = 1

For i = 1 To lastRow
cell = cell + 1
'change c to the column where the Duplicate text color will be change
If ActiveSheet.Range("C" & cell).Value = "Duplicate" Then

ActiveSheet.Range("c" & cell).Font.Color = vbRed
End If

Next i
'
End Sub



Sub GetDupes()
    
End Sub
