Attribute VB_Name = "Main"
Public numCols As Integer
Public table As ListObject
Public ws As Worksheet
Public emailHistWs As Worksheet
Public emailHistTable As ListObject

Public Function getData() As Object()

'Enea Zguro
'OGS Procurement Services
'9/10/24
'This function processes data in a table, formatting it in a Vendor Object (see Vendor Class for reference) array.

Dim numRows As Integer, rowsIncrement As Integer, vendorsIndexer As Integer, quarterIndexer As Integer, vendorRowIndexer As Integer

numRows = table.Range.SpecialCells(xlCellTypeVisible).rows.Count
numCols = table.Range.SpecialCells(xlCellTypeVisible).Columns.Count

'start at noninformational rows + 1
rowsIncrement = 2
vendorsIndexer = 0
'vendorRowIndexer = 1

'Error trap to account for empty table
On Error Resume Next
If table.DataBodyRange Is Nothing Or WorksheetFunction.CountA(table.DataBodyRange) = 0 Then
    GoTo endFunc
End If

Dim vendors() As Object
'subtract non informational rows + 1
ReDim vendors(0)

'Gather vendor information per row
Do While rowsIncrement < numRows + 1
    
    If UBound(vendors) < vendorsIndexer Then
        ReDim Preserve vendors(UBound(vendors) - LBound(vendors) + 1)
    End If
    
    quarterIndexer = 0
    Dim tempVendor As Vendor
    Set tempVendor = New Vendor
    
    If StrComp(table.Range(rowsIncrement, 4), "Yes") = 0 Then
        'GoTo nextVendor
        tempVendor.status = True
    End If
    'tempVendor.row = vendorRowIndexer
    tempVendor.row = rowsIncrement - 1
    
    'assuming the structure of name|contract|email
    If Not StrComp(table.Range(rowsIncrement, 1), "") = 0 Then
        tempVendor.name = table.Range(rowsIncrement, 1)
    End If
    
    If Not StrComp(table.Range(rowsIncrement, 2), "") = 0 Then
        tempVendor.contract = table.Range(rowsIncrement, 2)
    End If
    
    If Not StrComp(table.Range(rowsIncrement, 3), "") = 0 Then
        tempVendor.email = table.Range(rowsIncrement, 3)
    End If
    
    For i = 7 To numCols Step 1
            tempVendor.setQuarter(quarterIndexer) = "N/A"
        If StrComp(CStr(table.Range(rowsIncrement, i)), "Not Requested") = 0 Or StrComp(CStr(table.Range(rowsIncrement, i)), "Not Sent") = 0 Or StrComp(CStr(table.Range(rowsIncrement, i)), "Not Submitted") = 0 Then
            tempVendor.setQuarter(quarterIndexer) = "Not Requested"
        ElseIf StrComp(CStr(table.Range(rowsIncrement, i)), "Submitted") = 0 Or StrComp(CStr(table.Range(rowsIncrement, i)), "Approved") = 0 Then
            tempVendor.setQuarter(quarterIndexer) = "Submitted"
        ElseIf StrComp(CStr(table.Range(rowsIncrement, i)), "Submitted Incorrectly") = 0 Then
            tempVendor.setQuarter(quarterIndexer) = "Submitted Incorrectly"
        ElseIf IsDate(table.Range(rowsIncrement, i)) = True Then
            tempVendor.setQuarter(quarterIndexer) = CStr(table.Range(rowsIncrement, i))
        End If
        quarterIndexer = quarterIndexer + 1
    Next i
    
    Set vendors(vendorsIndexer) = tempVendor
    vendorsIndexer = vendorsIndexer + 1
'nextVendor:
rowsIncrement = rowsIncrement + 1
'vendorRowIndexer = vendorRowIndexer + 1
Loop

getData = vendors

If False Then
endFunc:
   MsgBox "No data to process. Please enter the data in the table!", , "Insert Data"
   End If
End Function

Function wsExists(name As String) As Boolean
    Dim tSheet As Worksheet
    On Error Resume Next
       Set tSheet = ThisWorkbook.Sheets(name)
    On Error GoTo 0
        wsExists = Not tSheet Is Nothing
End Function

Function tbExists(name As String) As Boolean
    On Error GoTo res
        tRes = StrComp(ws.ListObjects(name).name, name)
        If tRes = 0 Then
            tbExists = True
            GoTo ext
        End If
res:
    tbExists = False
ext:
End Function

Function getTableName(sht As String) As String
    getTableName = "The Name of the Current Table is: " & Worksheets(sht).ListObjects(1).name
End Function
Function setInformation()
    Set ws = ThisWorkbook.ActiveSheet
    Set table = ws.ListObjects(1)
    Set emailHistWs = Worksheets("Email History")
    Set emailHistTable = emailHistWs.ListObjects("EmailHistTable")
End Function
Sub addRow()
    setInformation
    If table.ListRows.Count = 0 Then
        table.ListRows.Add
    End If
    table.ListRows.Add
    emailHistTable.ListColumns.Add
    'refreshQuery
   ' With table.ListRows(table.ListRows.Count).Range
       ' .Borders(xlDiagonalDown).LineStyle = xlNone
       ' .Borders(xlDiagonalUp).LineStyle = xlNone
       ' .Borders(xlEdgeLeft).LineStyle = xlNone
       ' .Borders(xlEdgeTop).LineStyle = xlNone
       ' .Borders(xlEdgeBottom).LineStyle = xlNone
       ' .Borders(xlEdgeRight).LineStyle = xlNone
       ' .Borders(xlInsideVertical).LineStyle = xlNone
      '  .Borders(xlInsideHorizontal).LineStyle = xlNone
    'End With
End Sub
Sub addColumn()
    setInformation
    table.ListColumns.Add
    'refreshQuery
    'With table.ListColumns(table.ListColumns.Count).Range
       ' .Borders(xlDiagonalDown).LineStyle = xlNone
       ' .Borders(xlDiagonalUp).LineStyle = xlNone
        '.Borders(xlEdgeLeft).LineStyle = xlNone
       ' .Borders(xlEdgeTop).LineStyle = xlNone
       ' .Borders(xlEdgeBottom).LineStyle = xlNone
      '  .Borders(xlEdgeRight).LineStyle = xlNone
       ' .Borders(xlInsideVertical).LineStyle = xlNone
       ' .Borders(xlInsideHorizontal).LineStyle = xlNone
    'End With
End Sub
Sub deleteRow()
    setInformation
    On Error GoTo endSub
    table.ListRows(table.ListRows.Count).Delete
    'refreshQuery
endSub:
End Sub
Sub deleteColumn()
    setInformation
    If Not table.ListColumns.Count < 9 Then
        table.ListColumns(table.ListColumns.Count).Delete
        'refreshQuery
    End If
End Sub
Sub includeAll()
    setInformation
    Dim rows As Integer
    rows = 2
    Do While rows <= table.ListRows.Count + 1
        table.Range(rows, 4) = "Yes"
        rows = rows + 1
    Loop
End Sub
Sub excludeAll()
    setInformation
    Dim rows As Integer
    rows = 2
    Do While rows <= table.ListRows.Count + 1
        table.Range(rows, 4) = "No"
        rows = rows + 1
    Loop
End Sub
Sub refreshQuery()
    setInformation
    ThisWorkbook.RefreshAll
    Dim min As Integer, max As Integer
    min = ThisWorkbook.Sheets("Email History").ListObjects("queryTable").ListColumns.Count
    max = emailHistTable.ListColumns.Count
    If min = 1 Then
    GoTo endSub
    End If
    If Not min = max Then
        Do While max > min
            emailHistWs.Columns(max).Delete
            max = max - 1
        Loop
    End If
endSub:
End Sub

Sub updateHistory()
    setInformation
    emailHistTable.ListRows.Add
End Sub
Sub createWordDocument()

'Enea Zguro
'OGS Procurement Services
'9/10/24
'This sub is meant to generate a report of missing vendors' quarter sales reports in the form of a template email via Word.
refreshQuery
Dim WordApp As Word.Application, vendors() As Object, emails() As String, v As Variant, quarters() As String, headings() As String

setInformation
vendors = getData()

If IsEmpty(vendors) Then
    GoTo endSub
End If

'Error trap to stop generating report if not needed
'On Error Resume Next
'If UBound(vendors) < LBound(vendors) Then
    'GoTo endSub
'End If

'Subtract the number of columns (inclusive) of personal information
ReDim headings(numCols - 7)

Set WordApp = New Word.Application

'Start i at the column value of nonpersonal information
For i = 7 To numCols
    headings(i - 7) = table.HeaderRowRange(1, i)
Next i

Dim tempWs As Worksheet, rnge As Range
Set tempWs = ThisWorkbook.Worksheets("Validation Sheet")

'Email templates
Dim emailDraft As String
emailDraft = ThisWorkbook.Sheets("Customized Language").Range("B2").Value

WordApp.Visible = True
With WordApp
.Activate
.Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0, Visible:=True
Dim res As String, vendorInfo As String, notRQ As String, correct As String, incorrect As String, dueBy As String, dt As String, vowels() As String, qt() As String
ReDim vowels(10)
vowels = Split("a,e,i,o,u,A,E,I,O,U", ",")
Dim tempVAR As String
Dim qIndex As Integer
Dim vendorIndexer As Integer
vendorIndexer = 2
'emailHistTable.ListRows.Add
If Not emailHistTable.ListColumns.Count = 1 Then
    updateHistory
End If

For Each v In vendors
    If v Is Nothing Then
        GoTo nextVendor
    End If
    If Not v.status Then
        GoTo nextVendor
    End If
    For qIndex = 0 To UBound(v.quarter)
        tempVAR = v.Iquarter(qIndex)
        If Not StrComp(v.Iquarter(qIndex), "N/A") = 0 Then
            If Not StrComp(v.Iquarter(qIndex), "") = 0 Then
                GoTo verif
            End If
        End If
    Next qIndex
    GoTo nextVendor
verif:
    res = ""
    vendorInfo = ""
    notRQ = ""
    correct = ""
    incorrect = ""
    quarters = v.quarter
    'Gather email requirements for current vendor
    'Subtract number of personal columns (inclusive)
    For i = 0 To numCols - 7
        If StrComp(quarters(i), "Not Requested") = 0 Then
            notRQ = notRQ & CStr(headings(i)) & " BULLET" & vbNewLine
        End If
        If StrComp(quarters(i), "Submitted") = 0 Then
            correct = correct & CStr(headings(i)) & " BULLET" & vbNewLine
        End If
        If StrComp(quarters(i), "Submitted Incorrectly") = 0 Then
            incorrect = incorrect & CStr(headings(i)) & " BULLET" & vbNewLine
        End If
        If StrComp(headings(i), "Due By") = 0 Then
            dueBy = CStr(quarters(i))
        End If
    Next i
     
    res = emailDraft
    
    If tempWs.Range("C1").Value >= 12 Then
        res = Replace(res, "(morning)", "afternoon")
    Else
        res = Replace(res, "(morning)", "morning")
    End If
    
    'Replace specific vendor information and generate specific email
    res = Replace(res, "(vendor)", CStr(v.name))
    
    For Each vo In vowels
        flag = False
        If StrComp(Left(CStr(v.contract), 1), vo) = 0 Then
            flag = True
            GoTo found
        End If
    Next vo
found:
    If flag Then
        res = Replace(res, "(a)", "an")
    Else
        res = Replace(res, "(a)", "a")
    End If
    res = Replace(res, "(Insert Contract Name)", CStr(v.contract))
    res = Replace(res, "(received)(reason)", correct)
    res = Replace(res, "(incorrectly)(reason)", incorrect)
    res = Replace(res, "(notreceived)(reason)", notRQ)
    res = Replace(res, "(Insert Date)", dueBy)
    res = Replace(res, "(email)", CStr(v.email))
    
    'emailHistTable.Range(v.row, 5).Value = table.Range(v.row, 6).Value
    
    
    .Selection.TypeText res
nextVendor:
    If Not v.status Then
        emailHistTable.Range(emailHistTable.ListRows.Count, v.row).Value = "N/A"
    Else
        emailHistTable.Range(emailHistTable.ListRows.Count, v.row).Value = "Requested on: " + Format(Date, "mm-dd-yyyy") + " at " + Format(Time, "hh:mm:ss")
    End If
    
    table.Range(v.row + 1, 6).Value = table.Range(v.row + 1, 6).Value + 1
    vendorIndexer = vendorIndexer + 1
Next v
End With

'Apply numbered list to reasons and bullet points to reasons in email
'Highlight missing vendor information (if applicable)
With WordApp
    For Each Rng In .ActiveDocument.Words
        .Selection.Collapse
        Rng.Select
            If StrComp(Rng.Text, "BULLET") = 0 Then
                .Selection.TypeText " "
                .Selection.Range.ListFormat.ApplyBulletDefault
            End If
            If StrComp(Rng.Text, "REASONREPLACE") = 0 Then
                .Selection.TypeText " "
                .Selection.Range.ListFormat.ListIndent
            End If
            If StrComp(Rng.Text, "MissingMISSING ") = 0 Then
                .Selection.Range.HighlightColorIndex = wdYellow
                .Selection.TypeText "Missing "
            End If
            If StrComp(Rng.Text, "EmailMISSING") = 0 Then
                .Selection.Range.HighlightColorIndex = wdYellow
                .Selection.TypeText "Email"
            End If
            If StrComp(Rng.Text, "NameMISSING") = 0 Then
                .Selection.Range.HighlightColorIndex = wdYellow
                .Selection.TypeText "Name"
            End If
            If StrComp(Rng.Text, "NameOfContractMISSING ") = 0 Then
                .Selection.Range.HighlightColorIndex = wdYellow
                .Selection.TypeText "'Name of The Contract' "
            End If
            If StrComp(Rng.Text, "Note") = 0 Then
                .Selection.Font.Bold = True
            End If
            If StrComp(Rng.Text, "ContractContract ") = 0 Then
                .Selection.TypeText ""
            End If
    Next Rng
End With

Set WordApp = Nothing
endSub:
refreshQuery
ThisWorkbook.Save
End Sub
