Attribute VB_Name = "Main"
Option Explicit
Public numCols As Integer
Public table As ListObject
Public ws As Worksheet
Public emailHistWs As Worksheet
Public emailHistTable As ListObject
Public WordApp As Word.Application
Public documentList() As String
Public documentNumber As Integer
Public buttonValue As Integer
Public totalWrongReasons As Integer
Sub setbuttonValue()
    ActiveSheet.CommandButton1.BackColor = RGB(220, 105, 0)
    buttonValue = Not buttonValue
End Sub
Sub HowToUse()
' Directs the user to the sheet containing instructions on how to use the tracker.
    Sheets("How To Use").Select
End Sub
Public Sub formatSITable()

'Enea Zguro
'Procurement Services
'1/6/25
'This sub reads the table from email generator and creates a new worksheet for the user to input reasons for documents marked as submitted incorrectly
'These reasons will be read and displayed in the generated emails

    'define cell and row loop variables, refresh tables
    Dim c As Range, row As Range, thisCurrentTable As ListObject
    ThisWorkbook.RefreshAll
    
    'unhide template worksheet, turn off alerts and updating
    Application.ScreenUpdating = False
    Sheets("Sub Inc C").Visible = True
    Application.DisplayAlerts = False
    
    'clear previous reasons worksheet
    If Evaluate("ISREF('Previous Sub Incorrectly Reason'!A1)") Then
        Worksheets("Previous Sub Incorrectly Reason").Delete
    End If
    
    'copy current reasons into previous reasons worksheet
    Sheets("Submitted Incorrectly Reasons").Copy After:=Sheets("Email History")
    ActiveSheet.name = "Previous Sub Incorrectly Reason"
    Sheets("Previous Sub Incorrectly Reason").ListObjects(1).Unlink
    
    'clear current reasons worksheet
    If Evaluate("ISREF('Submitted Incorrectly Reasons'!A1)") Then
        Worksheets("Submitted Incorrectly Reasons").Delete
    End If
    Application.DisplayAlerts = True
    
    'copy template worksheet and format
    Sheets("Sub Inc C").Copy After:=Sheets("Email History")
    ActiveSheet.name = "Submitted Incorrectly Reasons"
    Sheets("Submitted Incorrectly Reasons").ListObjects(1).Unlink
    Set thisCurrentTable = Worksheets("Submitted Incorrectly Reasons").ListObjects(1)
    thisCurrentTable.DataBodyRange.ClearFormats
    
    'color 45 is orange format for submitted incorrectly
    'color 48 is grey format for unneeded cells
    'border 1, xlThin is standard border format for needed cells
    
    For Each row In thisCurrentTable.DataBodyRange.rows
        For Each c In row.Cells
            If StrComp(c.Value, "") = 0 Then
                c.Interior.ColorIndex = 48
            End If
            If StrComp(c.Value, "Submitted Incorrectly") = 0 Then
                c.Value = ""
                c.Interior.ColorIndex = 45
                c.BorderAround LineStyle:=1, Weight:=xlThin
            ElseIf StrComp(c.Value, "Submitted") = 0 Then
                c.Value = ""
                c.Interior.ColorIndex = 48
            ElseIf StrComp(c.Value, "Approved") = 0 Then
                c.Value = ""
                c.Interior.ColorIndex = 48
            ElseIf StrComp(c.Value, "Not Requested") = 0 Then
                c.Value = ""
                c.Interior.ColorIndex = 48
            ElseIf StrComp(c.Value, "Not Sent") = 0 Then
                c.Value = ""
                c.Interior.ColorIndex = 48
            ElseIf StrComp(c.Value, "Not Submitted") = 0 Then
                c.Value = ""
                c.Interior.ColorIndex = 48
            ElseIf StrComp(c.Value, "No Longer In Consideration") = 0 Then
                c.Value = ""
                c.Interior.ColorIndex = 48
            End If
        Next c
    Next row
    
    'hide template worksheet and turn on updates
    Sheets("Sub Inc C").Visible = False
    Application.ScreenUpdating = True
End Sub
Public Function emailValidation(email As String, status As Boolean) As Boolean
    'This function takes an email and boolean and determines email validation
    'The boolean status will always return true if it is false to allow empty emails on non-marked users
    
    If Not status Then
        emailValidation = True
        GoTo endfunc
    End If
    Dim arr, iPos As Long, iLen As Long

    arr = Split(email, "@")
    
    If UBound(arr) - LBound(arr) <> 1 Then Exit Function
    If Len(arr(0)) < 1 Then Exit Function
    
    iLen = Len(arr(1))
    If iLen < 3 Then Exit Function
    
    iPos = InStr(1, arr(1), ".", vbBinaryCompare)
    If iPos <= 1 Then Exit Function
    If iPos = iLen Then Exit Function

emailValidation = True
endfunc:
End Function
Public Function getData() As Object()

'Enea Zguro
'OGS Procurement Services
'9/10/24
'This function processes data in a table, formatting it in a Vendor Object (see Vendor Class for reference) array.

Dim numRows As Integer, rowsIncrement As Integer, vendorsIndexer As Integer, quarterIndexer As Integer, vendorRowIndexer As Integer, emptyRng As Boolean
setInformation

numRows = table.Range.SpecialCells(xlCellTypeVisible).rows.Count
numCols = table.Range.SpecialCells(xlCellTypeVisible).Columns.Count
emptyRng = True

'start at noninformational rows + 1
rowsIncrement = 2
vendorsIndexer = 0

'Error trap to account for empty table
On Error Resume Next
If table.DataBodyRange Is Nothing Or WorksheetFunction.CountA(table.DataBodyRange) = 0 Then
    GoTo endfunc
End If

Dim vendors() As Object
ReDim vendors(0)

'Gather vendor information per row
Do While rowsIncrement < numRows + 1
    emptyRng = True
    If UBound(vendors) < vendorsIndexer Then
        ReDim Preserve vendors(UBound(vendors) - LBound(vendors) + 1)
    End If
    
    quarterIndexer = 0
    Dim tempVendor As vendor
    Set tempVendor = New vendor
    
    'check if vendor was marked for email
    If StrComp(table.Range(rowsIncrement, 4), "Yes") = 0 Then
        tempVendor.status = True
    End If
    
    tempVendor.row = rowsIncrement - 1
    
    'assuming the structure of name|contract|email
    If Not StrComp(table.Range(rowsIncrement, 1), "") = 0 Then
        tempVendor.name = table.Range(rowsIncrement, 1)
    End If
    
    If Not StrComp(table.Range(rowsIncrement, 2), "") = 0 Then
        tempVendor.contract = table.Range(rowsIncrement, 2)
    End If
    
    If Not StrComp(table.Range(rowsIncrement, 3), "") = 0 Then
        If emailValidation(table.Range(rowsIncrement, 3), tempVendor.status) Then
            tempVendor.email = table.Range(rowsIncrement, 3)
        Else
            MsgBox "Invalid email for " & tempVendor.name & ".", vbOKOnly, "Email Validation"
        End If
    End If
    
    'start at first column of document collection
    Dim i As Integer
    For i = 6 To numCols Step 1
            tempVendor.setQuarter(quarterIndexer) = "N/A"
        If StrComp(CStr(table.Range(rowsIncrement, i)), "Not Requested") = 0 Or StrComp(CStr(table.Range(rowsIncrement, i)), "Not Sent") = 0 Or StrComp(CStr(table.Range(rowsIncrement, i)), "Not Submitted") = 0 Then
            tempVendor.setQuarter(quarterIndexer) = "Not Requested"
            emptyRng = False
        ElseIf StrComp(CStr(table.Range(rowsIncrement, i)), "Submitted") = 0 Or StrComp(CStr(table.Range(rowsIncrement, i)), "Approved") = 0 Then
            tempVendor.setQuarter(quarterIndexer) = "Submitted"
            emptyRng = False
        ElseIf StrComp(CStr(table.Range(rowsIncrement, i)), "Submitted Incorrectly") = 0 Then
            tempVendor.setQuarter(quarterIndexer) = "Submitted Incorrectly"
            emptyRng = False
        ElseIf IsDate(table.Range(rowsIncrement, i)) = True Then
            tempVendor.setQuarter(quarterIndexer) = CStr(table.Range(rowsIncrement, i))
            emptyRng = False
        End If
        quarterIndexer = quarterIndexer + 1
    Next i
    
    'skip users that have no docuement data entered
    If emptyRng Then
        tempVendor.status = False
    End If
    
    Set vendors(vendorsIndexer) = tempVendor
    vendorsIndexer = vendorsIndexer + 1
    rowsIncrement = rowsIncrement + 1
Loop

getData = vendors

If False Then
endfunc:
   MsgBox "No data to process. Please enter the data in the table!", , "Insert Data"
   End If
End Function
Function wsExists(name As String) As Boolean
    'helper function to check if a worksheet exists
    Dim tSheet As Worksheet
    On Error Resume Next
       Set tSheet = ThisWorkbook.Sheets(name)
    On Error GoTo 0
        wsExists = Not tSheet Is Nothing
End Function
Function tbExists(name As String) As Boolean
    'helper function to check if a table exists
    Dim tRes As String
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
    'helper function to get the name of the first table in a worksheet
    getTableName = "The Name of the Current Table is: " & Worksheets(sht).ListObjects(1).name
End Function
Sub clearWordDocuments()
    ''For i = 0 To UBound(documentList)
        'wd.Documents.Open (documentList(i))
    'Next i
End Sub
Function setInformation()
    'helper  function to set the respective information
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
End Sub
Sub addColumn()
    setInformation
    table.ListColumns.Add
End Sub
Sub deleteRow()
    setInformation
    On Error GoTo endSub
    table.ListRows(table.ListRows.Count).Delete
endSub:
End Sub
Sub deleteColumn()
    setInformation
    If Not table.ListColumns.Count < 8 Then
        table.ListColumns(table.ListColumns.Count).Delete
    End If
End Sub
Sub unCount()
    'removes the previous record of history from the history table and request count
    setInformation
    Dim vendors() As Object, i As Integer
    vendors = getData()
    For i = 2 To table.ListRows.Count + 1
        If vendors(i - 2).status Then
            table.Range(i, 5) = table.Range(i, 5) - 1
            emailHistTable.Range(emailHistTable.ListRows.Count, i - 1) = "N/A"
        End If
    Next i
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
    'refreshQuery refreshes the query connection used for the headers on the email history log
    'refreshQuery also updates the custom table to match the query and custom totals
    
    setInformation
    ThisWorkbook.RefreshAll
    Dim min As Integer, max As Integer, row As Integer, col As Integer
    min = ThisWorkbook.Sheets("Email History").ListObjects("queryTable").ListColumns.Count
    max = emailHistTable.ListColumns.Count
    
    'if only one user is available, go to custom totals section
    If min = 1 Then
        GoTo endSub
    End If
    
    If Not min = max Then
        Do While max > min
            emailHistWs.Columns(max).Delete
            max = max - 1
        Loop
    End If
    
    'custom total formula string concat workaround
    Dim S As String
    S = "`Requested on*`"
    S = Replace(S, "`", Chr(34)) 'replace special character with quotations
    row = emailHistTable.ListRows.Count + 1
    col = emailHistTable.ListColumns.Count
    
endSub:
    
    Dim rnge As Range, a As String, b As String, bw() As String
    a = "A" & CStr(row + 3)
    bw = Split((Columns(col).Address(, 0)), ":")
    b = bw(1) & CStr(row + 3)
    
    If IsEmpty(ThisWorkbook.Sheets("Email History").Range(a).Value) Then
        ThisWorkbook.Sheets("Email History").Range(a).Value = "=COUNTIF([Column1]," & S & ")"
    End If
    Set rnge = Range(a, b)
    ThisWorkbook.Sheets("Email History").Activate
    On Error GoTo endfunc
    emailHistTable.Range(row, 1).AutoFill Destination:=rnge
    
endfunc:
ThisWorkbook.Sheets("Email Generator").Activate
Worksheets("Email History").Cells.EntireColumn.autofit
End Sub
Sub updateHistory()
    'ThisWorkbook.Sheets("Email History").Protect "abc123"
    setInformation
    emailHistTable.ListRows.Add
End Sub
Sub createWordDocument()

'Enea Zguro
'OGS Procurement Services
'9/10/24
'This sub is meant to generate a report of missing vendors' quarter sales reports in the form of a template email via Word.

ufProg.StartUpPosition = 0
ufProg.Left = Application.Left + (0.5 * Application.Width) - (0.5 * ufProg.Width)
ufProg.Top = Application.Top + (0.5 * Application.Height) - (0.5 * ufProg.Height)
ufProg.Show
ufProg.LabelProg.Width = 0
ufProg.LabelCaption.Caption = "Reading Data, Please Wait..."

refreshQuery
Dim vendors() As Object, emails() As String, v As Variant, quarters() As String, headings() As String, out As Object, vowels() As String, vendorIndexer As Integer, mail As Object

setInformation
vendors = getData()

If IsEmpty(vendors) Then
    GoTo endSub
End If

'Subtract the number of columns (inclusive) of personal information
ReDim headings(numCols - 6)

'Start i at the column value of nonpersonal information
Dim i As Integer
For i = 6 To numCols
    headings(i - 6) = table.HeaderRowRange(1, i)
Next i

'Email templates
Dim emailDraft As String
If buttonValue = 2 Then
    emailDraft = ThisWorkbook.Sheets("Customized Language").Range("B2").Value
Else
    emailDraft = ThisWorkbook.Sheets("Customized Language").Range("B3").Value
End If

If buttonValue = 2 Then
    Set WordApp = New Word.Application
    WordApp.Visible = True
    
    With WordApp
        .Activate
        .Documents.Add template:="Normal", NewTemplate:=False, DocumentType:=0, Visible:=True
        WordApp.ActiveWindow.WindowState = wdWindowStateMinimize
    End With
ElseIf buttonValue = 0 Or buttonValue = 1 Then
    'Get the user's custom email signature
    Set out = CreateObject("Outlook.Application")
    Set mail = out.CreateItem(0)
    Dim emailSignature As String
    emailSignature = Environ("appdata") & "\Microsoft\Signatures\"
    If Dir(emailSignature, vbDirectory) <> vbNullString Then
        emailSignature = emailSignature & Dir$(emailSignature & "*.htm")
    Else
        emailSignature = ""
    End If
    emailSignature = CreateObject("Scripting.FileSystemObject").GetFile(emailSignature).OpenAsTextStream(1, -2).ReadAll
End If

Dim res As String

ReDim vowels(10)
vowels = Split("a,e,i,o,u,A,E,I,O,U", ",")
vendorIndexer = 2

If Not emailHistTable.ListColumns.Count = 1 Then
    updateHistory
End If

'skip empty vendors
Dim vStep As Integer
totalWrongReasons = 0
vStep = 0
For Each v In vendors
    If v Is Nothing Then
        GoTo nextvendor
    End If
    If Not v.status Then
        GoTo nextvendor
    End If
    
    res = formatVendorEmail(v, emailDraft, headings, vowels, vendorIndexer)
    
    If buttonValue = 2 Then
        WordApp.Selection.TypeText res
    Else
        Set mail = out.CreateItem(0)
        With mail
            .To = v.email
            .Subject = "Some Appropriate Subject"
            .HTMLBody = res & "<br>" & emailSignature & .HTMLBody
            If buttonValue = 1 Then
                .send
            Else
                .display
            End If
        End With
    End If
    
nextvendor:
    If Not v.status Then
        emailHistTable.Range(emailHistTable.ListRows.Count, v.row).Value = "N/A"
    Else
        table.Range(v.row + 1, 5).Value = table.Range(v.row + 1, 5).Value + 1
        emailHistTable.Range(emailHistTable.ListRows.Count, v.row).Value = "Requested on: " + Format(Date, "mm-dd-yyyy") + " at " + Format(Time, "hh:mm:ss")
    End If
    vendorIndexer = vendorIndexer + 1
Next v


'Apply numbered list to reasons and bullet points to reasons in email
'Highlight missing vendor information (if applicable)
Dim totalWords As Integer, cur As Integer, rng As Object
cur = 0
If buttonValue = 2 Then
    With WordApp
        totalWords = .ActiveDocument.Words.Count + totalWrongReasons
        .Application.ScreenUpdating = False
        .ActiveWindow.WindowState = wdWindowStateMinimize
        
        For Each rng In .ActiveDocument.Words
            cur = cur + 1
            
            .Selection.Collapse
            rng.Select
                If StrComp(rng.Text, "BULLET") = 0 Then
                    .Selection.TypeText " "
                    .Selection.Range.ListFormat.ApplyBulletDefault
                End If
                If StrComp(rng.Text, "REASONREPLACE") = 0 Then
                    .Selection.TypeText " "
                    .Selection.Range.ListFormat.ListIndent
                End If
                If StrComp(rng.Text, "MissingMISSING ") = 0 Then
                    .Selection.Range.HighlightColorIndex = wdYellow
                    .Selection.TypeText "Missing "
                End If
                If StrComp(rng.Text, "EmailMISSING") = 0 Then
                    .Selection.Range.HighlightColorIndex = wdYellow
                    .Selection.TypeText "Email"
                End If
                If StrComp(rng.Text, "NameMISSING") = 0 Then
                    .Selection.Range.HighlightColorIndex = wdYellow
                    .Selection.TypeText "Name"
                End If
                If StrComp(rng.Text, "NameOfContractMISSING ") = 0 Then
                    .Selection.Range.HighlightColorIndex = wdYellow
                    .Selection.TypeText "'Name of The Contract' "
                End If
                If StrComp(rng.Text, "Note") = 0 Then
                    .Selection.Font.Bold = True
                End If
                If StrComp(rng.Text, "ContractContract ") = 0 Then
                    .Selection.TypeText ""
                End If
                If StrComp(rng.Text, "REASONFORMAT") = 0 Then
                    .Selection.TypeText " "
                    .Selection.Range.ParagraphFormat.IndentFirstLineCharWidth (3)
                    .Selection.Range.ListFormat.ApplyOutlineNumberDefault (1)
                End If
                
                'update progres bar status
                ufProg.LabelCaption.Caption = Int((cur / totalWords) * 100) & "% Complete"
                ufProg.LabelProg.Width = (cur / totalWords) * (ufProg.FrameProg.Width)
                DoEvents
        Next rng
    End With
    'ReDim Preserve documentList(0)
    'documentList(0) = "Document" + CStr(documentNumber + 1)
    'documentNumber = documentNumber + 1
    'ReDim Preserve documentList(UBound(documentList) - LBound(documentList) + 1)
    WordApp.Application.ScreenUpdating = True
    WordApp.ActiveWindow.WindowState = wdWindowStateMaximize
    
    Set WordApp = Nothing
End If
endSub:
ufProg.LabelCaption.Caption = "Refreshing Data..."
refreshQuery
    'ThisWorkbook.Save
    'ThisWorkbook.SaveAs Filename:="Admin Evaluation Tracker (V2) Auto-Save"
    Unload ufProg
    Application.ScreenUpdating = True
    
End Sub
Function formatVendorEmail(v As Variant, template As String, headings() As String, vowels() As String, vendorIndexer As Integer) As String

'This function formats the vendor email based on the information given about the vendor
'The email is formated as either HTML or text based on the requested output

Dim vendorInfo As String, notRQ As String, correct As String, incorrect As String, dueBy As String, dt As String, qt() As String, tempVAR As String, qIndex As Integer, tempWs As Worksheet, res As String, quarters() As String, temp() As String, flag As Boolean
Set tempWs = ThisWorkbook.Worksheets("Validation Sheet")
    
    For qIndex = 0 To UBound(v.quarter)
            tempVAR = v.Iquarter(qIndex)
            If Not StrComp(v.Iquarter(qIndex), "N/A") = 0 Then
                If Not StrComp(v.Iquarter(qIndex), "") = 0 Then
                    GoTo verif
                End If
            End If
        Next qIndex
        GoTo nextvendor
verif:
        res = template
        vendorInfo = ""
        notRQ = ""
        correct = ""
        incorrect = ""
        quarters = v.quarter
        'Gather email requirements for current vendor
        'Subtract number of personal columns (inclusive)
        Dim i As Integer
        For i = 0 To numCols - 6
            If StrComp(quarters(i), "Not Requested") = 0 Then
                If buttonValue = 2 Then
                    'word format
                    notRQ = notRQ & CStr(headings(i)) & " BULLET" & vbNewLine
                Else
                    'html format
                    notRQ = notRQ & "<li><p style = 'margin-left:18.0pt'>" & CStr(headings(i)) & "</p></li>"
                End If
                totalWrongReasons = totalWrongReasons + 1
            End If
            If StrComp(quarters(i), "Submitted") = 0 Then
                If buttonValue = 2 Then
                    'word format
                    correct = correct & CStr(headings(i)) & " BULLET" & vbNewLine
                Else
                    'html format
                    correct = correct & "<li><p style = 'margin-left:18.0pt'>" & CStr(headings(i)) & "</p></li>"
                End If
                totalWrongReasons = totalWrongReasons + 1
            End If
            If StrComp(quarters(i), "Submitted Incorrectly") = 0 Then
                If buttonValue = 2 Then
                    'word format
                    incorrect = incorrect & CStr(headings(i)) & " BULLET" & vbNewLine
                Else
                    'html format
                    incorrect = incorrect & "<li><p style = 'margin-left:18.0pt'>" & CStr(headings(i)) & "</p></li>"
                End If
                totalWrongReasons = totalWrongReasons + 1
                temp = Split(CStr(Worksheets("Submitted Incorrectly Reasons").Cells(vendorIndexer, i + 2).Value), ",")
                Dim t As Variant
                If buttonValue = 0 Or buttonValue = 1 Then
                    incorrect = incorrect & "<ol type = 'i' list-style-type: decimal>"
                End If
                For Each t In temp
                    If buttonValue = 2 Then
                        incorrect = incorrect & t & " REASONFORMAT" & vbNewLine
                    Else
                        incorrect = incorrect & "<li><p>" & t & "</p></li>"
                    End If
                    totalWrongReasons = totalWrongReasons + 1
                Next t
                If buttonValue = 0 Or buttonValue = 1 Then
                    incorrect = incorrect & "</ol>"
                End If
            End If
            If StrComp(headings(i), "Due By") = 0 Then
                dueBy = CStr(quarters(i))
            End If
        Next i
         
        'Replace specific vendor information and generate specific email
        
        If tempWs.Range("C1").Value >= 12 Then
            res = Replace(res, "(DO NOT REPLACE ME!! 7)", "afternoon")
        Else
            res = Replace(res, "(DO NOT REPLACE ME!! 7)", "morning")
        End If
        
        res = Replace(res, "(DO NOT REPLACE ME!! 8)", CStr(v.name))
        
        'check for correct vowel
        Dim vo As Variant
        For Each vo In vowels
            flag = False
            If StrComp(Left(CStr(v.contract), 1), vo) = 0 Then
                flag = True
                GoTo found
            End If
        Next vo
found:
        If flag Then
            res = Replace(res, "(DO NOT REPLACE ME!! 1)", "an")
        Else
            res = Replace(res, "(DO NOT REPLACE ME!! 1)", "a")
        End If
        res = Replace(res, "(DO NOT REPLACE ME!! 2)", CStr(v.contract))
        res = Replace(res, "(DO NOT REPLACE ME!! 3)", correct)
        res = Replace(res, "(DO NOT REPLACE ME!! 4)", incorrect)
        res = Replace(res, "(DO NOT REPLACE ME!! 5)", notRQ)
        res = Replace(res, "(Insert Date)", dueBy)
        res = Replace(res, "(DO NOT REPLACE ME!! 6)", CStr(v.email))
        formatVendorEmail = res
nextvendor:
End Function
Private Sub generateHTML()
    Dim oIE As Object, oHEle As HTMLDivElement
    Dim oHDoc As HTMLDocument
    
    Set oIE = CreateObject("InternetExplorer.Application")
    
    With oIE
        .Visible = True
        .navigate "https://text-html.com"
    End With
    
    While oIE.readyState <> 4
        DoEvents
    Wend
 
    Set oHDoc = oIE.Document
    
    oHEle = oHDoc.getElementById("mc12")
    oHEle.getElementsByTagName("p").Item(1).innerHTML = "HELLO WORLD"
    
End Sub
