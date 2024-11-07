VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Enea Zguro
'OGS Procurement Services
'9/10/24
'This class is meant to hold vendor information and contract information per Vendor

Dim vContract As String, vName As String, vQuarters() As String, vEmail As String, vRow As Integer, vStatus As Boolean

Public Property Get name() As String
    name = vName
End Property

Public Property Let name(newName As String)
    vName = newName
End Property

Public Property Get row() As Integer
    row = vRow
End Property

Public Property Let row(rowSet As Integer)
    vRow = rowSet
End Property

Public Property Get status() As Boolean
    status = vStatus
End Property

Public Property Let status(statusSet As Boolean)
    vStatus = statusSet
End Property


Public Property Get email() As String
    email = vEmail
End Property

Public Property Let email(newEmail As String)
    vEmail = newEmail
End Property

Public Property Get contract() As String
    contract = vContract
End Property

Public Property Let contract(newContract As String)
    vContract = newContract
End Property

Public Property Get quarter() As String()
    quarter = vQuarters
End Property
Public Property Get Iquarter(index As Integer) As String
    Iquarter = vQuarters(index)
End Property


Public Property Let setQuarter(index As Integer, submission As String)
    If UBound(vQuarters) < index Then
        ReDim Preserve vQuarters(UBound(vQuarters) - LBound(vQuarters) + 1)
    End If
    vQuarters(index) = submission
    'ReDim Preserve vQuarters(UBound(vQuarters) - LBound(vQuarters) + 1)
End Property
Private Sub Vendor()
    vContract = "Missing Contract"
    vName = "Missing Name"
    vEmail = "Missing Email"
    vStatus = False
End Sub

Private Sub Class_Initialize()
    vName = "MissingMISSING NameMISSING"
    vContract = "MissingMISSING NameOfContractMISSING"
    vEmail = "MissingMISSING EmailMISSING"
    vStatus = False
    ReDim Preserve vQuarters(0)
End Sub


