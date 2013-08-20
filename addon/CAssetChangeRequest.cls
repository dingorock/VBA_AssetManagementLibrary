VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAssetChangeRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'### Private instance variables ###
Private requestCode As String
Private requestType As String
Private requestStatus As String
Private AssetType As String
Private AssetGroup As String
Private InstanceNumber As String
Private location As String
Private parent As String
Private promotionState As String
Private changeAssessment As String
Private QualificationStatus As String
Private OperationalStatus As String
Private materialsCertificates As String
Private wHSAssessmentRef As String
Private toolRecordRef As String
Private iQRef As String
Private serviceReports As String
Private meterReading As String
Private meterReadDate As String
Private owningGroup As String
Private manufacturer As String
Private Model As String
Private serialNumber As String
Private appSoftVersion As String
Private serviceContractRef As String
Private warrantyExpiry As String
Private requestComments As String
Private requestor As String
Private createdBy As String
Private entryDate As String

Private pAllRequests As Variant
Private pRequest As Variant
Private pRequestRows As Integer


Private pStatusArray(4) As Variant

Public Property Get StatusArray() As Variant
    pStatusArray(0) = "New"
    pStatusArray(1) = "More Info Required"
    pStatusArray(2) = "In Progress"
    pStatusArray(3) = "Rejected"
    pStatusArray(4) = "Completed"
    StatusArray = pStatusArray
End Property


'### Getter Properties ###
Public Property Get AllRequests() As Variant
    AllRequests = pAllRequests
End Property

Public Property Get NumberOfRequests() As Integer
    Dim wks As Worksheet
    Dim lookupColumn As String
    Dim lookupColumnRange As String
    
    Set wks = ThisWorkbook.Worksheets("ACRF")
    lookupColumn = "B"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    pRequestRows = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    
    NumberOfRequests = pRequestRows
End Property

Private Sub Class_Initialize()

Dim wks As Worksheet
Set wks = ThisWorkbook.Worksheets("ACRF")

Dim lookupColumnBaseCell As String
lookupColumnBaseCell = "A2"

ReDim pAllRequests(NumberOfRequests, 29)

pAllRequests(0, 0) = "requestCode"
pAllRequests(0, 1) = "requestType"
pAllRequests(0, 2) = "requestStatus"
pAllRequests(0, 3) = "assetType"
pAllRequests(0, 4) = "assetGroup"
pAllRequests(0, 5) = "instanceNumber"
pAllRequests(0, 6) = "location"
pAllRequests(0, 7) = "parent"
pAllRequests(0, 8) = "promotionState"
pAllRequests(0, 9) = "changeAssessment"
pAllRequests(0, 10) = "qualificationStatus"
pAllRequests(0, 11) = "operationalStatus"
pAllRequests(0, 12) = "materialsCertificates"
pAllRequests(0, 13) = "wHSAssessmentRef"
pAllRequests(0, 14) = "toolRecordRef"
pAllRequests(0, 15) = "iQRef"
pAllRequests(0, 16) = "serviceReports"
pAllRequests(0, 17) = "meterReading"
pAllRequests(0, 18) = "meterReadDate"
pAllRequests(0, 19) = "owningGroup"
pAllRequests(0, 20) = "manufacturer"
pAllRequests(0, 21) = "model"
pAllRequests(0, 22) = "serialNumber"
pAllRequests(0, 23) = "appSoftVtoolRecordRefersion"
pAllRequests(0, 24) = "serviceContractRef"
pAllRequests(0, 25) = "warrantyExpiry"
pAllRequests(0, 26) = "requestComments"
pAllRequests(0, 27) = "requestor"
pAllRequests(0, 28) = "createdBy"
pAllRequests(0, 29) = "entryDate"


For i = 0 To (NumberOfRequests - 1) Step 1
    For j = 0 To 29 Step 1
        ' i & j are reversed so the parameters are listed vertically,
        ' not horizontally like the acrf raw data table
        pAllRequests(i + 1, j) = wks.Range(lookupColumnBaseCell).Offset(i, j)
    Next
Next

End Sub

Public Function RequestByCode(requestCode As String) As Variant

Dim requestArray() As Variant
ReDim requestArray(1, 29)
Dim requestRow As Integer
requestRow = 0

For i = 0 To NumberOfRequests Step 1
    If (pAllRequests(i, 0) = requestCode) Then
        requestRow = i
        Exit For
    End If
Next

For j = 0 To 29 Step 1
    requestArray(0, j) = pAllRequests(0, j)
    requestArray(1, j) = pAllRequests(requestRow, j)
Next
    
RequestByCode = requestArray

End Function

Public Function RequestsForLookupCombobox() As Variant

Dim requestsArray() As Variant
ReDim requestsArray(NumberOfRequests - 1, 3)
Dim requestRow As Integer

For i = 1 To pRequestRows Step 1
    For j = 0 To 29 Step 1
        Select Case (j)
            Case 0
                requestsArray(i - 1, 0) = pAllRequests(i, j)
            Case 4
                requestsArray(i - 1, 1) = pAllRequests(i, j)
            Case 27
                requestsArray(i - 1, 2) = pAllRequests(i, j)
            Case 26
                requestsArray(i - 1, 3) = pAllRequests(i, j)
        End Select
    Next
Next

RequestsForLookupCombobox = requestsArray

End Function


Public Function RequestsForLookupComboboxFiltered(status As String) As Variant

If (status = "") Then
     RequestsForLookupComboboxFiltered = Me.RequestsForLookupCombobox
Else
    Dim requestsArray() As Variant
    Dim requestRow As Integer
    
    c = 0
    For k = 0 To (NumberOfRequests - 1)
        If (pAllRequests(k, 2) = status) Then
            c = c + 1
        End If
    Next
    
    If (c <> 0) Then
        ReDim requestsArray(c - 1, 3)
        
        c = 0
        For i = 0 To (NumberOfRequests - 1)
            If (pAllRequests(i, 2) = status) Then
                For j = 0 To 29 Step 1
                    Select Case (j)
                        Case 0
                            requestsArray(c, 0) = pAllRequests(i, j)
                        Case 4
                            requestsArray(c, 1) = pAllRequests(i, j)
                        Case 27
                            requestsArray(c, 2) = pAllRequests(i, j)
                        Case 26
                            requestsArray(c, 3) = pAllRequests(i, j)
                    End Select
                Next
                c = c + 1
            End If
        Next
    Else
        ReDim requestsArray(0, 3)
        requestsArray(0, 0) = ""
        requestsArray(0, 1) = ""
        requestsArray(0, 2) = ""
        requestsArray(0, 3) = ""
    End If
    RequestsForLookupComboboxFiltered = requestsArray
End If


End Function

Public Function HeadingsForLookupCombobox() As Variant

Dim requestsArray(0, 3) As Variant
'ReDim requestsArray(NumberOfRequests, 3)
Dim requestRow As Integer

For j = 0 To 29 Step 1
    Select Case (j)
        Case 0
            requestsArray(0, 0) = pAllRequests(i, j)
        Case 4
            requestsArray(0, 1) = pAllRequests(i, j)
        Case 27
            requestsArray(0, 2) = pAllRequests(i, j)
        Case 26
            requestsArray(0, 3) = pAllRequests(i, j)
    End Select
Next

HeadingsForLookupCombobox = requestsArray

End Function

