VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CeamData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'### Private instance variables ###
Private pManualActivities() As Variant
Private pAssetCategories() As Variant
Private pAssetTypes() As Variant
Private pAssetOwners() As Variant
Private pVirtualAssets() As Variant
Private pAllAssets() As Variant

'### Getter Properties ###
Public Property Get ManualActivities() As Variant
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")
    
    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "T"
    lookupColumnBaseCell = lookupColumn & "2"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
        
    Dim columnCount As Integer
    columnCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    
    ReDim pManualActivities(columnCount)
    
    For i = 0 To columnCount
        pManualActivities(i) = Range(lookupColumnBaseCell).Offset(i)
    Next

    ManualActivities = pManualActivities
End Property

Public Property Get AssetCategories() As Variant
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")
    
    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "A"
    lookupColumnBaseCell = lookupColumn & "2"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
        
    Dim columnCount As Integer
    columnCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    
    ReDim pAssetCategories(columnCount)
    
    For i = 0 To columnCount
        pAssetCategories(i) = Range(lookupColumnBaseCell).Offset(i)
    Next

    AssetCategories = pAssetCategories
End Property

Public Property Get AssetTypes() As Variant
    ReDim pAssetTypes(1)
    pAssetTypes(0) = "Capital"
    pAssetTypes(1) = "Rebuildable"
    AssetTypes = pAssetTypes
End Property

Public Property Get AssetOwners() As Variant
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")
    
    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "B"
    lookupColumnBaseCell = lookupColumn & "2"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
        
    Dim columnCount As Integer
    columnCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    
    ReDim pAssetOwners(columnCount, 1)
    
    For i = 0 To columnCount
        pAssetOwners(i, 0) = Range(lookupColumnBaseCell).Offset(i, 0)
        pAssetOwners(i, 1) = Range(lookupColumnBaseCell).Offset(i, 1)
    Next

    AssetOwners = pAssetOwners
End Property

Public Property Get VirtualAssets() As Variant
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")
    
    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "D"
    lookupColumnBaseCell = lookupColumn & "2"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
        
    Dim columnCount As Integer
    columnCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    
    ReDim pVirtualAssets(columnCount, 1)
    
    For i = 0 To columnCount
        pVirtualAssets(i, 0) = Range(lookupColumnBaseCell).Offset(i, 0)
        pVirtualAssets(i, 1) = Range(lookupColumnBaseCell).Offset(i, 1)
    Next

    VirtualAssets = pVirtualAssets
End Property

Public Property Get AllAssets() As Variant
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("ASSETS")
    
    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "A"
    lookupColumnBaseCell = lookupColumn & "2"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
        
    Dim columnCount As Integer
    columnCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    ReDim pAllAssets(columnCount, 4)
    
    For i = 0 To columnCount
        For j = 0 To 4
            pAllAssets(i, j) = wks.Range(lookupColumnBaseCell).Offset(i, j)
        Next
    Next
    
    AllAssets = pAllAssets
End Property

Function allAssetsForAssetGroup(AssetGroup As String) As Variant
'Dim pAllAssetsForAssetGroup() As Variant
Dim relevantRowNumbers() As Variant
Dim tempAllAssets() As Variant
tempAllAssets = Me.AllAssets

Dim counter As Integer
counter = 0

'get the row numbers that match
For i = 0 To UBound(tempAllAssets)
    If (tempAllAssets(i, 0) = AssetGroup) Then
        ReDim Preserve relevantRowNumbers(counter)
        relevantRowNumbers(counter) = i
        counter = counter + 1
    End If
Next

'load matched rows into new pallassets array
If (counter) Then
    ReDim pAllAssetsForAssetGroup(UBound(relevantRowNumbers), 4)
    For k = 0 To UBound(relevantRowNumbers)
        For j = 0 To 4
               pAllAssetsForAssetGroup(k, j) = tempAllAssets(relevantRowNumbers(k), j)
        Next
    Next
Else
    ReDim pAllAssetsForAssetGroup(0, 4)
End If

allAssetsForAssetGroup = pAllAssetsForAssetGroup
End Function

Sub xmlTest()

Dim xmlObject As MSXML2.DOMDocument60
Set xmlObject = New MSXML2.DOMDocument60


End Sub

