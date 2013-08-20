VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### INTERNAL CLASS INSTANCE VARIABLE DECLARATIONS ###

' ### PROPERTY DECLARATIONS FOR PUBLIC ACCESS VIA GETTERS & SETTERS ###
Private pDocNumber As String
Private pAssetGroupNumber As String
Private pAssetType As String
Private pName As String
Private pOwner As String
Private pDepartment As String
Private pCategory As String
Private pActivities() As Variant
Private pNewActivities() As Variant

Private pNextInstanceNumber As Integer
Private pNumberOfActivities As Integer


' ##DOCUMENT NUMBER REFERENCE
Public Property Let DocNumber(value As String)
    ' Check the length of the document reference
    If (Len(value) <> 6) Then
        While (Len(value) <> 6)
            value = InputBox("Document Number [ " & value & " ] is not a valid Windchill Document reference")
        Wend
    End If
    pDocNumber = value
End Property

Public Property Get DocNumber() As String
    DocNumber = pDocNumber
End Property

' ## ASSET GROUP NUMBER FROM WINDCHILL
Public Property Let AssetGroupNumber(value As String)

' Validation of input
Dim prefixValidated As Boolean
Dim lengthValidated As Boolean

' Initialise to False (not validated)
prefixValidated = False
lengthValidated = False

' #PREFIX VALIDATION
' Get the asset group prefix character W, T or K
Dim assetGroupPrefix As String
assetGroupPrefix = UCase(Left(value, 1))

Select Case (assetGroupPrefix)
    Case "T", "W", "K"
        prefixValidated = True
End Select

' #LENGTH VALIDATION
' Check the tool number format
If (Len(value) >= 6 _
  And Len(value) <= 7) Then
    lengthValidated = True
End If

' Final check of all validation test results
Dim validationResult As String
If (prefixValidated And lengthValidated) Then
    pAssetGroupNumber = value
Else
    pAssetGroupNumber = "TXXXXXX"
    
    validationResult = "Asset Group Number [ " & value & " ] is not a valid Windchill tool reference" & vbNewLine
    If (Not (prefixValidated)) Then validationResult = validationResult & vbNewLine & "- Prefix is not T, W or K"
    If (Not (lengthValidated)) Then validationResult = validationResult & vbNewLine & "- Incorrect length"
    
    msgboxResult = MsgBox(Prompt:=validationResult, _
            Buttons:=vbExclamation, _
            Title:="Input Error")
End If

End Property

Public Property Get AssetGroupNumber() As String
    AssetGroupNumber = pAssetGroupNumber
End Property

' #ASSET TYPE -- Capital or Rebuildable
Public Property Let AssetType(value As String)

Dim assetTypeValidation As Boolean
Dim validationResult As String
assetTypeValidation = False

' Create EAM DATA object and get an array of asset types
Dim eamData As CeamData
Set eamData = New CeamData
Dim AssetTypes() As Variant
AssetTypes = eamData.AssetTypes


For Each listItem In AssetTypes
    If (value = listItem) Then
        assetTypeValidation = True
        Exit For
    End If
Next

If (assetTypeValidation) Then
    pAssetType = value
Else
    pAssetType = "CAPITAL"
    
    validationResult = "Asset Type [ " & value & " ] is not a valid type." & vbNewLine & "Must be of type Capital or Rebuildable" _
                        & vbNewLine & "Setting type to CAPITAL as fallback"
    msgboxResult = MsgBox(Prompt:=validationResult, _
            Buttons:=vbExclamation, _
            Title:="Input Error")
End If

End Property

Public Property Get AssetType() As String
    AssetType = pAssetType
End Property

' #ASSET NAME
Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get Name() As String
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("ASSETS")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String

    lookupColumn = "A"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = "A2"
    
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    
    On Error Resume Next
    pName = WorksheetFunction.VLookup(pAssetGroupNumber, wks.Range(lookupListlookupColumnBaseCell & ":D" & rowCount), 4, 0)
    
    If (pName = "" Or pName = "#N/A") Then
        
        If (MsgBox("Open Part in Windchill to get the Tool name???", vbYesNo, "Can't find tool name: Check in Windchill?") = vbYes) Then
                    ActiveWorkbook.FollowHyperlink (Me.hyperlinkToAssetGroup)
        End If
        validationResult = "Name not found. Please enter a name for this asset group."
        pName = InputBox(Prompt:=validationResult, _
                Title:="Name not found")
    End If
    
    Name = pName
    
End Property


' #ASSET OWNER
Public Property Let Owner(value As String)
Dim assetOwnerValidation As Boolean
Dim validationResult As String
assetOwnerValidation = True

If (assetOwnerValidation) Then
    pOwner = value
Else
    pOwner = ""

    msgboxResult = MsgBox(Prompt:=validationResult, _
            Buttons:=vbExclamation, _
            Title:="Input Error")
End If

End Property

Public Property Get Owner() As String
    Owner = pOwner
End Property

Public Property Get Department() As String
    pDepartment = Left(pOwner, WorksheetFunction.Search(".", pOwner) - 1)
    Department = pDepartment
End Property


' #ASSET CATEGORY
Public Property Let Category(value As String)
Dim categoryValidation As Boolean
Dim validationResult As String
categoryValidation = False

Dim wks As Worksheet
Set wks = ThisWorkbook.Worksheets("Lookups")

Dim lookupColumn As String
Dim lookupColumnBaseCell As String
Dim lookupColumnRange As String

lookupColumn = "A"
lookupColumnRange = lookupColumn & ":" & lookupColumn
lookupListlookupColumnBaseCell = "A2"

Dim rowCount As Integer
rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1

Dim categoryList() As Variant
ReDim categoryList(rowCount)

For i = 0 To rowCount - 1 Step 1
    categoryList(i) = wks.Range(lookupListlookupColumnBaseCell).Offset(i)
    'Validate value, match against any category in list?
    If (value = categoryList(i)) Then
        categoryValidation = True
    End If
Next

If (categoryValidation) Then
    pCategory = value
Else
    pCategory = "Unknown"
    validationResult = "Unknown Category [ " & value & " ]."
    
    msgboxResult = MsgBox(Prompt:=validationResult, _
    Buttons:=vbExclamation, _
    Title:="Input Error")
End If

End Property

Public Property Get Category() As String
Category = pCategory
End Property

' #ASSET ACTIVITIES
Public Property Get Activities() As Variant
    
If (IsEmpty(pActivities(0))) Then
   
    Dim assetActivitiesValidation As Boolean
    Dim validationResult As String
    assetActivitiesValidation = True

    ' Get activities list
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("SCHEDULES_PIVOT")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "D"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = "C4"

    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1

    Dim activitiesCountForAssetGroup As Integer
    activitiesCountForAssetGroup = 0
    
    For i = 0 To rowCount Step 1
        tempAssetGroup = wks.Range(lookupListlookupColumnBaseCell).Offset(i, 0).value
        If (tempAssetGroup <> "") Then
            If (tempAssetGroup = pAssetGroupNumber) Then
                activitiesCountForAssetGroup = wks.Range(lookupListlookupColumnBaseCell).Offset(i, -1).value
                ReDim pActivities(activitiesCountForAssetGroup - 1)
                For j = 0 To (activitiesCountForAssetGroup - 1)
                    pActivities(j) = wks.Range(lookupListlookupColumnBaseCell).Offset(i + j + 1, 4).value
                Next
                Exit For
            End If
        End If
    Next
    If (activitiesCountForAssetGroup = 0) Then
        ReDim pActivities(0)
        pActivities(0) = ""
    End If
End If

Activities = pActivities

End Property

Public Function activityArray(theActivityCode As String) As Variant

    Dim theActivity As CActivity
    Set theActivity = New CActivity
    
    Dim tempArray(1, 16) As Variant
    'Headings
    tempArray(0, 0) = "Activity Code"
    tempArray(0, 1) = "Activity Description"
    tempArray(0, 2) = "Activity Type"
    tempArray(0, 3) = "Asset Group"
    tempArray(0, 4) = "Cause"
    tempArray(0, 5) = "Inventory CategorySet"
    tempArray(0, 6) = "Maintenance Routing"
    tempArray(0, 7) = "OSP"
    tempArray(0, 8) = "OSPLeadtime Op"
    tempArray(0, 9) = "OSPService Op"
    tempArray(0, 10) = "Planned"
    tempArray(0, 11) = "Service Interval"
    tempArray(0, 12) = "Service Vendor"
    tempArray(0, 13) = "Shutdown Required"
    tempArray(0, 14) = "Source"
    tempArray(0, 15) = "Template"

'    If (theActivityCode <> "<new activity>") Then
'        theActivity.initialiseActivityWithActivityCode (theActivityCode)

        'Data
        tempArray(1, 0) = theActivityCode
        tempArray(1, 1) = ""
        tempArray(1, 2) = theActivity.ActivityType
        tempArray(1, 3) = theActivity.AssetGroup
        tempArray(1, 4) = theActivity.Cause
        tempArray(1, 5) = theActivity.InventoryCategorySet
        tempArray(1, 6) = ""
        tempArray(1, 7) = theActivity.OSP
        tempArray(1, 8) = theActivity.OSPLeadtimeOp
        tempArray(1, 9) = theActivity.OSPserviceOp
        tempArray(1, 10) = theActivity.Planned
        tempArray(1, 11) = theActivity.ServiceInterval
        tempArray(1, 12) = theActivity.ServiceVendor
        tempArray(1, 13) = theActivity.ShutdownRequired
        tempArray(1, 14) = theActivity.Source
        tempArray(1, 15) = theActivity.Template
        tempArray(1, 16) = theActivity.TypeAttribute
'    End If
    
    activityArray = tempArray

End Function

Public Function scheduleArray() As Variant

    Dim tempArray(1, 10) As Variant
    'Headings
    tempArray(0, 0) = "ScheduleCode"
    tempArray(0, 1) = "AssetInstance"
    tempArray(0, 2) = "Activity"
    tempArray(0, 3) = "EffectiveFrom"
    tempArray(0, 4) = "EffectiveTo"
    tempArray(0, 5) = "RunToFailure"
    tempArray(0, 6) = "Interval"
    tempArray(0, 7) = "LastServiceDateActualStart"
    tempArray(0, 8) = "LastServiceDateActualEnd"
    tempArray(0, 9) = "LastServiceDateScheduledStart"
    tempArray(0, 10) = "pLastServiceDateScheduledEnd"
    'Data
    tempArray(1, 0) = ScheduleCode
    tempArray(1, 1) = assetInstance
    tempArray(1, 2) = Activity.activityCode
    tempArray(1, 3) = EffectiveFrom
    tempArray(1, 4) = EffectiveTo
    tempArray(1, 5) = RunToFailure
    tempArray(1, 6) = interval
    tempArray(1, 7) = LastServiceDateActualStart
    tempArray(1, 8) = LastServiceDateActualEnd
    tempArray(1, 9) = LastServiceDateScheduledStart
    tempArray(1, 10) = LastServiceDateScheduledEnd
    
    scheduleArray = tempArray

End Function

Public Property Get NumberOfActivities() As Integer
    If (pNumberOfActivities = -1) Then
        pNumberOfActivities = (UBound(Me.Activities, 1)) + 1
    End If
    
    NumberOfActivities = pNumberOfActivities
End Property


'## NEW ACTIVITY MANAGEMENT
Public Property Get NewActivities() As Variant
'    If (IsEmpty(pNewActivities(0))) Then
'
'    End If
    NewActivities = pNewActivities
End Property
Public Sub addActivity()
    If (Not (IsEmpty(pNewActivities(0)))) Then
        ReDim Preserve pNewActivities(UBound(pNewActivities) + 1)
    End If
    pNewActivities(UBound(pNewActivities, 1)) = "<new activity>"
End Sub

Public Sub removeActivity(activityIndex As Integer)
    If (IsEmpty(pNewActivities(0))) Then
        Exit Sub
    Else
        Dim newArray() As Variant
        Dim k As Integer
        k = 0
        If (pNewActivities(0) <> "" And UBound(pNewActivities, 1) > 0) Then
            ReDim newArray(UBound(pNewActivities, 1) - 1)
            For i = 0 To UBound(pNewActivities, 1)
                If (i = activityIndex) Then
                    'Skip i
                    i = i + 1
                End If
                newArray(k) = pNewActivities(i)
                k = k + 1
            Next
        Else
            Dim anArray() As Variant
            ReDim anArray(0)
            newArray = anArray
        End If
    End If
    pNewActivities = newArray
End Sub


' ### Instance Methods ###
Public Function hyperlinkToAssetGroup() As String
    hyperlinkToAssetGroup = "http://cltd-windchill.cochlear.com/Windchill/netmarkets/jsp/cochlear/DocByBasisRef.jsp?basisref=" & _
        pAssetGroupNumber
End Function
Public Function hyperlinkToSpecificationDoc() As String
    hyperlinkToSpecificationDoc = "http://cltd-windchill.cochlear.com/Windchill/netmarkets/jsp/cochlear/DocByBasisRef.jsp?basisref=" & _
        pDocNumber
End Function

Public Property Get NextInstanceNumber() As Integer
    Dim eamData As CeamData
    Set eamData = New CeamData
    Dim assetsForAssetGroup() As Variant

    assetsForAssetGroup = eamData.allAssetsForAssetGroup(pAssetGroupNumber)
    pNextInstanceNumber = 1
    If IsEmpty(assetsForAssetGroup(0, 0)) Then
        pNextInstanceNumber = 0
    Else
        For i = 0 To UBound(assetsForAssetGroup)
            If (assetsForAssetGroup(i, 4) > pNextInstanceNumber) Then
                pNextInstanceNumber = CInt(assetsForAssetGroup(i, 4))
            End If
        Next
    End If
    
    pNextInstanceNumber = pNextInstanceNumber + 1
    NextInstanceNumber = pNextInstanceNumber
    
End Property

' ### Public Class Methods ###
Public Sub initialiseAssetGroupFromRequest(selectedRequest As Variant)

pNumberOfActivities = -1
ReDim pActivities(0)
ReDim pNewActivities(0)

Me.AssetGroupNumber = selectedRequest(1, 4)
Me.DocNumber = selectedRequest(1, 14)
Me.AssetType = selectedRequest(1, 3)
Me.Owner = selectedRequest(1, 19)
'Me.Category = ""

'For Each code In Me.Activities
'    Debug.Print (code)
'Next

End Sub


Public Sub initialiseAssetGroupFromAssetGroupNumber(AssetGroup As String)
pNumberOfActivities = -1
ReDim pActivities(0)
ReDim pNewActivities(0)

Me.AssetGroupNumber = AssetGroup
End Sub


