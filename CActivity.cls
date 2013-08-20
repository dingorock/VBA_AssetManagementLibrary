VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### INTERNAL CLASS INSTANCE VARIABLE DECLARATIONS ###

' ### PROPERTY DECLARATIONS FOR PUBLIC ACCESS VIA GETTERS & SETTERS ###
Private pActivityCode As String
Private pActivityDescription As String
Private pActivityType As String
Private pServiceInterval As String
Private pPlanned As Boolean

Private pServiceVendor As String
Private pOSP As Boolean
Private pOSPLeadtimeOp As String
Private pOSPServiceOp As String

Private pMaintenanceRouting() As Variant

Private pTypeAttribute As String
Private pCause As String
Private pSource As String
Private pShutdownRequired As Boolean

Private pAssetGroup As String

' eAM STATIC READONLY VARIABLES
Private pTemplate As String
Private pInventoryCategorySet As String

' ##ACTIVITY CODE
Public Property Let activityCode(value As String)
    pActivityCode = value
End Property
Public Property Get activityCode() As String
    activityCode = pActivityCode
End Property

'## ACTIVITY DESCRIPTION
Public Property Get ActivityDescription() As String
    If (pPlanned And Me.ActivityType <> "Not Defined") Then
            pActivityDescription = ActivityType & " - " & _
                                pServiceInterval & " Monthly - " & _
                                Right(pActivityCode, Len(pActivityCode) - WorksheetFunction.Search("-", pActivityCode))
    Else
        'Manual Activity Description
        pActivityDescription = pActivityType & " (Manual)"
    End If
                                
    ActivityDescription = pActivityDescription
End Property

' ##ACTIVITY TYPE
Public Property Let ActivityType(value As String)

    Dim activityTypeValidation As Boolean
    Dim validationResult As String
    activityTypeValidation = False

    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "F"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = lookupColumn & "2"
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    Dim activitiesList() As String
    ReDim activitiesList(rowCount - 1)
    
    For i = 0 To rowCount Step 1
        activitiesList(i) = wks.Range(lookupListlookupColumnBaseCell).Offset(i)
        If (value = activitiesList(i)) Then activityTypeValidation = True
    Next
    
    If (activityTypeValidation) Then
        pActivityType = value
    Else
        validationResult = "Activity Type [ " & value & " ] is not valid. Use one of the following:" & vbNewLine
        For i = 0 To rowCount Step 1
            validationResult = validationResult & "-" & activitiesList(i) & vbNewLine
        Next
        pActivityType = InputBox(Prompt:=validationResult, Title:="Activity Type Error")
    End If
    
End Property
Public Property Get ActivityType() As String
    Dim typePrefix As String
    typePrefix = Left(pActivityCode, 1)
    Select Case (typePrefix)
        Case "C"
            pActivityType = "Calibration"
        Case "M"
            pActivityType = "Maintenance"
        Case "I"
            pActivityType = "Inspection"
        Case Else
            pActivityType = "Not Defined"
    End Select

    ActivityType = pActivityType
End Property


' ## SERVICE INTERVAL
Public Property Get ServiceInterval() As Integer
    If (pActivityCode <> "" And pActivityType <> "Not Defined") Then
        pServiceInterval = CInt(Mid(pActivityCode, 2, WorksheetFunction.Search("-", pActivityCode) - 2))
    Else
        pServiceInterval = 0
    End If
    ServiceInterval = pServiceInterval
End Property

' ##EXTERNAL VENDOR NAME
Public Property Let ServiceVendor(value As String)

    Dim serviceVendorValidation As Boolean
    Dim validationResult As String
    serviceVendorValidation = False

    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "K"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = lookupColumn & "2"
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    Dim vendorList() As String
    ReDim vendorList(rowCount - 1)
    
    For i = 0 To rowCount Step 1
        vendorList(i) = wks.Range(lookupListlookupColumnBaseCell).Offset(i)
        If (value = vendorList(i)) Then serviceVendorValidation = True
    Next
    
    If (serviceVendorValidation) Then
        pServiceVendor = value
    Else
        validationResult = "Vendor [ " & value & " ] is not valid. Use one of the following:" & vbNewLine
        For i = 0 To rowCount Step 1
            validationResult = validationResult & "-" & vendorList(i) & vbNewLine
        Next
        pServiceVendor = InputBox(Prompt:=validationResult, Title:="Vendor Error")
    End If

    pServiceVendor = value
End Property
Public Property Get ServiceVendor() As String
    ServiceVendor = pServiceVendor
End Property

' ##OSP ACTIVITY
Public Property Let OSP(value As Boolean)
    If (pOSP = "") Then
        pOSP = False
    Else
        pOSP = value
    End If
End Property
Public Property Get OSP() As Boolean
    If (pServiceVendor <> "") Then pOSP = True
    OSP = pOSP
End Property

' ##OSP Leadtime Standard Operation Code
Public Property Let OSPLeadtimeOp(value As String)
If (pOSP) Then
    Dim leadtimeOpValidation As Boolean
    Dim validationResult As String
    leadtimeOpValidation = False

    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "I"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = lookupColumn & "2"
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    Dim leadtimeOpList() As String
    
    ReDim leadtimeOpList(rowCount - 1)
    For i = 0 To rowCount Step 1
        leadtimeOpList(i) = wks.Range(lookupListlookupColumnBaseCell).Offset(i)
        If (value = leadtimeOpList(i)) Then leadtimeOpValidation = True
    Next
    
    If (leadtimeOpValidation) Then
        pOSPLeadtimeOp = value
    Else
        validationResult = "Leadtime Operation [ " & value & " ] is not valid. Use one of the following:" & vbNewLine
        For i = 0 To rowCount Step 1
            validationResult = validationResult & "-" & leadtimeOpList(i) & vbNewLine
        Next
        pOSPLeadtimeOp = InputBox(Prompt:=validationResult, Title:="Leadtime Operation Error")
    End If
Else
    pOSPLeadtimeOp = ""
End If
    
End Property
Public Property Get OSPLeadtimeOp() As String
If (pOSP) Then
    Dim leadtimeOpValidation As Boolean
    Dim validationResult As String
    leadtimeOpValidation = False
    
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "I"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = lookupColumn & "2"
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    Dim leadtimeOpList() As String
    
    ReDim leadtimeOpList(0, 1)
    For i = 0 To rowCount Step 1
        If (Left(wks.Range(lookupListlookupColumnBaseCell).Offset(i), 1) = "L") Then
            ReDim Preserve leadtimeOpList(i, 1)
            leadtimeOpList(i, 0) = wks.Range(lookupListlookupColumnBaseCell).Offset(i, 0)
            leadtimeOpList(i, 1) = wks.Range(lookupListlookupColumnBaseCell).Offset(i, 1)
        End If
    Next
    
    Dim tempOp As String
    For i = 0 To UBound(leadtimeOpList, 1) Step 1
        tempOp = Right(leadtimeOpList(i, 1), Len(leadtimeOpList(i, 1)) - WorksheetFunction.Search("-", leadtimeOpList(i, 1)))
        If (pServiceVendor = tempOp) Then
            leadtimeOpValidation = True
            pOSPLeadtimeOp = leadtimeOpList(i, 1)
        End If
    Next
    
    If (leadtimeOpValidation) Then
        pOSPLeadtimeOp = value
    Else
        validationResult = "Leadtime Operation [ " & value & " ] is not valid. Use one of the following:" & vbNewLine
        For i = 0 To rowCount Step 1
            validationResult = validationResult & "-" & leadtimeOpList(i) & vbNewLine
        Next
        pOSPLeadtimeOp = InputBox(Prompt:=validationResult, Title:="Leadtime Operation Error")
    End If
Else
    pOSPLeadtimeOp = ""
End If

End Property

' ##OSP Service Standard Operation Code
Public Property Let OSPserviceOp(value As String)
If (pOSP) Then
    pOSPServiceOp = value
Else
    pOSPServiceOp = ""
End Property
Public Property Get OSPserviceOp() As String
If (pOSP) Then
    Dim serviceOpValidation As Boolean
    Dim validationResult As String
    serviceOpValidation = False
    
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "I"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = lookupColumn & "2"
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    Dim serviceOpList() As String
    
    ReDim serviceOpList(0, 1)
    For i = 0 To rowCount Step 1
        If (Left(wks.Range(lookupListlookupColumnBaseCell).Offset(i), 1) = Left(pActivityType, 1)) Then
            ReDim Preserve serviceOpList(i, 1)
            serviceOpList(i, 0) = wks.Range(lookupListlookupColumnBaseCell).Offset(i, 0)
            serviceOpList(i, 1) = wks.Range(lookupListlookupColumnBaseCell).Offset(i, 1)
        End If
    Next
    
    Dim tempOp As String
    For i = 0 To UBound(serviceOpList, 1) Step 1
        tempOp = Right(serviceOpList(i, 1), Len(serviceOpList(i, 1)) - WorksheetFunction.Search("-", serviceOpList(i, 1)))
        If (pServiceVendor = tempOp) Then
            serviceOpValidation = True
            pOSPServiceOp = serviceOpList(i, 1)
        End If
    Next
    
    If (serviceOpValidation) Then
        pOSPServiceOp = value
    Else
        validationResult = "Service Operation [ " & value & " ] is not valid. Use one of the following:" & vbNewLine
        For i = 0 To rowCount Step 1
            validationResult = validationResult & "-" & serviceOpList(i) & vbNewLine
        Next
        pOSPServiceOp = InputBox(Prompt:=validationResult, Title:="Serivce Operation Error")
    End If
Else
    pOSPServiceOp = ""
End If
    
OSPserviceOp = pOSPServiceOp

End Property

' ##Maintenance Routings - OP-10, 20, 30 -> STD-OPERATION CODE
Public Property Get MaintenanceRouting() As Variant
    
    If (OSP) Then
        ReDim pMaintenanceRouting(2, 1)
    Else
        ReDim pMaintenanceRouting(0, 1)
    End If
    
    For i = 0 To UBound(pMaintenanceRouting, 1)
        pMaintenanceRouting(i, 0) = CInt(i + 1 & 0)
    Next
        
    If (Not (pOSP)) Then
        Select Case (pActivityType)
            Case "Calibration"
                pMaintenanceRouting(0, 1) = "CAL"
            Case "Maintenance"
                pMaintenanceRouting(0, 1) = "MNT"
            Case "Inspection"
                pMaintenanceRouting(0, 1) = "INS"
        End Select
    Else
        For i = 0 To UBound(pMaintenanceRouting, 1)
            Select Case (pMaintenanceRouting(i, 0))
                Case 10
                    pMaintenanceRouting(i, 1) = pOSPLeadtimeOp
                Case 20
                    pMaintenanceRouting(i, 1) = pOSPServiceOp
                Case 30
                    pMaintenanceRouting(i, 1) = pOSPLeadtimeOp
                    Select Case (pActivityType)
                        Case "Calibration"
                            pMaintenanceRouting(0, 1) = "CAL"
                        Case "Maintenance"
                            pMaintenanceRouting(0, 1) = "MNT"
                        Case "Inspection"
                            pMaintenanceRouting(0, 1) = "INS"
                    End Select
                End Select
        Next
    End If
        
    MaintenanceRouting = pMaintenanceRouting
End Property



' ##Activity attribute - TYPE ATTRIBUTE
Public Property Let TypeAttribute(value As String)
    pOSPServiceOp = value
End Property
Public Property Get TypeAttribute() As String
    TypeAttribute = pTypeAttribute
End Property

' ##Activity attribute - CAUSE
Public Property Let Cause(value As String)

    Dim causeValidation As Boolean
    Dim validationResult As String
    causeValidation = False

    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "G"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = lookupColumn & "2"
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    Dim causeList() As String
    ReDim causeList(rowCount - 1)
    
    For i = 0 To rowCount Step 1
        causeList(i) = wks.Range(lookupListlookupColumnBaseCell).Offset(i)
        If (value = causeList(i)) Then causeValidation = True
    Next
    
    If (causeValidation) Then
        pCause = value
    Else
        validationResult = "Activity Cause [ " & value & " ] is not valid. Use one of the following:" & vbNewLine
        For i = 0 To rowCount Step 1
            validationResult = validationResult & "-" & causeList(i) & vbNewLine
        Next
        pCause = InputBox(Prompt:=validationResult, Title:="Activity Cause Error")
    End If

    pCause = value
End Property
Public Property Get Cause() As String
    
Select Case (pActivityType)
    Case "Calibration", "Inspection"
        pCause = "Quality and Regulatory"
    Case "Maintenance"
        pCause = "Equipment"
    Case Else
        pCause = "Not Defined"
End Select

Cause = pCause

End Property

' ##Activity attribute - SOURCE
Public Property Let Source(value As String)

    Dim sourceValidation As Boolean
    Dim validationResult As String
    sourceValidation = False

    Dim wks As Worksheet
    Set wks = ThisWorkbook.Worksheets("Lookups")

    Dim lookupColumn As String
    Dim lookupColumnBaseCell As String
    Dim lookupColumnRange As String
    lookupColumn = "K"
    lookupColumnRange = lookupColumn & ":" & lookupColumn
    lookupListlookupColumnBaseCell = lookupColumn & "2"
    Dim rowCount As Integer
    rowCount = WorksheetFunction.CountA(wks.Range(lookupColumnRange)) - 1
    Dim sourceList() As String
    ReDim sourceList(rowCount - 1)
    
    For i = 0 To rowCount Step 1
        sourceList(i) = wks.Range(lookupListlookupColumnBaseCell).Offset(i)
        If (value = sourceList(i)) Then sourceValidation = True
    Next
    
    If (sourceValidation) Then
        pSource = value
    Else
        validationResult = "Activity Source [ " & value & " ] is not valid. Use one of the following:" & vbNewLine
        For i = 0 To rowCount Step 1
            validationResult = validationResult & "-" & sourceList(i) & vbNewLine
        Next
        pSource = InputBox(Prompt:=validationResult, Title:="Activity Source Error")
    End If

End Property
Public Property Get Source() As String

Select Case (pActivityType)
    Case "Calibration"
        pSource = "Calibration"
    Case "Maintenance"
        pSource = "Equipment"
    Case "Inspection"
        pSource = "Quality & Regulatory"
    Case Else
        pSource = "Not Defined"
End Select
    Source = pSource

End Property

' ##Activity attribute - SHUTDOWN
Public Property Let ShutdownRequired(value As Boolean)
If (value) Then
    pShutdownRequired = True
Else
    pShutdownRequired = False
End If
End Property
Public Property Get ShutdownRequired() As Boolean
If (pShutdownRequired) Then
    pShutdownRequired = True
Else
    pShutdownRequired = False
End If
ShutdownRequired = pShutdownRequired
End Property

' ## EAM READONLY PROPERTIES: Item Template
Public Property Get Template() As String
    pTemplate = "1:eAM Activities"
    Template = pTemplate
End Property

' ## EAM READONLY PROPERTIES: Item Category Set
Public Property Get InventoryCategorySet() As String
    pInventoryCategorySet = "EAM.ACTIVITY.EAM"
    InventoryCategorySet = pInventoryCategorySet
End Property

' ## EAM READONLY PROPERTIES: Planned
Public Property Get Planned() As String
    
If (Left(pActivityCode, 3) = "MA-") Then
    pPlanned = False
Else
    pPlanned = True
End If
    
Planned = pPlanned

End Property

Public Property Get AssetGroup() As String
    If (pActivityType <> "Not Defined") Then
        pAssetGroup = Mid(pActivityCode, WorksheetFunction.Search("-", pActivityCode) + 1, Len(pActivityCode))
    Else
        pAssetGroup = "Not Defined"
    End If
    AssetGroup = pAssetGroup

End Property


' ## PUBLIC CLASS FUNCTIONS
Function initialiseActivityWithActivityCode(code As String) As CActivity
    If (code = "<new activity>") Then
        theResult = InputBox("Enter an Activity Code:", "Create a new Activity: Enter new Code")
        code = theResult
    End If
    Me.activityCode = code
End Function

Private Sub Class_Initialize()
    pShutdownRequired = True
    pOSP = False
    pPlanned = True
End Sub
