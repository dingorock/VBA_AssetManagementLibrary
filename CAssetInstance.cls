VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAssetInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### INTERNAL CLASS INSTANCE VARIABLE DECLARATIONS ###

' ### PROPERTY DECLARATIONS FOR PUBLIC ACCESS VIA GETTERS & SETTERS ###
Private pAssetGroup As CAsset
Private pInstanceNumber As Integer
Private pCriticality As String
Private pParentAssetInstance As String
Private pIsSoftware As Boolean

Private pSchedules() As Variant

Private pOperationalStatus As String
Private pQualificationStatus As String
Private pDepartment As String
Private pOwner As String
Private pMake As String
Private pModel As String
Private pSerial As String
Private pContract As String
Private pWarrantyEndDate As String

Private newAssetGroup As Boolean

' eAM STATIC READONLY VARIABLES
' ### PROPERTY SETTERS & GETTERS

' ## ASSET GROUP
Public Property Let AssetGroup(value As CAsset)
    Set pAssetGroup = value
End Property
Public Property Get AssetGroup() As CAsset
    Set AssetGroup = pAssetGroup
End Property

' ## INSTANCE
Public Property Let InstanceNumber(value As Integer)
    pInstanceNumber = value
End Property
Public Property Get InstanceNumber() As Integer
    
    InstanceNumber = pInstanceNumber
End Property

' ## CRITICALITY
Public Property Let Criticality(value As String)
    pCriticality = value
End Property
Public Property Get Criticality() As String
    Criticality = pCriticality
End Property

' ## PARENT ASSET INSTANCE
Public Property Let ParentAssetInstance(value As String)
    pParentAssetInstance = value
End Property
Public Property Get ParentAssetInstance() As String
    ParentAssetInstance = pParentAssetInstance
End Property

' ## IS SOFTWARE?
Public Property Let IsSoftware(value As Boolean)
    pIsSoftware = value
End Property
Public Property Get IsSoftware() As Boolean
    If (Left(pAssetGroup.AssetGroupNumber, 1) = "K") Then
        pIsSoftware = True
    Else
        pIsSoftware = False
    End If
    IsSoftware = pIsSoftware
End Property

' ## ATTRIBUTE: OPERATIONAL STATUS
Public Property Let OperationalStatus(value As String)
    Dim operationalStatusValidation As Boolean
    Dim validationResult As String
    operationalStatusValidation = False
    Select Case (value)
        Case "Active", "Quarantined", "Obsolete", "Missing"
            operationalStatusValidation = True
            pOperationalStatus = value
        Case Else
            operationalStatusValidation = False
            pOperationalStatus = "Quarantined"
            validationResult = "Operational Status [ " & value & " ] is not valid." & vbNewLine & "-Set status to QUARANTINED"
            MsgBox Prompt:=validationResult, Title:="Qualification Status Error"
    End Select
End Property
Public Property Get OperationalStatus() As String
    OperationalStatus = pOperationalStatus
End Property

' ## ATTRIBUTE: QUALIFICATION STATUS
Public Property Let QualificationStatus(value As String)
    Dim qualificationStatusValidation As Boolean
    Dim validationResult As String
    qualificationStatusValidation = False
    Select Case (value)
        Case "Q"
            qualificationStatusValidation = True
            pQualificationStatus = "QUALIFIED"
        Case "NQ"
            qualificationStatusValidation = True
            pQualificationStatus = "NOT QUALIFIED"
        Case "NA"
            qualificationStatusValidation = True
            pQualificationStatus = "QUALIFICATION NOT REQUIRED"
        Case Else
            qualificationStatusValidation = False
            pQualificationStatus = "NOT QUALIFIED"
            validationResult = "Qualification Status [ " & value & " ] is not valid." & vbNewLine & "-Set status to NOT QUALIFIED"
            MsgBox Prompt:=validationResult, Title:="Qualification Status Error"
    End Select
End Property
Public Property Get QualificationStatus() As String
    QualificationStatus = pQualificationStatus
End Property

' ## ATTRIBUTE: OWNER
Public Property Let Owner(value As String)
    pOwner = value
End Property
Public Property Get Owner() As String
    Owner = pOwner
End Property

Public Property Get Department() As String
    pDepartment = Left(pOwner, WorksheetFunction.Search(".", pOwner) - 1)
    Department = pDepartment
End Property

' ## ATTRIBUTE: MAKE
Public Property Let Make(value As String)
    pMake = value
End Property
Public Property Get Make() As String
    Make = pMake
End Property

' ## ATTRIBUTE: MODEL
Public Property Let Model(value As String)
    pModel = value
End Property
Public Property Get Model() As String
    Model = pModel
End Property

' ## ATTRIBUTE: SERIAL NUMBER
Public Property Let Serial(value As String)
    pSerial = value
End Property
Public Property Get Serial() As String
    Serial = pSerial
End Property

' ## ATTRIBUTE: CONTRACT NUMBER / REFERENCE
Public Property Let Contract(value As String)
    pContract = value
End Property
Public Property Get Contract() As String
    Contract = pContract
End Property

' ## ATTRIBUTE: WARRANTY EXPIRY
Public Property Let WarrantyEndDate(value As String)
    pWarrantyEndDate = value
End Property
Public Property Get WarrantyEndDate() As String
    WarrantyEndDate = pWarrantyEndDate
End Property

Public Property Let Schedules(value As Variant)
    Schedules = value
End Property

Public Property Get Schedules() As Variant
If (Not (IsObject(pSchedules(0)))) Then
    activitiesCount = Me.AssetGroup.NumberOfActivities
    If (activitiesCount <> 0) Then
        'Used to build a schedule for each activity/instance combination provided
        'Then assign the schedule object into the pSchedules Array
        
        Dim tempSchedule As CSchedule
    
        'Create local temporary activity object
        Dim tempActivity As CActivity
    
        'Create local array of activities
        Dim activitiesObjects() As CActivity
        ReDim activitiesObjects(activitiesCount - 1)
        
        For i = 0 To (activitiesCount - 1)
            Set tempActivity = New CActivity
            tempActivity.initialiseActivityWithActivityCode (Me.AssetGroup.Activities(i))
            Set activitiesObjects(i) = tempActivity
        Next
    
        'Fill an array - pSchedules - with schedule objects
        ReDim pSchedules(Me.AssetGroup.NumberOfActivities - 1)
    
        For i = 0 To (Me.AssetGroup.NumberOfActivities - 1)
            Set tempActivity = New CActivity
            Set tempSchedule = New CSchedule
            Set pSchedules(i) = New CSchedule
            tempInstanceNumber = Me.InstanceNumber
            
            tempActivity.initialiseActivityWithActivityCode (Me.AssetGroup.Activities(i))
            tempSchedule.Activity = tempActivity
            tempSchedule.assetInstance = tempInstanceNumber
            
            Set pSchedules(i) = tempSchedule
        Next
    Else
        ReDim pSchedules(0)
        pSchedules(0) = ""
    End If
End If

Schedules = pSchedules

End Property


' ### Initialising Class Method
Private Sub Class_Initialize()
    pCriticality = "Normal"
    ReDim pSchedules(0)
    pSchedules(0) = ""
End Sub

Sub initialiseAssetInstanceFromRequest(request As Variant)

    'Create an asset group object for this requested instance
    Set pAssetGroup = New CAsset
    pAssetGroup.initialiseAssetGroupFromRequest (request)
    
    If (request(1, 7) = "") Then
        pParentAssetInstance = request(1, 6)
    Else
        pParentAssetInstance = request(1, 7)
    End If
    
    'Check if new of modified asset, for instance number
    If (request(1, 1) = "Create") Then
        pInstanceNumber = pAssetGroup.NextInstanceNumber
    Else
        If (IsNumeric(request(1, 5))) Then
            pInstanceNumber = request(1, 5)
        Else
            pInstanceNumber = 0
'            pInstanceNumber = CInt(Left(request(1, 5), (WorksheetFunction.Search(",", request(1, 5))) - 1))
'            MsgBox ("Instance number provided is not a valid number. Attempting to obtain the first number from the data.")
         End If
    End If
    
    Me.QualificationStatus = request(1, 10)
    Me.OperationalStatus = request(1, 11)
    pOwner = request(1, 19)
    pMake = request(1, 20)
    pModel = request(1, 21)
    pSerial = request(1, 22)
    pContract = request(1, 24)
    pWarrantyEndDate = request(1, 25)
    
    'ReDim pSchedules(0)
    Result = Me.Schedules
End Sub

'Sub initialiseAssetParentInstanceFromRequest(request As Variant)
'
'pParentAssetGroup.initialiseAssetGroupFromAssetGroupNumber (Left(request(1, 7), WorksheetFunction.Search("/", request(1, 7)) - 1))
'pParentAssetInstance.initialiseAssetParentInstanceFromRequest (request)
'pInstanceNumber = Right(request(1, 7), Len(request(1, 7)) - WorksheetFunction.Search("/", request(1, 7)))
'
''pCriticality = "Normal"
''pQualificationStatus = request(1, 10)
''pOperationalStatus = request(1, 11)
''pOwner = request(1, 19)
''pMake = request(1, 20)
''pModel = request(1, 21)
''pSerial = request(1, 22)
''pContract = request(1, 24)
''pWarrantyEndDate = request(1, 25)
'
'End Sub

Public Function schedulesArray() As Variant
    Dim anArray() As Variant
    ReDim anArray(Me.AssetGroup.NumberOfActivities)
    c = 0
    
    If (Not (IsObject(pSchedules(0)))) Then
        anArray(0) = ""
    Else
        For Each scheduleElement In pSchedules
            anArray(c) = scheduleElement.ScheduleCode
            c = c + 1
        Next
    End If
    schedulesArray = anArray
End Function

Public Function activitiesArray() As Variant
    Dim anArray() As Variant
    ReDim anArray(Me.AssetGroup.NumberOfActivities)
    
    If (AssetGroup.Activities(0) = "") Then
        anArray(0) = ""
    Else
        For c = 0 To UBound(AssetGroup.Activities, 1)
            Dim anActivity As CActivity
            Set anActivity = New CActivity
            anActivity.initialiseActivityWithActivityCode (AssetGroup.Activities(c))
            anArray(c) = anActivity
        Next
    End If
    activitiesArray = anArray
End Function


Public Function assetInstanceArray() As Variant

    Dim tempArray(1, 13) As Variant
    'Headings
    tempArray(0, 0) = "AssetGroup"
    tempArray(0, 1) = "AssetGroupName"
    tempArray(0, 2) = "InstanceNumber"
    tempArray(0, 3) = "Criticality"
    tempArray(0, 4) = "ParentAssetInstance"
    tempArray(0, 5) = "IsSoftware"
    tempArray(0, 6) = "OperationalStatus"
    tempArray(0, 7) = "QualificationStatus"
    tempArray(0, 8) = "Owner"
    tempArray(0, 9) = "Make"
    tempArray(0, 10) = "Model"
    tempArray(0, 11) = "Serial"
    tempArray(0, 12) = "Contract"
    tempArray(0, 13) = "WarrantyEndDate"
    
    'Data
    tempArray(1, 0) = AssetGroup.AssetGroupNumber
    tempArray(1, 1) = AssetGroup.Name
    tempArray(1, 2) = InstanceNumber
    tempArray(1, 3) = Criticality
    tempArray(1, 4) = ParentAssetInstance
    tempArray(1, 5) = IsSoftware
    tempArray(1, 6) = OperationalStatus
    tempArray(1, 7) = QualificationStatus
    tempArray(1, 8) = Owner
    tempArray(1, 9) = Make
    tempArray(1, 10) = Model
    tempArray(1, 11) = Serial
    tempArray(1, 12) = Contract
    tempArray(1, 13) = WarrantyEndDate
    
    assetInstanceArray = tempArray

End Function


Public Function updateInstanceFromAssetInstanceArray(i As Integer, newValue As Variant) As Boolean
newAssetGroup = False
'COULD USE THE FUNCTION ABOVE THAT PROVIDES AN ARRAY TO SEARCH & RETURN THE CORRECT INDEX

    Select Case i
'    0) = "AssetGroup"
'    1) = "InstanceNumber"
'    2) = "Criticality"
'    3) = "ParentAssetInstance"
'    4) = "IsSoftware"
'    5) = "OperationalStatus"
'    6) = "QualificationStatus"
'    7) = "Owner"
'    8) = "Make"
'    9) = "Model"
'    10) = "Serial"
'    11) = "Contract"
'    12) = "WarrantyEndDate"
    
    'Data
    Case 0
        Me.AssetGroup.AssetGroupNumber = newValue
    Case 1
        Me.AssetGroup.Name = newValue
        If (newValue = "" Or newValue = " ") Then newAssetGroup = True
    Case 2
        Me.InstanceNumber = newValue
    Case 3
        Me.Criticality = newValue
    Case 4
        Me.ParentAssetInstance = newValue
    Case 5
        Me.IsSoftware = newValue
    Case 6
        Me.OperationalStatus = newValue
    Case 7
        Me.QualificationStatus = newValue
    Case 8
        Me.Owner = newValue
    Case 9
        Me.Make = newValue
    Case 10
        Me.Model = newValue
    Case 11
        Me.Serial = newValue
    Case 12
        Me.Contract = newValue
    Case 13
        Me.WarrantyEndDate = newValue
    End Select
    
    updateInstanceFromAssetInstanceArray = True
End Function
