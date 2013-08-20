VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### INTERNAL CLASS INSTANCE VARIABLE DECLARATIONS ###

' ### PROPERTY DECLARATIONS FOR PUBLIC ACCESS VIA GETTERS & SETTERS ###
Private pScheduleCode As String
Private pAssetInstance As Integer
Private pActivity As CActivity
Private pEffectiveFrom As String
Private pEffectiveTo As String
Private pRunToFailure As Boolean
Private pInterval As Integer
Private pLastServiceDateActualStart As String
Private pLastServiceDateActualEnd As String
Private pLastServiceDateScheduledStart As String
Private pLastServiceDateScheduledEnd As String

Private pDefaultScheduleDate As String

' eAM STATIC READONLY VARIABLES
Private pSetName As String

' ### PROPERTY SETTERS & GETTERS
' ## SCHEDULE NAME / CODE
Private Function getDefaultScheduleDate() As String
    getDefaultScheduleDate = CStr(DateTime.Date)
End Function



Public Property Let ScheduleCode(value As String)
    pScheduleCode = value
End Property
Public Property Get ScheduleCode() As String
    pScheduleCode = Me.Activity.activityCode & "/" & Me.assetInstance
    ScheduleCode = pScheduleCode
End Property

' ## ASSET INSTANCE
Public Property Let assetInstance(value As Integer)
    pAssetInstance = value
End Property
Public Property Get assetInstance() As Integer
    assetInstance = pAssetInstance
End Property

' ## ACTIVITY
Public Property Let Activity(value As CActivity)
    Set pActivity = value
End Property
Public Property Get Activity() As CActivity
    Set Activity = pActivity
End Property

' ## EFFECTIVITY DATES ##
Public Property Let EffectiveFrom(value As String)
    pEffectiveFrom = value
End Property
Public Property Get EffectiveFrom() As String
    EffectiveFrom = pEffectiveFrom
End Property
Public Property Let EffectiveTo(value As String)
    pEffectiveTo = value
End Property
Public Property Get EffectiveTo() As String
    EffectiveTo = pEffectiveTo
End Property

' ## RUN TO FAILURE FLAG ##
Public Property Let RunToFailure(value As String)
    pRunToFailure = value
End Property
Public Property Get RunToFailure() As String
    RunToFailure = pRunToFailure
End Property

' ## SERVICE INTERVAL ##
Public Property Let ServiceInterval(value As Integer)
    pServiceInterval = value
End Property
Public Property Get ServiceInterval() As Integer
    ServiceInterval = pActivity.ServiceInterval
End Property

' ## LAST SERVICE DATES ##
' # ACTUAL
Public Property Let LastServiceDateActualEnd(value As String)
    pLastServiceDateActualEnd = value
End Property
Public Property Get LastServiceDateActualEnd() As String
    If (pLastServiceDateActualEnd = "") Then
        pLastServiceDateActualEnd = getDefaultScheduleDate
    End If
    LastServiceDateActualEnd = pLastServiceDateActualEnd
End Property
Public Property Let LastServiceDateActualStart(value As String)
    pLastServiceDateActualStart = value
End Property
Public Property Get LastServiceDateActualStart() As String
    If (pLastServiceDateActualStart = "") Then
        pLastServiceDateActualStart = getDefaultScheduleDate
    Else
        pLastServiceDateActualStart = pLastServiceDateActualStart
    End If
    LastServiceDateActualStart = pLastServiceDateActualStart
End Property

' ## LAST SERVICE DATES ##
' # SCHEDULED
Public Property Let LastServiceDateScheduledStart(value As String)
    pLastServiceDateScheduledStart = value
End Property
Public Property Get LastServiceDateScheduledStart() As String
    If (pLastServiceDateScheduledStart = "") Then
        pLastServiceDateScheduledStart = getDefaultScheduleDate
    Else
        pLastServiceDateActualStart = pLastServiceDateActualStart
    End If
    LastServiceDateScheduledStart = pLastServiceDateScheduledStart
End Property
Public Property Let LastServiceDateScheduledEnd(value As String)
    pLastServiceDateActualEnd = value
End Property
Public Property Get LastServiceDateScheduledEnd() As String
    If (pLastServiceDateScheduledEnd = "") Then
        pLastServiceDateScheduledEnd = getDefaultScheduleDate
    Else
        pLastServiceDateScheduledEnd = pLastServiceDateActualStart
    End If
    LastServiceDateScheduledEnd = pLastServiceDateScheduledEnd
End Property

' ##EAM STATIC PROPERTIES
Public Property Get SetName() As String
    pSetName = "MAIN"
    SetName = pSetName
End Property

Public Sub initialiseWithActivityAndAssetInstance(anActivity As CActivity, anAssetInstance As Integer)
    
    Me.assetInstance = anAssetInstance
    Set pActivity = anActivity
    Me.ScheduleCode = anActivity.activityCode & "/" & anAssetInstance

'    If (pAssetInstance.OperationalStatus = "Quarantined") Then
'        pEffectiveTo = DateTime.Now
'    End If
    
    pInterval = pActivity.ServiceInterval
End Sub

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

Public Function updateScheduleFromScheduleArray(scheduleObject As CSchedule, i As Integer, newValue As String) As Boolean
    
'COULD USE THE FUNCTION ABOVE THAT PROVIDES AN ARRAY TO SEARCH & RETURN THE CORRECT INDEX

    Select Case i
    '0) = "ScheduleCode"
    '1) = "AssetInstance"
    '2) = "Activity"
    '3) = "EffectiveFrom"
    '4) = "EffectiveTo"
    '5) = "RunToFailure"
    '6) = "Interval"
    '7) = "LastServiceDateActualStart"
    '8) = "LastServiceDateActualEnd"
    '9) = "LastServiceDateScheduledStart"
    '10) = "pLastServiceDateScheduledEnd"
    
    'Data
    Case 0
        scheduleObject.ScheduleCode = CInt(newValue)
    Case 1
        scheduleObject.assetInstance = CInt(newValue)
    Case 2
        scheduleObject.Activity.activityCode = CInt(newValue)
    Case 3
        scheduleObject.EffectiveFrom = CStr(newValue)
    Case 4
        scheduleObject.EffectiveTo = CInt(newValue)
    Case 5
        scheduleObject.RunToFailure = CBool(newValue)
    Case 6
        scheduleObject.ServiceInterval = CInt(newValue)
    Case 7
        scheduleObject.LastServiceDateActualStart = CStr(newValue)
    Case 8
        scheduleObject.LastServiceDateActualEnd = CStr(newValue)
    Case 9
        scheduleObject.LastServiceDateScheduledStart = CStr(newValue)
    Case 10
        scheduleObject.LastServiceDateScheduledEnd = CStr(newValue)
    End Select
    
    updateScheduleFromScheduleArray = True
End Function

