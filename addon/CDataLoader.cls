VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### PROPERTY DECLARATIONS FOR PUBLIC ACCESS VIA GETTERS & SETTERS ###

' ### PRIVATE DECLARATIONS FOR INTERNAL CLASS USE ###

' An array of all the dataload strings / csv files??
Private pDataloadAgenda() As Variant
Private pDataloadScripts(6) As Variant


Public Property Get Agenda() As Variant
    Agenda = pDataloadAgenda
End Property
Public Function changeAgendaScriptName(item As Integer, value As String)
    pDataloadAgenda(item, 0) = value
End Function

Public Function dlScriptFromScriptName(ScriptName As String) As Variant
    Dim scriptArray() As Variant
    For i = 0 To UBound(pDataloadAgenda, 1)
        If (pDataloadAgenda(i, 0) = ScriptName) Then
            scriptArray = pDataloadAgenda(i, 1)
            Exit For
        End If
    Next
    dlScriptFromScriptName = scriptArray
End Function

'### AGENDA MANAGEMENT
Private Sub Class_Initialize()
    'columns:
    '0: scriptName,
    '1: script DL array()
    
    ReDim pDataloadAgenda(0, 0)
    
'    'List of script arrays available from this class
'    pDataloadScripts(0) = "Asset Group Definition"
'    pDataloadScripts(1) = "Activity Definition"
'    pDataloadScripts(2) = "Asset Group Activity Association (New Asset Group or Activity Only)"
'    pDataloadScripts(3) = "Asset Manual Activity Association (New Asset Group Only)"
'    pDataloadScripts(4) = "Attribute Assignment (New Asset Group Only)"
'    pDataloadScripts(5) = "Schedule Template Definition (New Activity Only)"
'    pDataloadScripts(6) = "Asset Instance Definition"
'    pDataloadScripts(7) = "Asset Instance File Attachment"
'    pDataloadScripts(8) = "Asset Schedule Update (New Asset Only)"

    'List of script arrays available from this class
    pDataloadScripts(0) = "Asset Group Definition"
    pDataloadScripts(1) = "Asset Activities Definition"
    pDataloadScripts(2) = "Asset Manual Activity Association"
    pDataloadScripts(3) = "Asset Instance Definition"
    pDataloadScripts(4) = "Asset Instance File Attachment"
    pDataloadScripts(5) = "Asset Schedule Update"
    pDataloadScripts(6) = "Asset Obsolete"
    
End Sub

Public Property Get DataloadScripts() As Variant
    DataloadScripts = pDataloadScripts
End Property

Public Function agendaForListBox()
    Dim arrayForListBox() As Variant
    
    If (pDataloadAgenda(0, 0) = "") Then
        ReDim arrayForListBox(0)
        arrayForListBox(0) = ""
    Else
        ReDim arrayForListBox(UBound(pDataloadAgenda, 1))
        For i = 0 To UBound(pDataloadAgenda, 1)
            arrayForListBox(i) = pDataloadAgenda(i, 0)
        Next
    End If
    agendaForListBox = arrayForListBox
End Function

Public Sub printAgenda(Optional additionalMessage As String)
    Dim message As String
    message = additionalMessage
    message = message & vbNewLine & vbNewLine & _
        "Your Dataload Agenda currently has the following scripts:"
    
    If (Not (IsEmpty(pDataloadAgenda))) Then
        message = message & vbNewLine
        For i = 0 To UBound(pDataloadAgenda, 1)
            message = message & vbNewLine & "-" & pDataloadAgenda(i, 0)
        Next
    End If
    
    'Message to user
'''''    MsgBox (message)
End Sub

Public Function addScriptToAgenda(ScriptName As String, script() As Variant) As Boolean
    
    If (IsEmpty(pDataloadAgenda(0, 0)) Or pDataloadAgenda(0, 0) = "") Then
        'No data entered yet
        ReDim pDataloadAgenda(0, 1)
        pDataloadAgenda(i, 0) = ScriptName
        pDataloadAgenda(i, 1) = script
    Else
        'Create new array to copy existing agenda to, but with 1 additional element for the new script
        'Dim newArray() As Variant
        c = UBound(pDataloadAgenda, 1) + 1
        ReDim newArray(c, 1)
        For i = 0 To UBound(pDataloadAgenda, 1)
            newArray(i, 0) = pDataloadAgenda(i, 0)
            newArray(i, 1) = pDataloadAgenda(i, 1)
        Next
        'load new script into last element
        newArray(UBound(newArray, 1), 0) = ScriptName
        newArray(UBound(newArray, 1), 1) = script
        pDataloadAgenda = newArray
    End If
           
    'Message to user
    printAgenda ("Script '" & ScriptName & "' added to Agenda.")
    addScriptToAgenda = True
End Function

Public Sub removeScriptFromAgenda(ScriptName As String)
    Dim found As Boolean
    found = False
    
    For i = 0 To UBound(pDataloadAgenda, 1)
        If (pDataloadAgenda(i, 0) = ScriptName) Then
            'script to delete has been found at index i
            found = True
            Dim newArray() As Variant
            ' If last array element, don't redim, just reset the values
            If (UBound(pDataloadAgenda, 1) - 1 < 0) Then
                pDataloadAgenda(0, 0) = ""
                pDataloadAgenda(0, 1) = ""
            Else
                ReDim newArray((UBound(pDataloadAgenda, 1) - 1), 1)
                Dim c As Integer
                c = 0
                For j = 0 To UBound(newArray, 1)
                    If (j = i) Then
                        j = j + 1
                    End If
                    newArray(c, 0) = pDataloadAgenda(j, 0)
                    newArray(c, 1) = pDataloadAgenda(j, 1)
                    c = c + 1
                Next
                pDataloadAgenda = newArray
            End If
            Exit For
        End If
    Next
    
    'Message to user
    If (found) Then
        printAgenda ("Script '" & ScriptName & "' was removed from Agenda")
    Else
        printAgenda ("Script '" & ScriptName & "' could not be found in your Agenda")
    End If
        
End Sub

Public Function resetAgenda() As Boolean
    Result = MsgBox("Are you sure you want to remove ALL scripts from your Agenda?", vbYesNo, "Reset Agenda")
    If (Result = vbYes) Then
        ReDim pDataloadAgenda(0, 0)
        printAgenda ("All scripts removed from Agenda")
    Else
        printAgenda ("Cancelled reset of Agenda")
    End If
    
    If (Result = vbYes) Then
        resetAgenda = True
    Else
        resetAgenda = False
    End If
End Function



'### Instance functions for building Dataloder arrays
'### DEFINE AN ASSET GROUP: MASTER ITEM DEFINITION

Public Function dlScriptDefineAssetGroup(AssetGroup As CAsset) As Variant
    Dim dlArray(45) As Variant
    
    'Item name - Asset Group Number
    dlArray(0) = AssetGroup.AssetGroupNumber
    dlArray(1) = "TAB"
    'Asset Group Description
    dlArray(2) = AssetGroup.Name
    dlArray(3) = "*AT"
    dlArray(4) = "*UP"
    dlArray(5) = "*UP"
    dlArray(6) = "ENT"
    'Template to copy from
    If (AssetGroup.AssetType = "Capital") Then
        dlArray(7) = "1:eAM Asset Groups"
    Else
        dlArray(7) = "1:eAM Rebuildable Items"
    End If
    dlArray(8) = "TAB"
    dlArray(9) = "*AD"
    dlArray(10) = "*SL(0.3)"
    dlArray(11) = "TAB"
    dlArray(12) = "TAB"
    dlArray(13) = "TAB"
    dlArray(14) = "TAB"
    dlArray(15) = "TAB"
    dlArray(16) = "TAB"
    dlArray(17) = "*AO"
    dlArray(18) = "*UP"
    dlArray(19) = "*SAVE"
    dlArray(20) = "*AO"
    dlArray(21) = ""
    'Asset Group Prefix for new assets
    dlArray(22) = AssetGroup.AssetGroupNumber & "/"
    dlArray(23) = "TAB"
    'Next instance number
    dlArray(24) = AssetGroup.NextInstanceNumber
    dlArray(25) = "TAB"
    dlArray(26) = "*SAVE"
    dlArray(27) = "*AT"
    dlArray(28) = "*DN"
    dlArray(29) = "*UP"
    dlArray(30) = "ENT"
    dlArray(31) = "TAB"
    'Category
    dlArray(32) = "EAM.ASSET_GROUP.EAM"
    dlArray(33) = "*DN"
    'Category
    dlArray(34) = ""
    dlArray(35) = "*CW"
    dlArray(36) = "*SAVE"
    dlArray(37) = "*AT"
    dlArray(38) = "*UP"
    dlArray(39) = "*UP"
    dlArray(40) = "*UP"
    dlArray(41) = "*UP"
    dlArray(42) = "ENT"
    dlArray(43) = "TAB"
    dlArray(44) = "*SB"
    dlArray(45) = "*SAVE"

    dlScriptDefineAssetGroup = dlArray
End Function


Public Function dlScriptDefineActivity(Activities As Variant) As Variant

Dim dlArray() As Variant
ReDim dlArray(UBound(Activities, 1), 98)

For i = 0 To UBound(Activities, 1)
    Dim Activity As CActivity
    Set Activity = New CActivity
    Activity.initialiseActivityWithActivityCode (Activities(i))
    
    dlArray(i, 0) = Activity.activityCode
    dlArray(i, 1) = "TAB"
    dlArray(i, 2) = Activity.ActivityDescription  'Activity.ActivityType & " - " & Activity.ServiceInterval & " Monthly - " & Activity.AssetGroup
    dlArray(i, 3) = "*AT"
    dlArray(i, 4) = "*UP"
    dlArray(i, 5) = "*UP"
    dlArray(i, 6) = "ENT"
    dlArray(i, 7) = "1:eAM Activities"
    dlArray(i, 8) = "TAB"
    dlArray(i, 9) = "*AD"
    dlArray(i, 10) = "*SL(0.3)"
    dlArray(i, 11) = "TAB"
    dlArray(i, 12) = "TAB"
    dlArray(i, 13) = "TAB"
    dlArray(i, 14) = "TAB"
    dlArray(i, 15) = "TAB"
    dlArray(i, 16) = "TAB"
    dlArray(i, 17) = "*AO"
    dlArray(i, 18) = "*UP"
    dlArray(i, 19) = "*SAVE"
    dlArray(i, 20) = "*AV"
    dlArray(i, 21) = "*UP"
    dlArray(i, 22) = "*UP"
    dlArray(i, 23) = "ENT"
    dlArray(i, 24) = "TAB"
    dlArray(i, 25) = "TAB"
    dlArray(i, 26) = "*AO"
    dlArray(i, 27) = "TAB"
    dlArray(i, 28) = "TAB"
    dlArray(i, 29) = "TAB"
    dlArray(i, 30) = "TAB"
    dlArray(i, 31) = "TAB"
    dlArray(i, 32) = "TAB"
    dlArray(i, 33) = "TAB"
    dlArray(i, 34) = "TAB"
    dlArray(i, 35) = "TAB"
    dlArray(i, 36) = "TAB"
    dlArray(i, 37) = "TAB"
    dlArray(i, 38) = "TAB"
    dlArray(i, 39) = "TAB"
    dlArray(i, 40) = "TAB"
    dlArray(i, 41) = "TAB"
    dlArray(i, 42) = "TAB"
    dlArray(i, 43) = "TAB"
    dlArray(i, 44) = "TAB"
    dlArray(i, 45) = "TAB"
    dlArray(i, 46) = "TAB"
    dlArray(i, 47) = "TAB"
    dlArray(i, 48) = "TAB"
    dlArray(i, 49) = "TAB"
    dlArray(i, 50) = "TAB"
    dlArray(i, 51) = "TAB"
    dlArray(i, 52) = "TAB"
    dlArray(i, 53) = "TAB"
    dlArray(i, 54) = "TAB"
    dlArray(i, 55) = "TAB"
    dlArray(i, 56) = "TAB"
    dlArray(i, 57) = "TAB"
    dlArray(i, 58) = "TAB"
    dlArray(i, 59) = "TAB"
    dlArray(i, 60) = "TAB"
    dlArray(i, 61) = "TAB"
    dlArray(i, 62) = "TAB"
    dlArray(i, 63) = "TAB"
    dlArray(i, 64) = "TAB"
    dlArray(i, 65) = "TAB"
    dlArray(i, 66) = "TAB"
    dlArray(i, 67) = "TAB"
    dlArray(i, 68) = "TAB"
    dlArray(i, 69) = "TAB"
    dlArray(i, 70) = "TAB"
    dlArray(i, 71) = "TAB"
    'Type
    dlArray(i, 72) = Activity.ActivityType
    dlArray(i, 73) = "TAB"
    'Cause
    dlArray(i, 74) = Activity.Cause
    dlArray(i, 75) = "TAB"
    'Source
    dlArray(i, 76) = Activity.Source
    dlArray(i, 77) = "TAB"
    'Shutdown
    dlArray(i, 78) = Activity.ShutdownRequired
    dlArray(i, 79) = "*SAVE"
    dlArray(i, 80) = "*AT"
    dlArray(i, 81) = "*DN"
    dlArray(i, 82) = "*UP"
    dlArray(i, 83) = "ENT"
    dlArray(i, 84) = "TAB"
    dlArray(i, 85) = "EAM.ACTIVITY.EAM"
    dlArray(i, 86) = "*DN"
    dlArray(i, 87) = "Default.Default"
    dlArray(i, 88) = "*CW"
    dlArray(i, 89) = "*SAVE"
    dlArray(i, 90) = "*AT"
    dlArray(i, 91) = "*UP"
    dlArray(i, 92) = "*UP"
    dlArray(i, 93) = "*UP"
    dlArray(i, 94) = "*UP"
    dlArray(i, 95) = "ENT"
    dlArray(i, 96) = "TAB"
    dlArray(i, 97) = "*SB"
    dlArray(i, 98) = "*SAVE"
Next

    dlScriptDefineActivity = dlArray
End Function

''## For new activities, the activity must be associated with the asset group
''###### !! STILL NEED TO DETERMINE IF THE ASSET HAS NEW ACTIVITIES !!
'Public Function dlScriptActivityAssociationsForAssetGroup(AssetGroup As CAsset)
'    Dim dlArray(24) As Variant
'
'    dlArray(0) = "*QE"
'    dlArray(1) = Activity.activityCode
'    dlArray(2) = "*QR"
'    dlArray(3) = "TAB"
'    dlArray(4) = "*DN"
'    dlArray(5) = "TAB"
'    dlArray(6) = Activity.AssetGroup
'    dlArray(7) = "TAB"
'    dlArray(8) = "Medium"
'    dlArray(9) = "TAB"
'    dlArray(10) = "TAB"
'    dlArray(11) = "TAB"
'    dlArray(12) = "Equipment"
'    dlArray(13) = "TAB"
'    dlArray(14) = "Maintenance"
'    dlArray(15) = "TAB"
'    dlArray(16) = "TAB"
'    dlArray(17) = "ML"
'    dlArray(18) = "TAB"
'    dlArray(19) = "TAB"
'    dlArray(20) = "ML.MC.M"
'    dlArray(21) = "TAB"
'    dlArray(22) = "Maintenance"
'    dlArray(23) = "*PB"
'    dlArray(24) = "*SAVE"
'
'    dlScriptActivityAssociationsForAssetGroup = dlArray
'End Function

'## MANUAL ACTIVITY ASSOCIATIONS FOR AN ASSET GROUP
Public Function dlScriptManualActivityAssociationsForAssetGroup(AssetGroup As CAsset) As Variant
    Dim dlArray(9, 27) As Variant
    For i = 0 To 9
    'Load the standard data into each column
        If (i = 0) Then
            dlArray(i, 0) = "*QE"
            dlArray(i, 1) = "MA-%"
            dlArray(i, 2) = "*QR"
        Else
            dlArray(i, 0) = ""
            dlArray(i, 1) = ""
            dlArray(i, 2) = ""
        End If
        dlArray(i, 3) = "TAB"
        dlArray(i, 4) = "TAB"
        dlArray(i, 5) = "*IR"
        
        'Check for Rebuild
        If (AssetGroup.AssetType = "Rebuildable") Then
            dlArray(i, 6) = "*DN"
        Else
            dlArray(i, 6) = ""
        End If
        
        dlArray(i, 7) = "TAB"
        dlArray(i, 8) = AssetGroup.AssetGroupNumber
        dlArray(i, 9) = "TAB"
        dlArray(i, 10) = "Medium"
        dlArray(i, 11) = "TAB"
        dlArray(i, 12) = "TAB"
        dlArray(i, 13) = "TAB"
            
        dlArray(i, 15) = "TAB"
        
        dlArray(i, 17) = "TAB"
        dlArray(i, 18) = "TAB"
        dlArray(i, 19) = AssetGroup.Department
        dlArray(i, 20) = "TAB"
        dlArray(i, 21) = "TAB"
        
        dlArray(i, 23) = "TAB"
        dlArray(i, 24) = ""
        dlArray(i, 25) = "*PB"
        dlArray(i, 26) = "*SAVE"
        dlArray(i, 27) = "*DN"
        Select Case i
            Case 0
                dlArray(i, 14) = "Quality and Regulatory"
                dlArray(i, 16) = "Calibration"
                dlArray(i, 22) = AssetGroup.Owner & ".C"
            Case 1
                dlArray(i, 14) = "Equipment"
                dlArray(i, 16) = "Maintenance"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 2
                dlArray(i, 14) = ""
                dlArray(i, 16) = "Modification"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 3
                dlArray(i, 14) = "Equipment"
                dlArray(i, 16) = "Obsolete"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 4
                dlArray(i, 14) = "Quality and Regulatory"
                dlArray(i, 16) = "Quarantine"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 5
                dlArray(i, 14) = "Equipment"
                dlArray(i, 16) = "Rebuild"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 6
                dlArray(i, 14) = "Engineering"
                dlArray(i, 16) = "Relocation"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 7
                dlArray(i, 14) = "Equipment"
                dlArray(i, 16) = "Repair"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 8
                dlArray(i, 14) = ""
                dlArray(i, 16) = "Replace"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
            Case 9
                dlArray(i, 14) = "Quality and Regulatory"
                dlArray(i, 16) = "Inspection"
                dlArray(i, 22) = AssetGroup.Owner & ".M"
        End Select
     Next

    dlScriptManualActivityAssociationsForAssetGroup = dlArray
End Function

Public Function dlScriptAssetInstanceDefinition(assetInstance As CAssetInstance)
    Dim dlArray(87) As Variant
    
    dlArray(0) = assetInstance.AssetGroup.AssetGroupNumber
    dlArray(1) = "TAB"
    dlArray(2) = assetInstance.AssetGroup.AssetGroupNumber & "/" & assetInstance.InstanceNumber
    dlArray(3) = "TAB"
    dlArray(4) = assetInstance.AssetGroup.Name
    dlArray(5) = "TAB"
    If (assetInstance.AssetGroup.AssetType = "Capital") Then
        dlArray(6) = "%"
        dlArray(7) = "TAB"
    Else
        dlArray(6) = ""
        dlArray(7) = ""
    End If
    dlArray(8) = assetInstance.Department
    dlArray(9) = "TAB"
    dlArray(10) = assetInstance.Criticality
    dlArray(11) = "TAB"
    dlArray(12) = assetInstance.Owner & ".M"
    dlArray(13) = "TAB"
    dlArray(14) = "*SAVE"
    dlArray(15) = "*SL(2)"
    If (assetInstance.AssetGroup.AssetType = "Rebuildable") Then
        dlArray(16) = "TAB"
        dlArray(17) = "31.REBUILD"
        dlArray(18) = "TAB"
        dlArray(19) = "31LC10AA01"
        dlArray(20) = "TAB"
        dlArray(21) = "*SAVE"
        dlArray(22) = "*AN"
        dlArray(23) = "*SL(0.5)"
        dlArray(24) = "*SAVE"
        dlArray(25) = "TAB"
        dlArray(26) = "TAB"
        dlArray(27) = "TAB"
    Else
        dlArray(16) = ""
        dlArray(17) = ""
        dlArray(18) = ""
        dlArray(19) = ""
        dlArray(20) = ""
        dlArray(21) = ""
        dlArray(22) = ""
        dlArray(23) = ""
        dlArray(24) = ""
        dlArray(25) = ""
        dlArray(26) = ""
        dlArray(27) = ""
    End If
    dlArray(28) = "TAB"
    dlArray(29) = "TAB"
    dlArray(30) = assetInstance.ParentAssetInstance
    dlArray(31) = "TAB"
    dlArray(32) = "*SAVE"
    dlArray(33) = "*SL(0.5)"
    dlArray(34) = "*SP"
    
    dlArray(35) = ""
    'ATTRIBUTES
    dlArray(36) = "*QE"
    dlArray(37) = "TAB"
    dlArray(38) = assetInstance.AssetGroup.AssetGroupNumber & "/" & assetInstance.InstanceNumber
    dlArray(39) = "*QR"
    dlArray(40) = "*AA"
    dlArray(41) = "Default Attributes"
    dlArray(42) = "TAB"
    dlArray(43) = "TAB"
    dlArray(44) = assetInstance.OperationalStatus
    dlArray(45) = "TAB"
    dlArray(46) = assetInstance.QualificationStatus
    dlArray(47) = "TAB"
    dlArray(48) = assetInstance.Owner
    dlArray(49) = "TAB"
    dlArray(50) = assetInstance.Make
    dlArray(51) = "TAB"
    dlArray(52) = assetInstance.Model
    dlArray(53) = "TAB"
    dlArray(54) = assetInstance.Serial
    dlArray(55) = "TAB"
    dlArray(56) = assetInstance.Contract
    dlArray(57) = "TAB"
    dlArray(58) = assetInstance.WarrantyEndDate
    dlArray(59) = "TAB"
    dlArray(60) = "*AO"
    dlArray(61) = "*SAVE"
    dlArray(62) = "*SL(0.5)"
    dlArray(63) = "*SP"
    
    dlArray(64) = ""
    'ATTACHEMENTS
    dlArray(65) = "*QE"
    dlArray(66) = "TAB"
    dlArray(67) = assetInstance.AssetGroup.AssetGroupNumber & "/" & assetInstance.InstanceNumber
    dlArray(68) = "*QR"
    dlArray(69) = "*AV"
    dlArray(70) = "*UP"
    dlArray(71) = "*UP"
    dlArray(72) = "*UP"
    dlArray(73) = "ENT"

    'Tool Record URL
    dlArray(76) = "10"
    dlArray(77) = "TAB"
    dlArray(78) = "Asset Attachments"
    dlArray(79) = "TAB"
    dlArray(80) = "Tool / Equipment Specification " & assetInstance.AssetGroup.DocNumber & ": " & assetInstance.AssetGroup.Name
    dlArray(81) = "TAB"
    dlArray(82) = "Web Page"
    dlArray(83) = "TAB"
    dlArray(84) = assetInstance.AssetGroup.hyperlinkToSpecificationDoc
    dlArray(85) = "*SAVE"
    dlArray(86) = "*CW"
    dlArray(87) = "*CW"


    
    dlScriptAssetInstanceDefinition = dlArray
End Function

Public Function dlScriptAssetInstanceAttachment(assetInstance As CAssetInstance)
    Dim dlArray(19) As Variant
    
    dlArray(0) = "*QE"
    dlArray(1) = "TAB"
    dlArray(2) = assetInstance.AssetGroup.AssetGroupNumber & "/" & assetInstance.InstanceNumber
    dlArray(3) = "*QR"
    dlArray(4) = "*AV"
    dlArray(5) = "*UP"
    dlArray(6) = "*UP"
    dlArray(7) = "*UP"
    dlArray(8) = "ENT"
    dlArray(9) = "*DN"
    dlArray(10) = ""
    dlArray(11) = "20"
    dlArray(12) = "TAB"
    dlArray(13) = "Asset Attachments"
    dlArray(14) = "TAB"
    dlArray(15) = "Initial Service Report: " & assetInstance.AssetGroup.Name & ", " & assetInstance.AssetGroup.AssetGroupNumber & "-" & assetInstance.InstanceNumber
    dlArray(16) = "TAB"
    dlArray(17) = "File"
    dlArray(18) = "TAB"
    dlArray(19) = "Initial Service Report: " & assetInstance.AssetGroup.Name & ", " & assetInstance.AssetGroup.AssetGroupNumber & "-" & assetInstance.InstanceNumber & ".pdf"
    
    dlScriptAssetInstanceAttachment = dlArray
End Function


Public Function dlScriptAssetInstanceObsolete(assetInstance As CAssetInstance)
    Dim dlArray(9) As Variant
    
    dlArray(0) = "*QE"
    dlArray(1) = "TAB"
    dlArray(2) = assetInstance.AssetGroup.AssetGroupNumber & "/" & assetInstance.InstanceNumber
    dlArray(3) = "*QR"
    dlArray(4) = "*AT"
    dlArray(5) = "*DN"
    dlArray(6) = "ENT"
    dlArray(7) = "*AO"
    dlArray(8) = "*AY"
    dlArray(9) = "*SAVE"

    dlScriptAssetInstanceObsolete = dlArray
End Function

Public Function dlScriptActivityAssociationsForAssetInstance(assetInstance As CAssetInstance)
    dlScriptActivityAssociationsForAssetInstance = ""
End Function

Public Function dlScriptDefineSchedule(assetInstance As CAssetInstance) As Variant
    Dim dlArray() As Variant
    ReDim dlArray(UBound(assetInstance.Schedules, 1), 42)
    
    For i = 0 To UBound(dlArray, 1)
        If (assetInstance.AssetGroup.AssetType = "Capital") Then
            dlArray(i, 0) = ""
        Else
            dlArray(i, 0) = "*DN"
        End If
        dlArray(i, 1) = "TAB"
        dlArray(i, 2) = "TAB"
        dlArray(i, 3) = "TAB"
        dlArray(i, 4) = assetInstance.AssetGroup.AssetGroupNumber & "/" & assetInstance.InstanceNumber
        dlArray(i, 5) = "TAB"
        dlArray(i, 6) = assetInstance.AssetGroup.Activities(i) 'AssetInstance.Schedules(0).activityCode
        dlArray(i, 7) = "*AI"
        dlArray(i, 8) = "*SL(0.5)"
        dlArray(i, 9) = assetInstance.Schedules(i).ScheduleCode
        dlArray(i, 10) = "TAB"
        dlArray(i, 11) = "TAB"
        dlArray(i, 12) = "TAB"
        dlArray(i, 13) = assetInstance.Schedules(i).EffectiveFrom
        dlArray(i, 14) = "TAB"
        dlArray(i, 15) = assetInstance.Schedules(i).EffectiveTo
        dlArray(i, 16) = "TAB"
        dlArray(i, 17) = "TAB"
        dlArray(i, 18) = "TAB"
        dlArray(i, 19) = "TAB"
        dlArray(i, 20) = "TAB"
        dlArray(i, 21) = "TAB"
        dlArray(i, 22) = "TAB"
        dlArray(i, 23) = "TAB"
        dlArray(i, 24) = "TAB"
        dlArray(i, 25) = "*SB"
        dlArray(i, 26) = "*NB"
        dlArray(i, 27) = "TAB"
        dlArray(i, 28) = "TAB"
        dlArray(i, 29) = assetInstance.Schedules(i).ServiceInterval
        dlArray(i, 30) = "*AA"
        dlArray(i, 31) = assetInstance.Schedules(i).LastServiceDateActualStart
        dlArray(i, 32) = "TAB"
        dlArray(i, 33) = assetInstance.Schedules(i).LastServiceDateActualEnd
        dlArray(i, 34) = "TAB"
        dlArray(i, 35) = assetInstance.Schedules(i).LastServiceDateScheduledStart
        dlArray(i, 36) = "TAB"
        dlArray(i, 37) = assetInstance.Schedules(i).LastServiceDateScheduledEnd
        dlArray(i, 38) = "TAB"
        dlArray(i, 39) = "*CW"
        dlArray(i, 40) = "*SAVE"
        dlArray(i, 41) = "*CW"
        dlArray(i, 42) = "*CW"
    Next
    
    dlScriptDefineSchedule = dlArray
End Function


' ### Other Functions ###

' #Run Dataloader Program opening the template file
Public Function runDataLoader()
    runDataLoader = Shell("C:\Program Files\DataLoad\DataLoad.exe " & _
        "C:\Documents and Settings\TDoyle\My Documents\eAM Production Data\ACRF\DL.dld", _
        vbMaximizedFocus)
End Function

Public Sub exportScriptToCSV(ScriptName As String)
 
    scriptArray = dlScriptFromScriptName(ScriptName)
    numberOfDims = Main.getNumberOfDimensionsInArray(scriptArray)
    
    filepath = InputBox("Export to filename:", "Export Script", "c:\temp\dl\" & ScriptName & ".tsv")
        
    Open filepath For Output As #1
    For i = 0 To UBound(scriptArray, 1)
        If (numberOfDims > 1) Then
            For j = 0 To (UBound(scriptArray, numberOfDims))
                rowText = rowText & scriptArray(i, j)
                'Don't add a tab at the end of the row
                If (j <= UBound(scriptArray, numberOfDims)) Then
                    rowText = rowText & vbTab
                End If
            Next
            Print #1, rowText
            rowText = ""
        Else
            dim1rowText = dim1rowText & scriptArray(i)
            If (i <> UBound(scriptArray, 1)) Then
                dim1rowText = dim1rowText & vbTab
            End If
        End If
    Next
    
    If (numberOfDims = 1) Then
        Print #1, dim1rowText
    End If
    
    Close #1
    Result = MsgBox("Script exported to file: " & filepath, vbInformation, "Export Script")
End Sub


Public Sub exportScriptsArrayToCSV(exportArray As Variant)
 
    Dim filepath() As Variant
    Dim numberOfScripts As Integer
    
    numberOfScripts = UBound(exportArray, 1)
    ReDim filepath(numberOfScripts)
    
    
    For i = 0 To numberOfScripts
        Dim theScriptName As String
        theScriptName = exportArray(i)
        
        scriptArray = Me.dlScriptFromScriptName(theScriptName)
        numberOfDims = Main.getNumberOfDimensionsInArray(scriptArray)
        
        ReDim Preserve filepath(i)
        filepath(i) = "c:\temp\dl\" & exportArray(i) & ".tsv"
            
        Open filepath(i) For Output As #1
        For f = 0 To UBound(scriptArray, 1)
            If (numberOfDims > 1) Then
                For j = 0 To (UBound(scriptArray, numberOfDims))
                    rowText = rowText & scriptArray(f, j)
                    'Don't add a tab at the end of the row
                    If (j <= UBound(scriptArray, numberOfDims)) Then
                        rowText = rowText & vbTab
                    End If
                Next
                Print #1, rowText
                rowText = ""
            Else
                dim1rowText = dim1rowText & scriptArray(f)
                If (f <> UBound(scriptArray, 1)) Then
                    dim1rowText = dim1rowText & vbTab
                End If
            End If
        Next
        
        If (numberOfDims = 1) Then
            Print #1, dim1rowText
        End If
        
        Close #1
    Next
    
    Dim filesSavedList As String
    For k = 0 To UBound(filepath, 1)
        If (k = numberOfScripts) Then
            filesSavedList = filesSavedList & filepath(k)
        Else
            filesSavedList = filesSavedList & filepath(k) & vbNewLine
        End If
    Next
    Result = MsgBox("Scripts exported to files: " & vbNewLine & vbNewLine & filesSavedList, vbInformation, "Export Script")

End Sub


'Public Function getNumberOfDimensionsInArray(theArray As Variant) As Integer
'    On Error GoTo resolveDimension
'    For dimensions = 1 To 600000
'        errorCheck = LBound(theArray, dimensions)
'    Next
'resolveDimension:
'    getNumberOfDimensionsInArray = dimensions - 1
'End Function
