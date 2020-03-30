'This Action Script is NOT Device assigned and should be used as a global recurring action that could be scheduled or run manually as needed
Option Explicit
'***DO NOT MODIFY ABOVE THIS LINE AT RISK OF BREAKING THE SCRIPT**
' ******************* Configuration *************************
'Set to True if you want to *DISABLE* interfaces, Set to False if you do not want to *DISABLE* any interface active monitors
Dim bRunDisable : bRunDisable = True
'This should match a portion of the comment of the interface active monitors you'd like to *DISABLE*
Dim sDisableComment : sDisableComment = "GigabitEthernet"
Dim sDisableComment2 : sDisableComment2 = "" 'Optional: Enter second comment to match, leave blank if you don't need it
Dim sDisableComment3 : sDisableComment3 = "" 'Optional: Enter third comment to match, leave blank if you don't need it
Dim sDisableComment4 : sDisableComment4 = "" 'Optional: Enter fourth comment to match, leave blank if you don't need it

'Set to True if you want to *ENABLE* interfaces, Set to False if you do not want to *ENABLE* any interface active monitors
Dim bRunEnable : bRunEnable = True
'This should match a portion of the comment of the interface active monitors you'd like to *ENABLE*
Dim sEnableComment : sEnableComment = "Trunk" 'Enter first comment to match
Dim sEnableComment2 : sEnableComment2 = "VLAN" 'Optional: Enter second comment to match, leave blank if you don't need it
Dim sEnableComment3 : sEnableComment3 = "to switch" 'Optional: Enter third comment to match, leave blank if you don't need it
Dim sEnableComment4 : sEnableComment4 = "" 'Optional: Enter fourth comment to match, leave blank if you don't need it

'When disabling, 0 means disable all that match, 1 means disable only those that match and are down
Dim bDownDisable : bDownDisable = 0

'Device Group(s): Set to the device group you'd like to run the script against
'This can be either a device group or dynamic group name. It must match exactly!
'Leaving this variable blank is the equivalent of "All Devices"
Dim sDeviceGroup : sDeviceGroup = ""

'Do a dry run? If True, then SQL commands will not be executed, only logged.
Dim bDryRun : bDryRun = False
' ***************** End Configuration ***********************

'***DO NOT MODIFY BELOW THIS LINE AT RISK OF BREAKING THE SCRIPT***

'Variable declaration
Dim nPivotActiveMonitorTypeToDeviceID 'This will be used for tracking the nPivotActiveMonitorTypeToDeviceID
Dim sDisableQuery 'This will be used for settings the disable query depending on if you set bDisable to 1
Dim nCurrDevID : nCurrDevID = 0 'This will be used for tracking which nDeviceID we are working on
Dim nNewDevID 'This will be used for tracking which nDeviceID we are working on
Dim bContinue : bContinue = True 'This is a true/false flag to specifying whether to continue in the script or not
Dim sGroupQuery : sGroupQuery = "" 'This will be filled in later depending on what type of group is entered in the 'sDeviceGroup' variable
Dim nDisableCount : nDisableCount = 0 'This will be used to keep count of how many devices we disabled interfaces on
Dim nEnableCount : nEnableCount = 0 'This will be used to keep count of how many devices we enabled interfaces on
Dim sFinalEnableComment : sFinalEnableComment = "" 'This will set the final enable comment TSQL
Dim sFinalDisableComment : sFinalDisableComment = "" 'This will set the final disable comment TSQL

'Optional: TSQL adjustments: You can adjust the TSQL matching here by adjusting where % (wildcards) are placed
If Len(sEnableComment) > 0 Then
	sFinalEnableComment = "(sComment like '%" & sEnableComment & "%' "
End If
If Len(sEnableComment2) > 0 Then
	sFinalEnableComment = sFinalEnableComment & "or sComment like '%" & sEnableComment2 & "%' "
End If
If Len(sEnableComment3) > 0 Then
	sFinalEnableComment = sFinalEnableComment & "or sComment like '%" & sEnableComment3 & "%' "
End If
If Len(sEnableComment4) > 0 Then
	sFinalEnableComment = sFinalEnableComment & "or sComment like '%" & sEnableComment4 & "%' "
End If
sFinalEnableComment = sFinalEnableComment & ") "

If Len(sDisableComment) > 0 Then
	sFinalDisableComment = "(sComment like '%" & sDisableComment & "%' "
End If
If Len(sDisableComment2) > 0 Then
	sFinalDisableComment = sFinalDisableComment & "or sComment like '%" & sDisableComment2 & "%' "
End If
If Len(sDisableComment3) > 0 Then
	sFinalDisableComment = sFinalDisableComment & "or sComment like '%" & sDisableComment3 & "%' "
End If
If Len(sDisableComment4) > 0 Then
	sFinalDisableComment = sFinalDisableComment & "or sComment like '%" & sDisableComment4 & "%' "
End If
sFinalDisableComment = sFinalDisableComment & ") "

'Create DB Object
Dim oDB : Set oDB = Context.GetDB
'Create WUG Event Helper Object
Dim oEvent : Set oEvent = CreateObject("CoreAsp.EventHelper")

'****BEGIN MAIN SCRIPT****
'Get the device group type
GetDeviceGroupType sDeviceGroup
'If the continue flag is true, keep going
If bContinue = True Then
 Dim sSqlToEnable : sSqlToEnable = "SELECT nPivotActiveMonitorTypeToDeviceID, PAMTD2.nDeviceID FROM PivotActiveMonitorTypeToDevice PAMTD2" & _
 " join ActiveMonitorType on ActiveMonitorType.nActiveMonitorTypeID = PAMTD2.nActiveMonitorTypeID " & _
 " where (PAMTD2.nActiveMonitorTypeID = 22 and " & sFinalEnableComment & _
 " and bDisabled = 1) " & sGroupQuery
 Dim sSqlToDisable : sSqlToDisable = "SELECT nPivotActiveMonitorTypeToDeviceID, PAMTD2.nDeviceID FROM PivotActiveMonitorTypeToDevice PAMTD2" & _
 " join ActiveMonitorType on ActiveMonitorType.nActiveMonitorTypeID = PAMTD2.nActiveMonitorTypeID " & _
 " join Device D on D.nDeviceID = PAMTD2.nDeviceID " & _
 " where (PAMTD2.nActiveMonitorTypeID = 22 and " & sFinalDisableComment & _
 " and bDisabled = 0) " & sGroupQuery
 Dim sSqlToDisable1 : sSqlToDisable1 = "SELECT nPivotActiveMonitorTypeToDeviceID, PAMTD2.nDeviceID FROM PivotActiveMonitorTypeToDevice PAMTD2" & _
 " join ActiveMonitorType on ActiveMonitorType.nActiveMonitorTypeID = PAMTD2.nActiveMonitorTypeID" & _
 " where (PAMTD2.nActiveMonitorTypeID = 22 and " & sFinalDisableComment & _
 " and bDisabled = 0 and nPivotActiveMonitorTypeToDeviceID in (" & _
 " select nPivotActiveMonitorTypeToDeviceID from PivotActiveMonitorTypeToDevice PAMTD" & _
 " join MonitorState MS on MS.nMonitorStateID = PAMTD.nMonitorStateID" & _
 " where nInternalMonitorState = 1))" & sGroupQuery
 
 If bDownDisable = 0 Then sDisableQuery = sSqlToDisable
 If bDownDisable = 1 Then sDisableQuery = sSqlToDisable1
 'Run the disable/enable SQLs and send change events
 If bRunDisable Then
 	Context.NotifyProgress "Disable query: " & sDisableQuery
  DisableInterfacesSQL sDisableQuery
 Else
 	Context.NotifyProgress "*DISABLE* interfaces skipped due to bRunDisable flag"
 End If
 If bRunEnable Then
  Context.NotifyProgress "Enable query: " & sSqlToEnable
  EnableInterfacesSQL sSqlToEnable
 Else
 	Context.NotifyProgress "*ENABLE* interfaces skipped due to bRunDisable flag"
 End If
End If
Context.NotifyProgress nEnableCount & " devices had interfaces enabled"
Context.NotifyProgress nDisableCount & " devices had interfaces disabled"
'****END MAIN SCRIPT****

' **********************
' * GetDeviceGroupType *
' **********************
Sub GetDeviceGroupType(sDeviceGroup)
Context.NotifyProgress "****** GetDeviceGroupType sub procedure start." & vbCrLf
'This gets the device group ID
If Len(sDeviceGroup) > 0 Then
 Context.NotifyProgress "Finding group ID..."
 Dim nDeviceGroupID : nDeviceGroupID = GetSQLOutputToInt("select min(nDeviceGroupID) from DeviceGroup where sGroupName = '" & sDeviceGroup &"'")
 Context.NotifyProgress nDeviceGroupID & " is the group ID"
 If nDeviceGroupID > 0 Then
  'This gets whether that device group is dynamic or not
  Dim bDynamicGroup : bDynamicGroup = GetSQLOutputToBoolean("select bDynamicGroup from DeviceGroup where nDeviceGroupID = " & nDeviceGroupID)
  Context.NotifyProgress bDynamicGroup & " is the value of bDynamicGroup for nDeviceGroupID " & nDeviceGroupID
  If bDynamicGroup = 0 Then
 	 sGroupQuery = " and nPivotActiveMonitorTypeToDeviceID in (select nPivotActiveMonitorTypeToDeviceID from PivotDeviceToGroup PDTG " & _
 	 " join PivotActiveMonitorTypeToDevice PAMTD3 on PAMTD3.nDeviceID = PDTG.nDeviceID where nDeviceGroupID = " & nDeviceGroupID & ")"
  Else
 	 Dim sFilter : sFilter = GetSQLOutputToString("select sFilter from DeviceGroup where nDeviceGroupID = " & nDeviceGroupID)
 	 Context.NotifyProgress sFilter
	 sGroupQuery = " and PAMTD2.nDeviceID in ( " & sFilter &")"
  End If
Else
	Context.NotifyProgress "Group named " & sDeviceGroup & " was not found. Stopping script."
	bContinue = False
 End If
End If
Context.NotifyProgress "****** GetDeviceGroupType sub procedure end." & vbCrLf
End Sub

' ***********************
' * EnableInterfacesSQL *
' ***********************
Sub EnableInterfacesSQL(sSqlToEnable)
'Run the SQL statement
Dim sRS : Set sRS = oDB.Execute(sSqlToEnable)
If Not sRS.BOF And Not sRS.EOF Then
 sRS.moveFirst()
 While Not sRS.EOF
  nPivotActiveMonitorTypeToDeviceID = sRS("nPivotActiveMonitorTypeToDeviceID")
  nNewDevID = sRS("nDeviceID")
  Context.NotifyProgress "Processing device ID: " & nNewDevID
  EnableMonitor(nPivotActiveMonitorTypeToDeviceID)
  If nNewDevID <> nCurrDevID And nCurrDevID <> 0 Then
  	nEnableCount = nEnableCount + 1
  	SendChangeEvent(nCurrDevID)
  End If
  sRS.MoveNext()
  nCurrDevID = nNewDevID
  If sRS.EOF = True Then
   nEnableCount = nEnableCount + 1
   SendChangeEvent(nCurrDevID)
  End If
 Wend
End If
End Sub

' ************************
' * DisableInterfacesSQL *
' ************************
Sub DisableInterfacesSQL(sDisableQuery)
'Run the SQL statement
Dim sRS : Set sRS = oDB.Execute(sDisableQuery)
If Not sRS.BOF And Not sRS.EOF Then
 sRS.moveFirst()
 While Not sRS.EOF
  nPivotActiveMonitorTypeToDeviceID = sRS("nPivotActiveMonitorTypeToDeviceID")
  nNewDevID = sRS("nDeviceID")
  Context.NotifyProgress "Processing device ID: " & nNewDevID
  DisableMonitor(nPivotActiveMonitorTypeToDeviceID)
  If nNewDevID <> nCurrDevID And nCurrDevID <> 0 Then
  	nDisableCount = nDisableCount + 1
  	SendChangeEvent(nCurrDevID)
  End If
  sRS.MoveNext()
  nCurrDevID = nNewDevID
  If sRS.EOF = True Then
   nDisableCount = nDisableCount + 1
   SendChangeEvent(nCurrDevID)
  End If
 Wend
End If
End Sub

' *********************
' * GetSQLOutputToInt *
' *********************
Function GetSQLOutputToInt(sSql)
GetSQLOutputToInt = 0
Dim sRet, sRS
Context.NotifyProgress sSql
Set sRS = oDB.Execute(sSql)
If Not sRS.EOF And Not sRS.BOF Then
 sRet = Trim(sRS.GetString)
 If Len(sRet) > 0 Then
  If Right(sRet, 1) = vbCr Then
   sRet = Left(sRet, Len(sRet) - 1)
  End If
 End If
 If IsNumeric(sRet) Then
  GetSQLOutputToInt = CInt(sRet)
 End If
End If
End Function

' *************************
' * GetSQLOutputToBoolean *
' *************************
Function GetSQLOutputToBoolean(sSql)
GetSQLOutputToBoolean = 0
Dim sRet, sRS
Context.NotifyProgress sSql
Set sRS = oDB.Execute(sSql)
If Not sRS.EOF And Not sRS.BOF Then
 sRet = Trim(sRS.GetString)
 If Len(sRet) > 0 Then
 	If Right(sRet, 1) = vbCr Then
 		sRet = Left(sRet, Len(sRet) - 1)
 	End If
 End If
 If Len(sRet) > 0 Then
  GetSQLOutputToBoolean = CBool(sRet)
 End If
End If
End Function

' ************************
' * GetSQLOutputToString *
' ************************
Function GetSQLOutputToString(sSql)
GetSQLOutputToString = 0
Dim sRet, sRS
Context.NotifyProgress sSql
Set sRS = oDB.Execute(sSql)
If Not sRS.EOF And Not sRS.BOF Then
 sRet = Trim(sRS.GetString)
 If Right(sRet, 1) = vbCr Or Right(sRet, 1) = vbLf Then
  sRet = Left(sRet, Len(sRet) - 1)
 End If
 If Len(sRet) > 0 Then
  GetSQLOutputToString = Trim(sRet)
 End If
End If
End Function


' ******************
' * DisableMonitor *
' ******************
Sub DisableMonitor(nPivotID)
Dim sSql : sSql = "Update PivotActiveMonitorTypeToDevice set bDisabled = 1 where nPivotActiveMonitorTypeToDeviceID = " & nPivotID
If bDryRun = False Then 
 Dim sRS : Set sRS = oDb.Execute(sSql)
 Context.NotifyProgress sSql & " was run."
Else
 Context.NotifyProgress sSql & " would have run"
End If
End Sub

' ******************
' * EnableMonitor *
' ******************
Sub EnableMonitor(nPivotID)
Dim sSql : sSql = "Update PivotActiveMonitorTypeToDevice set bDisabled = 0 where nPivotActiveMonitorTypeToDeviceID = " & nPivotID
If bDryRun = False Then 
 Dim sRS : Set sRS = oDb.Execute(sSql)
 Context.NotifyProgress sSql & " was run."
Else
 Context.NotifyProgress sSql & " would have run"
End If
End Sub

' *******************
' * SendChangeEvent *
' *******************
Sub SendChangeEvent(nDeviceID)
'Variables for device change event
Const DCT_MODIFIED = 2
Const DCIT_DEVICE  = 1
If bDryRun = False Then 
 oEvent.SendChangeEvent DCT_MODIFIED, nDeviceID, DCIT_DEVICE
 Context.NotifyProgress nDeviceID & " was sent a change event."
Else
 Context.NotifyProgress nDeviceID & " would have been sent a change event."
End If
End Sub