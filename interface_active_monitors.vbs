' This Action Script is NOT Device assigned and should be used 
' as a global recurring action that could be scheduled or run manually as needed.
Option Explicit
'DO NOT MODIFY ABOVE THIS LINE AT RISK OF BREAKING THE SCRIPT
' ******************* Configuration *************************
'This should match a portion of the comment of the interface active monitors you'd like to *ENABLE*
Dim sEnableComment : sEnableComment = "xs"

'This should match a portion of the comment of the interface active monitors you'd like to *DISABLE*
Dim sDisableComment : sDisableComment = ""

'When disabling, 0 means disable all that match, 1 means disable only those that match and are down
Dim bDownDisable : bDownDisable = 0

'Name(s) of device primary roles to consider
Dim aRoles : aRoles = "Switch,Router,Firewall"

' ***************** End Configuration ***********************
'DO NOT MODIFY BELOW THIS LINE AT RISK OF BREAKING THE SCRIPT

'This query can be modified to enable/disable interface active monitors or enable/disable those NOT LIKE '%sComment%'
Dim sSqlToEnable : sSqlToEnable = "SELECT nPivotActiveMonitorTypeToDeviceID, PAMTD2.nDeviceID FROM PivotActiveMonitorTypeToDevice PAMTD2" & _
" join ActiveMonitorType on ActiveMonitorType.nActiveMonitorTypeID = PAMTD2.nActiveMonitorTypeID " & _
" join Device D on D.nDeviceID = PAMTD2.nDeviceID " & _
" join DeviceType DT on DT.nDeviceTypeID = D.nDeviceTypeID" & _
" where (PAMTD2.nActiveMonitorTypeID = 22 and sComment like '%" & sComment & "%' " & _
" and bDisabled = 1)"

Dim sSqlToEnable1 : sSqlToDisable = "SELECT nPivotActiveMonitorTypeToDeviceID, PAMTD2.nDeviceID FROM PivotActiveMonitorTypeToDevice PAMTD2" & _
" join ActiveMonitorType on ActiveMonitorType.nActiveMonitorTypeID = PAMTD2.nActiveMonitorTypeID " & _
" join Device D on D.nDeviceID = PAMTD2.nDeviceID " & _
" join DeviceType DT on DT.nDeviceTypeID = D.nDeviceTypeID" & _
" where (PAMTD2.nActiveMonitorTypeID = 22 and sComment like '%" & sComment & "%' " & _
" and bDisabled = 0)"

If bToDisable = 0 Then sQuery = sSqlToDisable
If bToDisable = 1 Then sQuery = sSqlToDisable & " and nInternalstate=down"

'Create DB object
Dim oDB : Set oDB = Context.GetDB
'Create WUG Event Helper Object
Dim oEvent : Set oEvent = CreateObject("CoreAsp.EventHelper")

'Variable declaration
Dim nPivotActiveMonitorTypeToDeviceID, nNewDevID, nCount, sQuery
Dim nCurrDevID : nCurrDevID = 0


'Run the SQL statement
Dim sRS : Set sRS = oDB.Execute(sQuery)
If Not sRS.BOF And Not sRS.EOF Then
 sRS.moveFirst()
 While Not sRS.EOF
  nPivotActiveMonitorTypeToDeviceID = sRS("nPivotActiveMonitorTypeToDeviceID")
  nNewDevID = sRS("nDeviceID")
  Context.NotifyProgress "Processing device ID: " & nNewDevID
  DisableMonitor(nPivotActiveMonitorTypeToDeviceID)
  If nNewDevID <> nCurrDevID And nCurrDevID <> 0 Then
  	SendChangeEvent(nCurrDevID)
  End If
  sRS.MoveNext()
  nCurrDevID = nNewDevID
  If sRS.EOF = True Then
   SendChangeEvent(nCurrDevID)
  End If
 Wend
End If

' *********************
' * GetSQLOutputToInt *
' *********************
Function GetSQLOutputToInt(sSql)
GetSQLOutputToInt = 0
Dim sRet, sRS
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

' ******************
' * DisableMonitor *
' ******************
Sub DisableMonitor(nPivotID)
Dim sSql : sSql = "Update PivotActiveMonitorTypeToDevice set bDisabled = 1 where nPivotActiveMonitorTypeToDeviceID = " & nPivotID
Dim sRS : Set sRS = oDb.Execute(sSql)
Context.NotifyProgress sSql & " was run."
End Sub

' ******************
' * EnableMonitor *
' ******************
Sub EnableMonitor(nPivotID)
Dim sSql : sSql = "Update PivotActiveMonitorTypeToDevice set bDisabled = 0 where nPivotActiveMonitorTypeToDeviceID = " & nPivotID
Dim sRS : Set sRS = oDb.Execute(sSql)
Context.NotifyProgress sSql & " was run."
End Sub

' *******************
' * SendChangeEvent *
' *******************
Sub SendChangeEvent(nDeviceID)
'Variables for device change event
Const DCT_MODIFIED = 2
Const DCIT_DEVICE  = 1
oEvent.SendChangeEvent DCT_MODIFIED, nDeviceID, DCIT_DEVICE
Context.NotifyProgress nDeviceID & " was sent a change event."
End Sub
