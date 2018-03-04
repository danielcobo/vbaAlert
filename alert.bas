Attribute VB_Name = "alert_"
Option Explicit

'Display a message
'-----------------
'Returns true or false based on user responese
Function alert(str) As Boolean
    MsgBox str, vbInformation + vbOKOnly + vbApplicationModal + vbMsgBoxSetForeground
    alert = True
End Function

'Display an error message
'------------------------
'returns true or false based on user response
Function alertErr(str) As Boolean
    MsgBox str, vbCritical + vbOKOnly + vbApplicationModal + vbMsgBoxSetForeground
    alertErr = True
End Function

'Yes/No dialog
'-------------
'Returns true or false based on user response
Function alertYesNo(str As String) As Boolean
    Dim result As Integer
    result = MsgBox(str, vbQuestion + vbYesNo + vbApplicationModal + vbMsgBoxSetForeground)
    
    If result = 6 Then
        alertYesNo = True
    Else   'result = 7
        alertYesNo = False
    End If
End Function

'OK/Cancel dialog
'----------------
'Returns true or false based on user response
Function alertOKCancel(str As String) As Boolean
    Dim result As Integer
    result = MsgBox(str, vbQuestion + vbOKCancel + vbApplicationModal + vbMsgBoxSetForeground)

    If result = 1 Then
        alertOKCancel = True
    Else   'result = 2
        alertOKCancel = False
    End If
End Function

'Input prompt
'------------
'Returns value user inputed
'or "" if canceled or left empty by user
Function alertInput(strMsg As String) As String
    alertInput = InputBox(strMsg)
End Function
