VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Introduction:
'
'dbCon (Database Connection) was developed to make opening and managing ADO
'connections a little easier. I hope you find it useful and would be happy
'to hear any suggestions or comments you might have.  Since this is my first
'public offering I am eager for feedback.  I can be reached at praxis454@hotmail.com
'
'Error Index:
'
'8000:  Name aready in use.
'       Occurs when user attempts to assign a name to a new recordset that is the
'       same as a recordset currently open.
'8001:  Connection string not set.
'       Occurs when user attempts to open a conneciton to a database without first
'       setting a connection string
'8002:  Connection string can not be zero length
'       Occurs when user attempts to set a blank as connection string
'8003:  Recordset name can not be found
'       Occurs when the user attempts to refrence a recordset that is not currently
'       open or has never been opened
'8099:  Unanticipated Error
'       Occurs when the error that occurs is not handled

Dim strConnection As String
Dim cnnConnection As ADODB.Connection
Dim rstRecord() As ADODB.Recordset
Dim blnConnected As Boolean
Dim intRecordCount As Integer
Dim strRecName() As String
Dim intCurRecOpen As Integer
Dim strName As String

Public Enum ConnectionState
    dcOpen = 1
    dcClosed = 2
End Enum

Public Sub Connect()
'Purpose: Open connection to the database
'Created: 5-16-00
'Last Modified: 5-17-00

On Error GoTo errH

If strConnection = "" Then
    Err.Raise 8001, "dbCon", "Connection string must be set before attempting to open a connection"
    Exit Sub
End If

If blnConnected = True Then
    If MsgBox("Connection to data base already made, do you wish to use the new connection", vbYesNo, "Disconnect?") = vbNo Then
        Exit Sub
    End If
End If

    
Set cnnConnection = New ADODB.Connection
cnnConnection = strConnection
cnnConnection.Open

blnConnected = True

Exit Sub
errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & " Source:clsdbCon_Connect"
        Exit Sub
End Select

End Sub

Public Sub Disconnect()
'Purpose: Close connection to the database
'Created: 5-16-00
'Last Modified: 5-17-00

On Error GoTo errH

If cnnConnection.State = 1 Then
    cnnConnection.Close
End If

blnConnected = False

Exit Sub

errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & " Source: clsdbCon_Disconnect"
        Exit Sub
End Select

End Sub

Public Sub OpenRecordset(Name As String, SQL As String)
'Purpose:   Open a recordset assigning it the identifier variable Name, and opening
'           the conneciton with the provide variable SQL
'Created: 5-17-00
'Last Modified: 5-17-00

On Error GoTo errH
Dim i As Integer

Name = StrConv(Name, vbUpperCase)

'Check if recordset name is already in use
If intRecordCount <> 0 Then
    For i = 0 To intRecordCount
        If strRecName(i) = Name Then
            Err.Raise 8000, "dbCon", "Recordset name " & StrConv(Name, vbProperCase) & " already in use."
            Exit Sub
        End If
    Next i
End If

'increase the count of records open by 1
intRecordCount = intRecordCount + 1

'Redim the recordset array to reflect the new count of recordsets opened
ReDim Preserve rstRecord(intRecordCount)

'Open the new recordset
Set rstRecord(intRecordCount) = New ADODB.Recordset

rstRecord(intRecordCount).CursorLocation = adUseClient
rstRecord(intRecordCount).Open SQL, cnnConnection, adOpenDynamic, adLockBatchOptimistic

'Redim the Recordset name look up index and add the new name
ReDim Preserve strRecName(intRecordCount)
strRecName(intRecordCount) = Name

intCurRecOpen = intCurRecOpen + 1


Exit Sub
errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & " Source:clsdbCon_OpenRecord"
        Exit Sub
End Select

End Sub

Public Sub CloseRecordset(Name As String)
'Purpose:   Close an open recordset
'Created: 5-17-00
'Last Modified: 5-17-00

On Error GoTo errH
Dim i As Integer
Name = StrConv(Name, vbUpperCase)

For i = 0 To intRecordCount
    If strRecName(i) = Name Then
        rstRecord(i).Close
        strRecName(i) = ""
        Exit Sub
    End If
Next i

If i = intRecordCount + 1 Then
    Err.Raise 8003, "dbCon", StrConv(Name, vbProperCase) & " can not be found, make sure you have opened this recordset"
    Exit Sub
End If

intCurRecOpen = intCurRecOpen - 1

Exit Sub
errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & " Source:dbCon_CloseRecord"
        Exit Sub
End Select

End Sub

Property Let ConnectionString(Connection As String)
'Purpose: Set the database conneciton string
'Created: 5-16-00
'Last Modified: 5-17-00

On Error GoTo errH

If Connection = "" Then
    Err.Raise 8002, "dbCon", "Connection can not be a zero length value"
    Exit Property
End If

strConnection = Connection

Exit Property
errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & " Source: clsdbCon_ConnectionString"
        Exit Property
End Select

End Property

Property Get Recordset(Name As String) As ADODB.Recordset
'Purpose:   Allow user to access recordset
'Created: 5-17-00
'Last Modified: 5-17-00

On Error GoTo errH
Dim i As Integer
Name = StrConv(Name, vbUpperCase)

For i = 1 To intRecordCount
If Name = strRecName(i) Then
    Set Recordset = rstRecord(i)
    Exit Sub
End If
Next i

If i = intRecordCount + 1 Then
    Err.Raise 8003, "dbCon", StrConv(Name, vbProperCase) & " can not be found, make sure you have opened this recordset"
    Exit Sub
End If
Exit Sub
errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & " Source:dbCon_Recordset"
        Exit Sub
End Select


End Property

Property Get Status()

On Error GoTo errH

If blnConnected = True Then
    Status = dcOpen
Else
    Status = dcClosed
End If

    
Exit Property
errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & "Source:dbCon_Status"
        Exit Property
End Select
End Property


Property Get RecordsetCount()
On Error GoTo errH


RecordsetCount = intCurRecOpen
    
Exit Property
errH:
Select Case Err.Number
    Case Else
        Err.Raise 8099, "dbCon", Err.Number & " " & Err.Description & "Source:dbCon_RecordsetCount"
        Exit Property
End Select
End Property


