VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDatClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38AABC080324"
'// THIS Code was a little bit re-designed to work with the MS Access Database
'// Feel free to use this code for your own needs
'// It would be nice if you could give a credit to me in ya AppZ
'// Please contact me for further information and questions:  marcuslauermann@gmx.net

Option Explicit
Public Conn As Connection

'//Initialisierung der Klasse
Private Sub Class_Initialize()
On Error GoTo err
Set Conn = New Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & App.Path & "\blob.mdb" & ";" & _
"Mode=ReadWrite|Share Deny None;Persist Security Info=False;Jet OLEDB"
Conn.Open App.Path & "\blob.mdb"
err:
Debug.Print "Class_Initialize: " & err.Number & " " & err.Description: Resume Next
End Sub

'//Klassendaten löschen
Private Sub Class_Terminate()
On Error GoTo err
    Set Conn = Nothing
err:
Debug.Print "Class_Terminate: " & err.Number & " " & err.Description: Resume Next
End Sub


