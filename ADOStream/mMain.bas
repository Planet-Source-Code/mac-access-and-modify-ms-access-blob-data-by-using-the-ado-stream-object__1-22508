Attribute VB_Name = "mMain"
'// THIS Code was a little bit re-designed to work with the MS Access Database
'// Feel free to use this code for your own needs
'// It would be nice if you could give a credit to me in ya AppZ
'// Please contact me for further information and questions:  marcuslauermann@gmx.net

Option Explicit

Public sDBPath As Variant

Dim cn As CDatClass
Dim rs As Recordset

Sub Main()

Set cn = New CDatClass
Set rs = New Recordset

    sDBPath = getstring(HKEY_CURRENT_USER, "Software\ACGC", "dbpath")

        rs.Open "SELECT * from tblPersonal", cn.Conn, adOpenKeyset, adLockReadOnly

    Debug.Print "tblPersonal - RecordCount: " & rs.RecordCount

frmMain.Show

End Sub
