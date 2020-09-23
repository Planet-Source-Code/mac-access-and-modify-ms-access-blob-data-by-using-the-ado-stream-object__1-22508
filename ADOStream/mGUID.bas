Attribute VB_Name = "mGUID"
'// THIS Code was a little bit re-designed to work with the MS Access Database
'// Feel free to use this code for your own needs
'// It would be nice if you could give a credit to me in ya AppZ
'// Please contact me for further information and questions:  marcuslauermann@gmx.net

Option Explicit

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

Public Function GetGUID() As String
'(c) 2001 Marcus Lauermann

Dim udtGUID As GUID

If (CoCreateGuid(udtGUID) = 0) Then

GetGUID = _
String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
End If

End Function


