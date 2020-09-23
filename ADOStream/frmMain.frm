VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMain 
   Caption         =   "HOWTO: Access and Modify MS Access 2000 BLOB Data by Using the ADO Stream Object"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows-Standard
   Begin RichTextLib.RichTextBox rtfComments 
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "some Comments"
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   11
      ToolTipText     =   "Telephone @Work"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Telephone - Private"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   5
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Country"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "City"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   7
      ToolTipText     =   "Zip-Code"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Street"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Surname"
      Top             =   600
      Width           =   3135
   End
   Begin MSDataListLib.DataCombo cboPID 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Select a Person an hit ENTER to load the Recordset"
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtPersonal 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Given Name"
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton cmdSave2DB 
      Caption         =   "Save to DB"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave2File 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   9720
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Article ID: Q258038"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HOWTO: Access and Modify SQL Server BLOB Data by Using the ADO Stream Object"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   5055
   End
   Begin VB.Image Picture1 
      BorderStyle     =   1  'Fest Einfach
      Height          =   2895
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// THIS Code was a little bit re-designed to work with the MS Access Database
'// Feel free to use this code for your own needs
'// It would be nice if you could give a credit to me in ya AppZ
'// Please contact me for further information and questions:  marcuslauermann@gmx.net

Option Explicit

Dim cn As CDatClass
Dim rs As Recordset
Dim bc As BindingCollection
Dim mstream As ADODB.Stream

Private Sub cboPID_KeyPress(KeyAscii As Integer)
If cboPID.Text = "" Then Exit Sub
    If KeyAscii = &HD Then
        LoadPersonal (cboPID.BoundText)
    End If
End Sub

Private Sub cmdSave2File_Click()

Set cn = New CDatClass
Set rs = New Recordset

'On Error GoTo Command1_Error

rs.Open "Select * FROM tblPersonal WHERE PID =" & cboPID.BoundText, cn.Conn, adOpenKeyset, adLockOptimistic

Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open
mstream.Write rs.Fields("Foto").Value

With cdOpen
    .FileName = txtPersonal(0).Text & txtPersonal(1).Text & ".bmp"
    .Filter = "Image (*.*)|*.*"
    .ShowSave
    
    If Len(.FileName) <> 0 Then
        mstream.SaveToFile .FileName, adSaveCreateOverWrite
        Picture1.Picture = LoadPicture(.FileName)
    End If
End With
rs.Close
cn.Conn.Close

MsgBox "Done saving image from database to " & cdOpen.FileName
Exit Sub
Command1_Error:
    MsgBox Str(err) & " - " & Error, vbExclamation
    
End Sub

Private Sub cmdSave2DB_Click()

Set cn = New CDatClass
Set rs = New ADODB.Recordset

rs.Open "Select * FROM tblPersonal WHERE PID =" & cboPID.BoundText, cn.Conn, adOpenKeyset, adLockOptimistic

On Error GoTo Command2_Error

Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open

With cdOpen
    .FileName = ""
    .Filter = "Image (*.*)|*.*"
    .ShowOpen
    
    If Len(.FileName) <> 0 Then
        
        mstream.LoadFromFile .FileName
        rs.Fields("Foto").Value = mstream.Read
        Picture1.Picture = LoadPicture(.FileName)
        rs.Update
    End If
End With

rs.Close
cn.Conn.Close
MsgBox "Done saving image " & cdOpen.FileName
Exit Sub
Command2_Error:
    MsgBox Str(err) & " - " & Error, vbExclamation
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()

    Set cn = New CDatClass
    Set rs = New Recordset
    
    rs.Open "Select *,Vorname+' '+Nachname AS PNAME FROM tblPersonal ORDER BY Vorname ASC", cn.Conn, adOpenKeyset, adLockOptimistic

        With cboPID
        Set .DataSource = rs
            .DataField = "PID"
        Set .RowSource = rs
            .BoundColumn = "PID"
            .BoundText = rs!PID
            .ListField = "PNAME"
            .Text = ""
        End With

End Sub

Private Sub LoadPersonal(sPID As String)
Me.MousePointer = vbHourglass
    Set cn = New CDatClass
    Set rs = New Recordset
    Set bc = New BindingCollection
        rs.Open "SELECT * FROM tblPersonal WHERE PID=" & sPID, cn.Conn, adOpenKeyset, adLockOptimistic
    With bc
    
    Set .DataSource = rs
        .Add txtPersonal(0), "Text", "Vorname"
        .Add txtPersonal(1), "Text", "Nachname"
        .Add txtPersonal(2), "Text", "Straße"
        .Add txtPersonal(3), "Text", "PLZ"
        .Add txtPersonal(4), "Text", "Ort"
        .Add txtPersonal(5), "Text", "Land"
        .Add txtPersonal(6), "Text", "TelPrv"
        .Add txtPersonal(7), "Text", "TelWrk"
        .Add rtfComments, "Text", "Bemerkungen"
    End With
    
    DisplayBlob
    
Me.MousePointer = vbDefault
End Sub

Private Sub DisplayBlob()
Dim sFilename As String
Set cn = New CDatClass
Set rs = New Recordset

On Error GoTo Command1_Error
sFilename = App.Path & "\Temp\" & mGUID.GetGUID & ".tmp"

rs.Open "Select * FROM tblPersonal WHERE PID =" & cboPID.BoundText, cn.Conn, adOpenKeyset, adLockOptimistic

If rs.RecordCount = 0 Then cmdSave2DB.Enabled = False: cmdSave2File.Enabled = False: Picture1.Picture = Nothing: Exit Sub

Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open
mstream.Write rs.Fields("Foto").Value

mstream.SaveToFile sFilename, adSaveCreateOverWrite
        Picture1.Picture = LoadPicture(sFilename)

cmdSave2DB.Enabled = True: cmdSave2File.Enabled = True

rs.Close
cn.Conn.Close

Command1_Error:
    If err.Number <> 0 Then MsgBox Str(err) & " - " & Error, vbExclamation
End Sub
