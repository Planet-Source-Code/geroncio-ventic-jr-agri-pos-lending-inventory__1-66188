VERSION 5.00
Begin VB.Form frmCompactDB 
   Caption         =   "Compact DataBase"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form2"
   ScaleHeight     =   3870
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompactdba 
      Caption         =   "Compact Database"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmRepairDB.frx":0000
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label lblNewSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   4935
   End
   Begin VB.Label lblFreeSpace 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
End
Attribute VB_Name = "frmCompactDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbasize As Long
Private Sub cmdCompactdba_Click()
On Error GoTo err
If MsgBox("Are you sure", vbYesNo) = vbYes Then
      DB.Close
    If Dir(App.Path & "\CompactCon.mdb") <> "" Then
      Kill App.Path & "\CompactCon.mdb"
    End If
    DBEngine.CompactDatabase App.Path & "\myDB.mdb", App.Path & "\CompactCon.mdb", , , ";pwd=matrix-se"
    Kill App.Path & "\myDB.mdb"
    Name App.Path & "\CompactCon.mdb" As App.Path & "\myDB.mdb"
    PathName = App.Path & "\myDB.MDB"
    'On Error GoTo err
    dbasize = FileLen(PathName)
    lblNewSize = "Compacted Database size : " & Format((dbasize / 1024) / 1024, "standard") & "MB."
    OpenDB
End If

err:
 If err.Number = 3356 Then
   MsgBox "Error occured while trying to compact database Restart your Computer and try again", vbExclamation
   Exit Sub
End If
End Sub



Private Sub Form_Activate()
lblSize = "Current Database size: " & Format((dbasize / 1024) / 1024, "standard") & "MB."

End Sub

Private Sub Form_Load()
    Dim fs, d, s
    Dim drvpath As String
    Dim freeSpace As Long
    drvpath = App.Path
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvpath))
    freeSpace = d.AvailableSpace / 1024 / 1024
    s = "Drive " & Left(App.Path, 1) & " has "
    lblFreeSpace = s & FormatNumber(freeSpace, 0) & "MB free"
PathName = App.Path & "\myDB.MDB"
On Error GoTo err
dbasize = FileLen(PathName)
If freeSpace * 1024 * 1024 < dbasize Then
  lblNewSize = "Not enough space to compact database clear some space on drive " & Left(App.Path, 1)
  cmdCompactdba.Enabled = False
End If
err:
Exit Sub
End Sub

Private Sub Label1_Click()

End Sub

