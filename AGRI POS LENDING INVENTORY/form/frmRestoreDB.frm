VERSION 5.00
Begin VB.Form frmRestoreDB 
   Caption         =   "Restore DataBase"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   ScaleHeight     =   3270
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore Database"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2325
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   5535
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "..."
      Height          =   285
      Left            =   5880
      TabIndex        =   0
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Restore from"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Current Database Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label lblSelectedDba 
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
      Left            =   255
      TabIndex        =   4
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   2280
      Width           =   5895
   End
End
Attribute VB_Name = "frmRestoreDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbasize As Long
Dim dbasize2 As Long
Dim PathName As String

Private Sub cmdRestore_Click()
If MsgBox("Restoring database from location " & txtSource & " will replace existing database files.Do you want to Contunue", vbYesNo) = vbYes Then
DoRestore txtSource.Text, App.Path
If NoDba = True Then
 MsgBox "Database Restored Click Ok to Exit Program"
 frmRestoreDB.Hide
 Unload frmRestoreDB
End If
Else
 lblStatus.Caption = "Database Restore Canceled"
End If
End Sub

Private Sub cmdSource_Click()
On Error GoTo Erro
Dim strTemp As String
strTemp = fBrowseForFolder(Me.hwnd, "Restore From")
If strTemp <> "" Then
    txtSource = strTemp
    dbasize2 = FileLen(txtSource & "\myDB.MDB")
    lblSelectedDba = "Selected Backup Database is : " & Format((dbasize2 / 1024) / 1024, "standard") & "MB."
    cmdRestore.Enabled = True
End If
Erro:
    Select Case err.Number
       Case 53 'File Not Found
          lblSelectedDba = "No Backup at this location"
          cmdRestore.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB."
End Sub

Private Sub Form_Load()
PathName = App.Path & "\myDB.MDB"
On Error GoTo err
dbasize = FileLen(PathName)
err:
Exit Sub
End Sub


