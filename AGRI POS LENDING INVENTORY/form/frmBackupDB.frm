VERSION 5.00
Begin VB.Form frmBackupDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BackUp Database"
   ClientHeight    =   3705
   ClientLeft      =   4350
   ClientTop       =   3780
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5940
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   2640
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      DrawMode        =   10  'Mask Pen
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      ScaleHeight     =   285
      ScaleWidth      =   4695
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.TextBox txtDestination 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   4815
   End
   Begin VB.CommandButton cmdDestination 
      Caption         =   "..."
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup Database"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblDbaSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Current Database Size is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Backup Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   3975
   End
End
Attribute VB_Name = "frmBackupDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Value As Long         ' current progress value
Dim Interval As Double  ' amount to draw for each percent
Dim cap As String          ' caption

Dim dbasize As Long
Dim PathName As String

Private Sub cmdBackup_Click()
Me.MousePointer = 11
picProgress.BorderStyle = 1
If txtDestination <> "" Then
DoBackup PathName, txtDestination
ElseIf txtDestination = "" Then
  MsgBox "You must specify a distination for the backup", vbCritical
End If
picProgress.Visible = True
 Value = 0
    Interval = picProgress.ScaleWidth / 100
    cmdBackup.Enabled = False
    cmdClose.Enabled = False
    Timer1.Enabled = True

End Sub

Private Sub cmdClose_Click()
'Call coolClose(Me, 6)
Unload Me
End Sub

Private Sub cmdDestination_Click()
Dim strTemp As String

strTemp = fBrowseForFolder(Me.hwnd, "Select backup path")
If strTemp <> "" Then
    txtDestination = strTemp
End If
If txtDestination.Text = "" Then
cmdBackup.Enabled = False
Else
cmdBackup.Enabled = True
End If
End Sub

Private Sub Form_Activate()
lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB."
cmdBackup.Enabled = False
'txtDestination.Enabled = False
txtDestination.Locked = True

End Sub

Private Sub Form_Load()
'SetRegion
PathName = App.Path & "\myDB.MDB"
dbasize = FileLen(PathName)
End Sub
'Private Sub SetRegion()
 '   On Error Resume Next
  '  If hRgn Then DeleteObject hRgn
   ' hRgn = GetBitmapRegion(Me.Picture, RGB(255, 0, 255))
    'SetWindowRgn Me.hwnd, hRgn, True
'End Sub


Private Sub Timer1_Timer()
 Value = Value + 5
    
    ' set caption
    cap = Value & "%"
    
    
    With picProgress
        .Cls
        ' center the caption
        .CurrentX = (.ScaleWidth - .TextWidth(cap)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(cap)) \ 2
        picProgress.Print cap
        ' draw a filled rect
        picProgress.Line (0, 0)-(Interval * Value, .ScaleHeight), RGB(0, 0, 0), BF
        .Refresh
    End With
    
    ' stop at 100
    If Value = 100 Then
    Me.MousePointer = 0
    MsgBox "BackUp Has been Completed", vbInformation + vbOKOnly
Unload Me
         

    End If
    
    
    
End Sub
