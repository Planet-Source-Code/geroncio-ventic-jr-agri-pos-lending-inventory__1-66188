VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4305
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      DrawMode        =   10  'Mask Pen
      Height          =   135
      Left            =   11040
      ScaleHeight     =   75
      ScaleWidth      =   3435
      TabIndex        =   9
      Top             =   9120
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   195
      Left            =   13080
      TabIndex        =   8
      Top             =   9000
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   13065
      TabIndex        =   7
      Top             =   9000
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   3600
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin MSComctlLib.ProgressBar Progload 
         Height          =   135
         Left            =   240
         TabIndex        =   4
         Top             =   3720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   11
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   375
         Left            =   6240
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   3480
         Width           =   75
      End
      Begin VB.Label lblCopyright 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   15
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   3105
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblCopyright 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Copyright    © 2005"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   5280
         TabIndex        =   1
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Version  1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5280
         TabIndex        =   2
         Top             =   2940
         Width           =   1530
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   5280
         TabIndex        =   3
         Top             =   2580
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit





Private Sub Form_Load()
On Error Resume Next
Me.MousePointer = ccHourglass
Me.Text1.Text = frmLogin.txtUsername
Me.Text2.Text = frmLogin.txtpassword

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label1.ForeColor = &HFF&
 Progload.Value = Progload.Value + 0.25
 If Progload.Value = 0.25 Then
 Label1.Caption = " Loading...."
 ElseIf Progload.Value = 25 Then
Label1.Caption = "Connecting to DataBase....."
ElseIf Progload.Value = 75 Then
Label1.Caption = " Please Wait...."
End If

    'If the Progress Bar (ProgLoad) is 100% then your function happens.
    If Progload.Value = 100 Then
        
        'Your function, can be anything. Open another form
        'Unloads this form
        Me.Hide
        MDIForm1.Show
    End If
   
End Sub
