VERSION 5.00
Begin VB.Form frmchangeuser 
   BackColor       =   &H00FDECD7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Username or Password"
   ClientHeight    =   6720
   ClientLeft      =   4905
   ClientTop       =   2745
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6600
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6375
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   600
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   1455
         ScaleWidth      =   1455
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNewUsername 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   2
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox txtNewPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         TabIndex        =   0
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&ancel"
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdchange 
         Caption         =   "&Change"
         Height          =   495
         Left            =   840
         TabIndex        =   4
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "New Username:"
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
         Left            =   1080
         TabIndex        =   15
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
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
         Left            =   960
         TabIndex        =   14
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   1200
         X2              =   5040
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Username:"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         Height          =   3255
         Left            =   840
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " NOTE: You must never forgot your USERNAME and   your  PASSWORD,  in oder to access this system."
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   3615
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3480
      TabIndex        =   9
      Top             =   10440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   10680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtOldPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3480
      TabIndex        =   7
      Top             =   10680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtOldUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3960
      TabIndex        =   6
      Top             =   10440
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmchangeuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdchange_Click()
On Error Resume Next
Dim Empcn As New ADODB.Connection
Dim Emprs As New ADODB.Recordset

If txtNewUsername.Text = "" Then MsgBox "Please Fill new Username", vbInformation, "Invalid": Exit Sub
If txtNewPassword.Text = "" Then MsgBox "Please Fill new Password", vbInformation, "Invalid": Exit Sub
txtOldUsername.Text = MDIForm1.Text1.Text
txtOldPassword.Text = MDIForm1.Text2.Text
If txtOldUsername.Text = Text4.Text Then
    If txtOldPassword.Text = Text3.Text Then
    Call connect(Empcn, App.Path & "\myDB.mdb")
    Call SetRs(Emprs, Empcn, "Select * FROM UserAccount WHERE Username like '" & txtOldUsername.Text & "'")

   
    
    Emprs.Fields(0) = txtNewUsername.Text
    Emprs.Fields(1) = txtNewPassword.Text
    Emprs.Update
    Emprs.Requery
    MsgBox "The Changes will take effect on your next Log- in", vbInformation
    
    txtOldUsername.Text = txtNewUsername.Text
    txtOldPassword.Text = txtNewPassword.Text
    txtNewUsername.Text = ""
    txtNewPassword.Text = ""
    Text4.Text = ""
    Text3.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Unload Me
    Else
    MsgBox "Password is not correct", vbCritical, "System  Setting"
    Exit Sub
    End If
Else
    MsgBox "Username is not correct", vbCritical, "System  Setting"
    Exit Sub
End If
Set Emprs = Nothing
Set Empcn = Nothing


End Sub




