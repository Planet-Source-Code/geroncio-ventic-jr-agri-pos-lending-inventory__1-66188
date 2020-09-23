VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   5490
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3243.672
   ScaleMode       =   0  'User
   ScaleWidth      =   7436.452
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   4800
      TabIndex        =   5
      Top             =   3480
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5040
      TabIndex        =   0
      Top             =   1800
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   6240
      TabIndex        =   2
      Top             =   3480
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2520
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   3
      Top             =   1800
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   -720
      Width           =   9000
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Dim cnuser As New ADODB.Connection
Dim rsuser As New ADODB.Recordset
Dim retvalue As Integer
Dim Worm As Integer

Private Sub cmdcancel_Click()
   
    LoginSucceeded = False
    
    End
End Sub

Private Sub cmdok_Click()
     On Error Resume Next
   
   Call connect(cnuser, App.path & "\myDB.mdb")
   Call SetRs(rsuser, cnuser, "Select * From UserAccount")
   
   If Found(rsuser, "UserName", txtUserName.Text) = False Then
    
        
        
   MsgBox "User Name doesn't Exist", vbInformation, "Agri System"
   
   Else
   
   If password = txtPassword.Text Then
   Me.Hide
   frmSplash.Show
   frmSplash.Enabled = True
   
   Else
   MsgBox "User Name supplied and or Password is invalid....", vbInformation, "Agri System"
   
   End If
   End If
   
   Set cnuser = Nothing
   Set rsuser = Nothing
   txtUserName.Text = ""
   txtPassword.Text = ""
   txtUserName.SetFocus
   
   
End Sub

Private Sub Form_Load()
On Error Resume Next
txtUserName.TabIndex = 0
Call developer


End Sub


Public Sub developer()
On Error Resume Next
retvalue = GetSetting("A", "0", "Runcount")
Worm = Val(retvalue) + 1
SaveSetting "A", "0", "RunCount", Worm

If Worm > 99999 Then
    MsgBox "This is the End of the trial run" & vbCrLf & "This Is only  Good for Thesis ", 16, "Sorry"
    MsgBox "Email the developer of this system at" & vbCrLf & "venticjojo@yahoo.com" & vbCrLf & "Venticjojo05@hotmail.com" & vbCrLf & "or call 09186070112" & vbCrLf & "For the reactivation or full version of this system", 6, "System Developer"
    Unload Me
End If
End Sub

