VERSION 5.00
Begin VB.Form frmAddnewUser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   2970
   ClientLeft      =   4920
   ClientTop       =   4365
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4650
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtconfirmpassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Confirm Pasword"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Password"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "User Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddnewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub Text1_Change()

End Sub

Private Sub cmdOk_Click()
Dim CN1 As New ADODB.Connection
Dim RS1 As New ADODB.Recordset

    If txtUsername.Text = "" Then
        MsgBox "Please Filled all the fields", vbInformation
        Exit Sub
        End If
            If txtpassword.Text = "" Then
                MsgBox "Please Filled all the fields", vbInformation
                Exit Sub
                End If
                    If txtconfirmpassword.Text = "" Then
                        MsgBox "Please Filled all the fields", vbInformation
                        Exit Sub
                        End If




If txtpassword.Text <> Me.txtconfirmpassword.Text Then
MsgBox "Verified password is incorrect", vbInformation
Exit Sub
End If

Call connect(CN1, App.Path & "\mydb.mdb")
Call SetRs(RS1, CN1, "select * from UserAccount WHERE Username ='" & txtUsername & "'")

    If Not RS1.EOF Then
        MsgBox "Username already exist", vbInformation
        Me.txtUsername.SelStart = 0
        Me.txtUsername.SelLength = Len(txtUsername.Text)
        Me.txtUsername.SetFocus

            Else
            
                RS1.AddNew
                RS1.Fields("Username") = txtUsername.Text
                RS1.Fields("Password") = txtpassword.Text
                RS1.Update
                RS1.Requery

                    MsgBox "New user has been Successfully Added", vbInformation
                    Me.txtUsername = ""
                    Me.txtpassword = ""
                    Me.txtconfirmpassword = ""
                    Me.txtUsername.SetFocus
End If

Set CN1 = Nothing
Set RS1 = Nothing
End Sub
