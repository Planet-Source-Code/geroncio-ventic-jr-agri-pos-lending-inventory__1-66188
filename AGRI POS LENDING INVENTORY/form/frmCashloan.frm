VERSION 5.00
Begin VB.Form frmCashloan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Loan"
   ClientHeight    =   3090
   ClientLeft      =   4905
   ClientTop       =   4080
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   4425
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtamount 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtcn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Contract Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmCashloan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
Dim CNloan As New ADODB.Connection
Dim RSloan As New ADODB.Recordset

Call connect(CNloan, App.Path & "\myDB.mdb")
Call SetRs(RSloan, CNloan, "SELECT * FROM Debtrecord")

RSloan.AddNew
RSloan.Fields("contractnumber") = Me.txtcn.Text
RSloan.Fields("name") = Me.txtname.Text
RSloan.Fields("Item") = "Cash"
RSloan.Fields("Quantity") = "Cash"
RSloan.Fields("cost") = Me.txtamount.Text
RSloan.Fields("totalcost") = Me.txtamount.Text
RSloan.Fields("Date") = Date

RSloan.Update
RSloan.Requery

MsgBox "new Cash Loan has been Added", vbInformation
Unload Me

Set RSloan = Nothing
Set CNloan = Nothing


End Sub

Private Sub txtamount_KeyPress(KeyAscii As Integer)
If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyInsert Then
KeyAscii = 0
End If
End Sub
