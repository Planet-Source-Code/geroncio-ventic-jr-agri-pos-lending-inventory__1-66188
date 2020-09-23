VERSION 5.00
Begin VB.Form frmReplenish 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replenish Stocks"
   ClientHeight    =   3675
   ClientLeft      =   3540
   ClientTop       =   4080
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8205
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text2 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text4 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text5 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text6 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Replenish"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   5640
      ScaleHeight     =   1905
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.PictureBox Picture2 
         Height          =   2175
         Left            =   0
         Picture         =   "frmreplenishequipment.frx":0000
         ScaleHeight     =   2115
         ScaleWidth      =   2355
         TabIndex        =   1
         Top             =   -240
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "StockID"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Unit "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Description"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Received Quantity"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Stock OnHand:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmReplenish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
'Dim RScmb As New adodb.Recordset
'Dim CNcmb As New adodb.Connection

'Call connect(CNcmb, App.Path & "\Tailor.mdb")
'Call setRs(RScmb, CNcmb, "SELECT * FROM stocks WHERE StockID='" & Combo1.Text & "'")

'With RScmb
'Text1.Text = .Fields("Item")
'Text2.Text = .Fields("Category")
'Text4.Text = .Fields("Unit")
'Text5.Text = .Fields("Quantity")
'End With

'Set RScmb = Nothing
'Set CNcmb = Nothing

End Sub

Private Sub Combo1_Click()
Dim RScmb As New ADODB.Recordset
Dim CNcmb As New ADODB.Connection



Call connect(CNcmb, App.Path & "\myDB.mdb")
Call SetRs(RScmb, CNcmb, "SELECT * FROM stocks WHERE StockID='" & Combo1.Text & "'")

With RScmb
Text1.Text = .Fields("Item")
Text2.Text = .Fields("Category")
Text4.Text = .Fields("Unit")
Text5.Text = .Fields("Quantity")
End With

Set RScmb = Nothing
Set CNcmb = Nothing
End Sub

Private Sub Command1_Click()
Dim CNR As New ADODB.Connection
Dim RSR As New ADODB.Recordset
Dim ans As String


If Text6.Text = "" Then
    MsgBox "Can not replenish, Input Received Quantity to replenish", vbCritical
Exit Sub

Else
Call connect(CNR, App.Path & "\myDB.mdb")
Call SetRs(RSR, CNR, "SELECT * FROM Stocks WHERE StockID ='" & Combo1.Text & "'")

With RSR
.Fields("Quantity") = .Fields("Quantity") + Val(Text6.Text)
.Update
End With
Set RSR = Nothing
Set CNR = Nothing

MsgBox "Stocks has been Replenished", vbInformation
ans = MsgBox("Would you like to Replenish Another Item? ", vbYesNo)
If ans = vbYes Then
Combo1.Text = ""
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Else
Unload Me

End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim CNre As New ADODB.Connection
Dim RSre As New ADODB.Recordset

Text1.Enabled = False
Text2.Enabled = False
Text4.Enabled = False
Text5.Enabled = False



Call connect(CNre, App.Path & "\myDB.mdb")
Call SetRs(RSre, CNre, "SELECT * FROM Stocks Order by stockID ASC")

Combo1.Clear

With RSre
    While Not .EOF
        Combo1.AddItem .Fields("StockID")
        .MoveNext
    Wend
End With
Set RSre = Nothing
Set CNre = Nothing

End Sub


Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyInsert Then
KeyAscii = 0
End If
End Sub
