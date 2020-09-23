VERSION 5.00
Begin VB.Form frmOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Item"
   ClientHeight    =   4980
   ClientLeft      =   4515
   ClientTop       =   3120
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6375
   Begin VB.CommandButton cmdCash 
      Caption         =   "&Cash"
      Height          =   495
      Left            =   2640
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   13110
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   360
      ScaleHeight     =   3585
      ScaleWidth      =   5625
      TabIndex        =   4
      Top             =   360
      Width           =   5655
      Begin VB.TextBox txtCN 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtCost 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblcontract 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contract Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   12
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   10
         Top             =   2520
         Width           =   390
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   13605
      TabIndex        =   3
      Top             =   4590
      Width           =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   13890
      TabIndex        =   0
      Top             =   5340
      Width           =   45
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CommandMode As Boolean


Private Sub cmdCash_Click()
Dim cndebt As New ADODB.Connection
Dim rsdebt As New ADODB.Recordset
'Dim reply As String
'reply = ""
'reply = InputBox("Enter Contract Number", "Add Cash Loan Record")
'If reply = "" Then
'Exit Sub

'Else
Call connect(cndebt, App.Path & "\myDB.mdb")
Call SetRs(rsdebt, cndebt, "Select * from custprofile Where contractnumber = '" & txtcn.Text & "'")
Unload Me
If Not rsdebt.EOF Then
frmCashloan.txtcn.Text = rsdebt.Fields("contractNumber")
frmCashloan.txtname.Text = rsdebt.Fields("Name")
frmCashloan.Show vbModal
Else

MsgBox "Contract Number Don't Exist", vbInformation
Exit Sub




End If
Set cndebt = Nothing
Set rsdebt = Nothing

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim CNs As New ADODB.Connection
Dim Rss As New ADODB.Recordset
Dim reply As String




Call connect(CNs, App.Path & "\myDB.mdb")
Call SetRs(Rss, CNs, "SELECT * FROM Debtrecord")


CommandMode = True
Rss.addnew
Rss.Fields("ContractNumber") = Me.txtcn.Text
Rss.Fields("Name") = Me.txtname.Text
Rss.Fields("Item") = Me.Combo1.Text
Rss.Fields("Quantity") = Me.txtQuantity.Text
Rss.Fields("cost") = Me.txtCost.Text
Rss.Fields("TotalCOst") = Me.Text5.Text
Rss.Fields("Date") = Date
Rss.Update
Rss.Requery

Set CNs = Nothing
Set Rss = Nothing



Dim Cnss As New ADODB.Connection
Dim Rsss As New ADODB.Recordset


Call connect(Cnss, App.Path & "\myDB.mdb")
Call SetRs(Rsss, Cnss, "SELECT * FROM STocks WHERE Stocks.item= '" & Combo1.Text & "'")

CommandMode = False
Rsss.Fields("Quantity") = Me.Text2.Text
Rsss.Update


MsgBox " New Debt Item Has been Added", vbInformation
reply = MsgBox("Would you like to Add another Item?", vbYesNo)
If reply = vbYes Then
txtQuantity.Text = ""
Else
Unload Me
End If



Set Cnss = Nothing
Set Rsss = Nothing

End Sub

Private Sub Combo1_Change()
Dim RSe As New ADODB.Recordset
Dim cne As New ADODB.Connection

Call connect(cne, App.Path & "\mydb.mdb")
Call SetRs(RSe, cne, "SELECT * FROM stocks WHERE stocks.item = '" & Combo1.Text & "'")

Me.txtCost.Text = RSe.Fields("UnitCost")
Me.Text1.Text = RSe.Fields("Quantity")

Set RSe = Nothing
Set cne = Nothing
End Sub

Private Sub Combo1_Click()
Dim RSe As New ADODB.Recordset
Dim cne As New ADODB.Connection

Call connect(cne, App.Path & "\mydb.mdb")
Call SetRs(RSe, cne, "SELECT * FROM stocks WHERE stocks.item = '" & Combo1.Text & "'")

Me.txtCost.Text = RSe.Fields("UnitCost")
Me.Text1.Text = RSe.Fields("Quantity")

Set RSe = Nothing
Set cne = Nothing
End Sub

Private Sub Form_Load()
Dim RSre As New ADODB.Recordset
Dim CNre As New ADODB.Connection

cmdSave.Enabled = False
Call connect(CNre, App.Path & "\mydb.mdb")
Call SetRs(RSre, CNre, "SELECT * FROM Stocks Order by stockid asc")
Combo1.Clear

With RSre
    While Not .EOF
        Combo1.AddItem .Fields("Item")
        .MoveNext
    Wend
End With
Set RSre = Nothing
Set CNre = Nothing
End Sub

Private Sub txtQuantity_Change()

Me.Text5.Text = Val(txtQuantity.Text) * Val(txtCost.Text)
Text2.Text = Val(Text1.Text) - Val(txtQuantity.Text)
If (Val(Text1.Text) <= Val(txtQuantity.Text)) Then
cmdSave.Enabled = False
MsgBox " Not enough Stocks on hand" & vbCrLf & "" & vbCrLf & "Need to replenish stock", vbInformation
txtQuantity.SetFocus
Exit Sub
Else
cmdSave.Enabled = True

End If








End Sub
