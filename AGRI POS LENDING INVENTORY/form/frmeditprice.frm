VERSION 5.00
Begin VB.Form frmeditprice 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Edit Price"
   ClientHeight    =   5010
   ClientLeft      =   4140
   ClientTop       =   3525
   ClientWidth     =   5745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   5745
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "&Set"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtnewprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtoldprice 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2640
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
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
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
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
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
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
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
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "New Price"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Old Price"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   645
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
      TabIndex        =   9
      Top             =   2160
      Width           =   1110
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
      TabIndex        =   8
      Top             =   840
      Width           =   795
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
      TabIndex        =   7
      Top             =   1320
      Width           =   630
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
      TabIndex        =   6
      Top             =   1680
      Width           =   330
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
      TabIndex        =   5
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frmeditprice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdset_Click()
Dim RScmb As New ADODB.Recordset
Dim CNcmb As New ADODB.Connection
Dim addnew As Boolean




Call connect(CNcmb, App.Path & "\myDB.mdb")
Call SetRs(RScmb, CNcmb, "SELECT * FROM stocks WHERE StockID='" & Combo1.Text & "'")

addnew = False
RScmb.Fields("UnitCOst") = txtnewprice.Text
RScmb.Update
RScmb.Requery

'With RScmb
'Text1.Text = .Fields("Item")
'Text2.Text = .Fields("Category")
'Text4.Text = .Fields("Unit")
'Text5.Text = .Fields("Quantity")
'txtoldprice.Text = .Fields("UnitCost")

'End With
MsgBox "New Price has been Set", vbInformation
Unload Me
Set RScmb = Nothing
Set CNcmb = Nothing
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
txtoldprice.Text = .Fields("UnitCost")

End With

Set RScmb = Nothing
Set CNcmb = Nothing
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
