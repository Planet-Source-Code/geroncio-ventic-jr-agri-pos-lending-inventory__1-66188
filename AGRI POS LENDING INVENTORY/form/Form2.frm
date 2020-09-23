VERSION 5.00
Begin VB.Form frmAddstock 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Stocks"
   ClientHeight    =   3495
   ClientLeft      =   4515
   ClientTop       =   3120
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7500
   Begin VB.ComboBox cbocategory 
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtstockID 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtdescription 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtunit 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtunitcost 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "StockID:"
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
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Unit Cost"
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
      Left            =   360
      TabIndex        =   11
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   2880
      Width           =   720
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
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Category:"
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
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Unit"
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
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   360
   End
End
Attribute VB_Name = "frmAddstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim CNAstocks As New ADODB.Connection
Dim RSAstocks As New ADODB.Recordset

Dim CNx As New ADODB.Connection
Dim RSx As New ADODB.Recordset

Dim count As String



Call connect(CNAstocks, App.Path & "\myDb.mdb")
Call SetRs(RSAstocks, CNAstocks, "SELECT * FROM Stocks WHERE Item='" & txtdescription.Text & "'")
If Not RSAstocks.EOF Then
MsgBox " Item Already Exist,only need to Replenish for new stocks to  be added", vbInformation, "Inventory"
txtdescription.Text = ""
cbocategory.Text = ""
txtunit.Text = ""
txtunitcost.Text = ""
txtQuantity.Text = ""

Exit Sub
Else
With RSAstocks
.AddNew
.Fields("StockID") = txtstockID.Text
.Fields("Item") = txtdescription.Text
.Fields("Category") = cbocategory.Text
.Fields("Unit") = txtunit.Text
.Fields("Unitcost") = txtunitcost.Text
.Fields("Quantity") = txtQuantity.Text
.Update
End With
MsgBox "Item has been added to the stocks", vbInformation, "Inventory"
End If
Set RSAstocks = Nothing
Set CNAstocks = Nothing

txtdescription.Text = ""
cbocategory.Text = ""
txtunit.Text = ""
txtunitcost.Text = ""
txtQuantity.Text = ""
txtstockID.Enabled = False


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call connect(CNx, App.Path & "\myDB.mdb")
Call SetRs(RSx, CNx, "SELECT * FROM Stocks")

count = RSx.RecordCount + 1
    Select Case Len(count)
        Case 1: txtstockID.Text = "0000" & RSx.RecordCount + 1
        Case 2: txtstockID.Text = "000" & RSx.RecordCount + 1
        Case 3: txtstockID.Text = "00" & RSx.RecordCount + 1
        Case 4: txtstockID.Text = "0" & RSx.RecordCount + 1
        Case 5: txtstockID.Text = RSx.RecordCount + 1
    End Select
Set CNx = Nothing
Set RSx = Nothing

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

txtstockID.Enabled = False

Dim CNcount As New ADODB.Connection
Dim RSCount As New ADODB.Recordset

Dim CNcategory As New ADODB.Connection
Dim rscategory As New ADODB.Recordset

Dim count As String

Call connect(CNcount, App.Path & "\myDB.mdb")
Call SetRs(RSCount, CNcount, "SELECT * FROM Stocks")

count = RSCount.RecordCount + 1
    Select Case Len(count)
        Case 1: txtstockID.Text = "0000" & RSCount.RecordCount + 1
        Case 2: txtstockID.Text = "000" & RSCount.RecordCount + 1
        Case 3: txtstockID.Text = "00" & RSCount.RecordCount + 1
        Case 4: txtstockID.Text = "0" & RSCount.RecordCount + 1
        Case 5: txtstockID.Text = RSCount.RecordCount + 1
    End Select
Set CNcount = Nothing
Set RSCount = Nothing

cbocategory.Clear
Call connect(CNcategory, App.Path & "\myDB.mdb")
Call SetRs(rscategory, CNcategory, "SELECT * FROM stockscategory order by stockscategory ASC")
With rscategory
    While Not .EOF
        cbocategory.AddItem .Fields("StocksCategory")
        .MoveNext
        Wend
    End With
Set CNcategory = Nothing
Set rscategory = Nothing

End Sub

Private Sub txtquantity_KeyPress(KeyAscii As Integer)
If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyInsert Then
KeyAscii = 0
End If
End Sub

Private Sub txtunitcost_KeyPress(KeyAscii As Integer)
If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyInsert Then
KeyAscii = 0
End If
End Sub
