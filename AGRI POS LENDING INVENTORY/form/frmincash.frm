VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmincash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "In Cash Basis"
   ClientHeight    =   5100
   ClientLeft      =   2955
   ClientTop       =   3705
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   10095
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3465
      ScaleWidth      =   3945
      TabIndex        =   14
      Top             =   240
      Width           =   3975
      Begin VB.TextBox txtCN 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtCost 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   840
         Width           =   1215
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
         Left            =   240
         TabIndex        =   24
         Top             =   1320
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
         Left            =   240
         TabIndex        =   23
         Top             =   1800
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
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   390
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CN"
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
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   270
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
         Index           =   6
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   4200
      ScaleHeight     =   3465
      ScaleWidth      =   5625
      TabIndex        =   7
      Top             =   240
      Width           =   5655
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   915
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "00.00"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   615
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "00.00"
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   555
         Left            =   2280
         TabIndex        =   8
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   720
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   11955
      TabIndex        =   6
      Top             =   9240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.TextBox txttotal1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   11880
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   11865
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   9300
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   30
      Left            =   14040
      TabIndex        =   2
      Top             =   9240
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print OR"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmincash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim count1 As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Cmdok_Click()
Dim Cn1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim Cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim ans As String

If Me.txtQuantity.Text = "" Then MsgBox "Please input Quantity", vbInformation: Exit Sub
If Me.txtName.Text = "" Then MsgBox "Please type Customer Name", vbInformation: Exit Sub

Call connect(Cn1, App.Path & "\myDB.mdb")
Call SetRs(rs1, Cn1, "SELECT * FROM InCash")
rs1.AddNew
rs1.Fields("CN") = Me.txtCN.Text
rs1.Fields("Name") = Me.txtName.Text
rs1.Fields("Item") = Me.Combo1.Text
rs1.Fields("Quantity") = Me.txtQuantity.Text
rs1.Fields("Cost") = Me.txtCost.Text
rs1.Fields("totalcost") = Me.txttotal1.Text
rs1.Fields("Date") = Date
rs1.Update
rs1.Requery

Call connect(Cn2, App.Path & "\mydb.mdb")
Call SetRs(rs2, Cn2, "Select * from Incash Where Cn='" & txtCN.Text & "' AND Date ='" & Date & "'")
Set DataGrid1.DataSource = rs2



Dim Cnss As New ADODB.Connection
Dim Rsss As New ADODB.Recordset


Call connect(Cnss, App.Path & "\myDB.mdb")
Call SetRs(Rsss, Cnss, "SELECT * FROM STocks WHERE Stocks.item= '" & Combo1.Text & "'")

Rsss.Fields("Quantity") = Me.Text2.Text
Rsss.Update


MsgBox " New  Item Has been Added", vbInformation
ans = MsgBox("Would you like to Add another Item?", vbYesNo)
If ans = vbYes Then

txtQuantity.Text = ""
txtCost.Text = ""
Combo1.Text = "Select an Item"
Exit Sub
Else
Cmdok.Enabled = False
cmdprint.Enabled = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim cnno As New ADODB.Connection
Dim rsno As New ADODB.Recordset
Dim cnno1 As New ADODB.Connection
Dim rsno1 As New ADODB.Recordset
Dim cnno2 As New ADODB.Connection
Dim rsno2 As New ADODB.Recordset

Call connect(cnno, App.Path & "\myDB.mdb")
Call SetRs(rsno, cnno, "SELECT * From InCash Where Cn='" & txtCN.Text & "'AND DATE='" & Date & "'")
Set DataGrid1.DataSource = rsno

Call connect(cnno1, App.Path & "\myDB.mdb")
Call SetRs(rsno1, cnno1, "Select sum(totalcost)as total from InCash Where CN = '" & txtCN.Text & "'") 'And Date='" & Date & "'")
With frmincash
.txttotal.Text = Format(rsno1.Fields("Total"), " #,##0.00")
End With
Call connect(cnno2, App.Path & "\myDB.mdb")
Call SetRs(rsno2, cnno2, "Select sum(totalcost)as buo from InCash Where CN = '" & txtCN.Text & "'")
With frmincash
.txttotal1.Text = Format(rsno2.Fields("buo"))
End With


Exit Sub

End If
Set Rsss = Nothing
Set Cnss = Nothing
Set rs1 = Nothing
Set Cn1 = Nothing
Set rs2 = Nothing
Set Cn2 = Nothing

Set cnno = Nothing
Set rsno = Nothing
Set cnno1 = Nothing
Set rsno1 = Nothing
Set cnno2 = Nothing
Set rsno2 = Nothing
End Sub

Private Sub cmdPrint_Click()
If txtcash.Text = "" Then MsgBox "Please Type Cash Payment First Before Printing An OR", vbInformation: Exit Sub
Dim cnp As New ADODB.Connection
Dim rsp As New ADODB.Recordset

Call connect(cnp, App.Path & "\myDb.mdb")
Call SetRs(rsp, cnp, "SELECT * from InCash Where Cn='" & txtCN.Text & "'")
Set DataReport6.DataSource = rsp
Unload Me
DataReport6.Show

End Sub

Private Sub Combo1_Change()
On Error Resume Next
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
On Error Resume Next
cmdprint.Enabled = False
lblName(5).Visible = False
txtChange.Visible = False
txtName.Text = ""
Combo1.Text = "Select an Item"
txtQuantity.Text = ""
txtCost.Text = ""


Dim CNcount As New ADODB.Connection
Dim RSCount As New ADODB.Recordset
Dim RSre As New ADODB.Recordset
Dim CNre As New ADODB.Connection



txtCN.Enabled = False
Call connect(CNcount, App.Path & "\myDB.mdb")
Call SetRs(RSCount, CNcount, "Select * from InCash")

count1 = RSCount.RecordCount + 1
    Select Case Len(count1)
        Case 1: txtCN.Text = "0000000" & RSCount.RecordCount + 1
        Case 2: txtCN.Text = "000" & RSCount.RecordCount + 1
        Case 3: txtCN.Text = "00" & RSCount.RecordCount + 1
        Case 4: txtCN.Text = "0" & RSCount.RecordCount + 1
        Case 5: txtCN.Text = RSCount.RecordCount + 1
    End Select
Set CNcount = Nothing
Set RSCount = Nothing



'cmdSave.Enabled = False
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

Private Sub Text3_Change()

End Sub

Private Sub txtcash_Change()
Me.txtChange.Text = (Val(txtcash.Text) - Val(txttotal1.Text))
If Val(txtcash.Text) >= Val(txttotal1.Text) Then
lblName(5).Visible = True
txtChange.Visible = True
Else
lblName(5).Visible = False
txtChange.Visible = False
End If

End Sub

Private Sub txtQuantity_Change()
Me.txttotal1.Text = (Val(Me.txtQuantity.Text) * Val(Me.txtCost.Text))

Text2.Text = Val(Text1.Text) - Val(txtQuantity.Text)
If (Val(Text1.Text) <= Val(txtQuantity.Text)) Then
Cmdok.Enabled = False
MsgBox " Not enough Stocks on hand" & vbCrLf & "" & vbCrLf & "Need to replenish stock", vbInformation
txtQuantity.SetFocus
Exit Sub
Else
Cmdok.Enabled = True
End If
End Sub

