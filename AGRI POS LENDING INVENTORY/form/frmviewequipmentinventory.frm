VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmviewinventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Master File"
   ClientHeight    =   8070
   ClientLeft      =   1995
   ClientTop       =   2355
   ClientWidth     =   11685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11685
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   8505
      TabIndex        =   5
      Top             =   240
      Width           =   8535
      Begin VB.ComboBox cbocategory 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Filter:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6585
      ScaleWidth      =   8505
      TabIndex        =   3
      Top             =   1080
      Width           =   8535
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Stock ID"
            Object.Width           =   1500
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item"
            Object.Width           =   5290
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Unit"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Category"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Print"
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Replenish"
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewequipmentinventory.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmviewequipmentinventory.frx":08DA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmviewinventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub cbocategory_Change()
Dim cncategory2 As New ADODB.Connection
Dim rscategory2 As New ADODB.Recordset

Dim cncategory3 As New ADODB.Connection
Dim rscategory3 As New ADODB.Recordset

If cbocategory.Text = "All Categories" Then

Call connect(cncategory3, App.path & "\myDB.mdb")
Call SetRs(rscategory3, cncategory3, "SELECT * FROM Stocks order by StockID ASC")

ListView1.ListItems.Clear
    
        With rscategory3
        While Not .EOF
        
            Set x = ListView1.ListItems.Add(, , .Fields("StockID"), , 1)
                x.SubItems(1) = .Fields("Item")
                x.SubItems(2) = .Fields("Unit")
                x.SubItems(3) = .Fields("Quantity")
                x.SubItems(4) = .Fields("Category")
                x.SubItems(5) = .Fields("UnitCost")
            .MoveNext
        Wend
    End With
    
  

    Set rscategory3 = Nothing
    Set cncategory3 = Nothing


Else

Call connect(cncategory2, App.path & "\myDB.mdb")
Call SetRs(rscategory2, cncategory2, "SELECT * FROM Stocks WHERE Category='" & cbocategory.Text & "'")

ListView1.ListItems.Clear

    With rscategory2
        While Not .EOF
            Set x = ListView1.ListItems.Add(, , .Fields("StockID"), , 2)
                x.SubItems(1) = .Fields("Item")
                x.SubItems(2) = .Fields("Unit")
                x.SubItems(3) = .Fields("Quantity")
                x.SubItems(4) = .Fields("Category")
                 x.SubItems(5) = .Fields("UnitCost")
            .MoveNext
        Wend
    End With
    Set rscategory2 = Nothing
    Set cncategory2 = Nothing
End If


End Sub

Private Sub cbocategory_Click()
Dim cncategory2 As New ADODB.Connection
Dim rscategory2 As New ADODB.Recordset

Dim cncategory3 As New ADODB.Connection
Dim rscategory3 As New ADODB.Recordset

If cbocategory.Text = "All Categories" Then

Call connect(cncategory3, App.path & "\myDB.mdb")
Call SetRs(rscategory3, cncategory3, "SELECT * FROM Stocks order by StockID ASC")
ListView1.ListItems.Clear
    
        With rscategory3
        While Not .EOF
            Set x = ListView1.ListItems.Add(, , .Fields("StockID"), , 1)
                x.SubItems(1) = .Fields("Item")
                x.SubItems(2) = .Fields("Unit")
                x.SubItems(3) = .Fields("Quantity")
                x.SubItems(4) = .Fields("Category")
                 x.SubItems(5) = .Fields("UnitCost")
            .MoveNext
        Wend
    End With
    
   

    Set rscategory3 = Nothing
    Set cncategory3 = Nothing


Else

Call connect(cncategory2, App.path & "\myDB.mdb")
Call SetRs(rscategory2, cncategory2, "SELECT * FROM Stocks WHERE Category='" & cbocategory.Text & "'")

ListView1.ListItems.Clear

    With rscategory2
        While Not .EOF
            Set x = ListView1.ListItems.Add(, , .Fields("StockID"), , 2)
                x.SubItems(1) = .Fields("Item")
                x.SubItems(2) = .Fields("Unit")
                x.SubItems(3) = .Fields("Quantity")
                x.SubItems(4) = .Fields("Category")
                 x.SubItems(5) = .Fields("UnitCost")
            .MoveNext
        Wend
    End With
    Set rscategory2 = Nothing
    Set cncategory2 = Nothing
End If

End Sub

Private Sub Command2_Click()
Dim CNprint As New ADODB.Connection
Dim RSprint As New ADODB.Recordset

Dim CNprint1 As New ADODB.Connection
Dim RSprint1 As New ADODB.Recordset

If cbocategory.Text = "All Categories" Then

    Call connect(CNprint, App.path & "\myDB.mdb")
    Call SetRs(RSprint, CNprint, "SELECT * FROM stocks Order by stockID asc")
    Set DataReport3.DataSource = RSprint
    Unload Me
    DataReport3.Show
    Set CNprint = Nothing
    Set RSprint = Nothing
    

Else

    Call connect(CNprint1, App.path & "\myDB.mdb")
    Call SetRs(RSprint1, CNprint1, "SELECT * FROM Stocks WHERE Category='" & cbocategory & "'")
    Set DataReport3.DataSource = RSprint1
    Unload Me
    DataReport3.Show
    Set CNprint1 = Nothing
    Set RSprint1 = Nothing

End If


End Sub

Private Sub Command3_Click()
If frmReplenish.Combo1.Text = "" Then
    MsgBox "Select Item to Replenish", vbInformation
    Exit Sub
Else
    Unload Me
    frmReplenish.Show vbModal
End If
End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Form_Load()
ListView1.BackColor = &HFDECD7
Label1.BackColor = &HFDECD7
Picture2.BackColor = &HFDECD7
cbocategory.BackColor = &HFFFFFF

Dim cncategory1 As New ADODB.Connection
Dim rscategory1 As New ADODB.Recordset

Call connect(cncategory1, App.path & "\MyDB.mdb")
Call SetRs(rscategory1, cncategory1, "SELECT * FROM stockscategory  order by stockscategory ASC")
With rscategory1
    While Not .EOF
        cbocategory.AddItem .Fields("StocksCategory")
        .MoveNext
        Wend
    End With
Set cncategory1 = Nothing
Set rscategory1 = Nothing
End Sub

Private Sub ListView1_Click()
On Error Resume Next

With frmReplenish
.Combo1.Text = ListView1.SelectedItem
.Text1.Text = ListView1.SelectedItem.SubItems(1)
.Text2.Text = ListView1.SelectedItem.SubItems(4)
.Text4.Text = ListView1.SelectedItem.SubItems(2)
.Text5.Text = ListView1.SelectedItem.SubItems(3)
End With

End Sub

