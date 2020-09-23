VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "INVENTORY and POS"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   -2175
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4C93
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":EEA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13703
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6840
      Left            =   0
      ScaleHeight     =   6840
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   975
      Width           =   300
   End
   Begin VB.PictureBox Picture7 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      DrawMode        =   10  'Mask Pen
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   11880
      TabIndex        =   2
      Top             =   0
      Width           =   11880
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   360
         TabIndex        =   6
         Top             =   120
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1429
         ButtonWidth     =   1455
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "User Setting's"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add New User"
                     Text            =   "Add New User"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Change UserName or Password"
                     Text            =   "Change UserName or Password"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "BackUp Database"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add New Item's"
                     Text            =   "Add New Item's"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Replenish Stocks"
                     Text            =   "Replenish Stocks"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "View Inventory"
                     Text            =   "View Inventory Master File"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AddNew StocksCategory"
                     Text            =   "AddNew StocksCategory"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   7440
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   6960
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   7440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   6960
         Width           =   150
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7815
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Picture         =   "MDIForm1.frx":186E0
            Text            =   "U S E R :"
            TextSave        =   "U S E R :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Enabled         =   0   'False
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            TextSave        =   "8/4/06"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            TextSave        =   "9:51 PM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6840
      Left            =   11580
      ScaleHeight     =   6840
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   975
      Width           =   300
   End
   Begin VB.Menu MnuSetup 
      Caption         =   "Setup"
      Begin VB.Menu mnuUser 
         Caption         =   "User Setting's"
         Begin VB.Menu mnuAdduser 
            Caption         =   "Add New User"
         End
         Begin VB.Menu mnua 
            Caption         =   "-"
         End
         Begin VB.Menu mnuChange 
            Caption         =   "Change UserName or Password"
         End
      End
      Begin VB.Menu mnuD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilities 
         Caption         =   "DataBase Utility"
         Begin VB.Menu mnubackup 
            Caption         =   "BackUp Database"
         End
      End
      Begin VB.Menu mnuc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnuinventory 
         Caption         =   "Inventories"
         Begin VB.Menu mnuaddnewcateg 
            Caption         =   "Add New Product Catagory"
         End
         Begin VB.Menu mnujo 
            Caption         =   "-"
         End
         Begin VB.Menu mnunewitem 
            Caption         =   "New Item's"
         End
         Begin VB.Menu mnuaaa 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnunewprice 
            Caption         =   "Set New Price"
         End
         Begin VB.Menu mnumark69 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReplenishequipment 
            Caption         =   "Replenish Stocks"
         End
         Begin VB.Menu mnuxx 
            Caption         =   "-"
         End
         Begin VB.Menu mnuviewequipmentinventory 
            Caption         =   "View  Inventory"
         End
      End
      Begin VB.Menu mnuaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewCustProfile 
         Caption         =   "Add New Customer Profile"
      End
      Begin VB.Menu mnuaddnewdebt 
         Caption         =   "Add New Debt Items"
         Begin VB.Menu mnuContractNumber 
            Caption         =   "By Customer Contract Number"
         End
         Begin VB.Menu mnuabcd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCustomerName 
            Caption         =   "By Customer Name"
         End
      End
      Begin VB.Menu mnuPaymenyfordebt 
         Caption         =   "Payment For Debt Item's"
         Begin VB.Menu mnupaymentBCCN 
            Caption         =   "By Customer Contract Number"
         End
         Begin VB.Menu mnuabcde 
            Caption         =   "-"
         End
         Begin VB.Menu mnupaymentBCN 
            Caption         =   "By Customer Name"
         End
      End
      Begin VB.Menu mnuj 
         Caption         =   "-"
      End
      Begin VB.Menu mnucash 
         Caption         =   "Cash Loan"
         Begin VB.Menu mnucashBCCN 
            Caption         =   "By Customer Contract Number"
         End
         Begin VB.Menu mnuxtian1 
            Caption         =   "-"
         End
         Begin VB.Menu mnucashBCN 
            Caption         =   "By Customer Name"
         End
      End
      Begin VB.Menu mnujojo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuincash 
         Caption         =   "Sell In Cash"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "Report"
      Begin VB.Menu mnuSIIC 
         Caption         =   "Sold Item in Cash by Date"
      End
      Begin VB.Menu mnulito 
         Caption         =   "-"
      End
      Begin VB.Menu mnudebtitem 
         Caption         =   "Debt Item"
      End
      Begin VB.Menu mnuFullyPaid 
         Caption         =   "Fully Paid Item's"
      End
      Begin VB.Menu mnujojo2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelinquentAccount 
         Caption         =   "Delinquent Account"
         Begin VB.Menu mnunamerar 
            Caption         =   "Name"
         End
         Begin VB.Menu mnuCnnum 
            Caption         =   "Contract Number"
         End
      End
      Begin VB.Menu mnuCWRB 
         Caption         =   "Customer's With Remaining Balance"
         Begin VB.Menu mnunname 
            Caption         =   "Name"
         End
         Begin VB.Menu mnucnwb 
            Caption         =   "Contract Number"
         End
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
MDIForm1.Text1.Text = frmSplash.Text1.Text
MDIForm1.Text2.Text = frmSplash.Text2.Text
'MDIForm1.Text1.Text = frmSplash.Text1.Text
'MDIForm1.Text2.Text = frmSplash.Text2.Text
With StatusBar1
        .Panels(2).Text = " " & UserName
        '.Panels(1).Text = "Grading System"
        .Panels(3).Text = "" & password
        .Panels(3).Visible = False
    End With
    

End Sub

Private Sub MDIForm_Resize()
MDIForm1.WindowState = 2
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub mnuabout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuaddnewcateg_Click()
frmnewcategory.Show vbModal
End Sub

'Private Sub mnuaddnewdebt_Click()
'Dim cndebt As New ADODB.Connection
'Dim rsdebt As New ADODB.Recordset
'Dim reply As String
'reply = ""
'reply = InputBox("Enter Contract Number", "Add New Debt Record")
'If reply = "" Then
'Exit Sub

'Else
'Call connect(cndebt, App.Path & "\myDB.mdb")
'Call SetRs(rsdebt, cndebt, "Select * from custprofile Where contractnumber = '" & reply & "'")

'If Not rsdebt.EOF Then
'frmOrder.txtCN.Text = rsdebt.Fields("contractNumber")
'frmOrder.txtName.Text = rsdebt.Fields("Name")
'frmOrder.Show vbModal
'Else

'MsgBox "Contract Number Don't Exist", vbInformation
'Exit Sub

'End If


'End If
'End Sub

Private Sub mnuAdduser_Click()
frmAddnewUser.Show vbModal
End Sub

Private Sub mnubackup_Click()
frmBackupDB.Show vbModal
End Sub

Private Sub mnucashBCCN_Click()
Dim cndebt As New ADODB.Connection
Dim rsdebt As New ADODB.Recordset
Dim reply As String
reply = ""
reply = InputBox("Enter Contract Number", "Add Cash Loan Record")
If reply = "" Then
Exit Sub

Else
Call connect(cndebt, App.path & "\myDB.mdb")
Call SetRs(rsdebt, cndebt, "Select * from custprofile Where contractnumber = '" & reply & "'")

If Not rsdebt.EOF Then
frmCashloan.txtcn.Text = rsdebt.Fields("contractNumber")
frmCashloan.txtname.Text = rsdebt.Fields("Name")
frmCashloan.Show vbModal
Else

MsgBox "Contract Number Don't Exist", vbInformation
Exit Sub

End If


End If
Set cndebt = Nothing
Set rsdebt = Nothing
End Sub

Private Sub mnucashBCN_Click()
Dim cndebt As New ADODB.Connection
Dim rsdebt As New ADODB.Recordset
Dim reply As String
reply = ""
reply = InputBox("Enter Custome Name", "Add Cash Loan Record")
If reply = "" Then
Exit Sub

Else
Call connect(cndebt, App.path & "\myDB.mdb")
Call SetRs(rsdebt, cndebt, "Select * from custprofile Where Name = '" & reply & "'")

If Not rsdebt.EOF Then
frmCashloan.txtcn.Text = rsdebt.Fields("contractNumber")
frmCashloan.txtname.Text = rsdebt.Fields("Name")
frmCashloan.Show vbModal
Else

MsgBox "Customer Name Don't Exist", vbInformation
Exit Sub

End If


End If
Set cndebt = Nothing
Set rsdebt = Nothing
End Sub

Private Sub mnuChange_Click()
frmchangeuser.Show vbModal
End Sub

Private Sub mnuCnnum_Click()
On Error Resume Next
 Dim reply As String
 Dim cne As New ADODB.Connection
 Dim RSe As New ADODB.Recordset
 Dim cnp As New ADODB.Connection
 Dim rsp As New ADODB.Recordset
        
        reply = ""
        reply = InputBox("Enter Contract Number", "Search")
        
If reply = "" Then
Exit Sub

Else
         
        Call connect(cne, App.path & "\myDB.mdb")
        Call SetRs(RSe, cne, "SELECT * From Debtrecord Where ContractNumber= '" & reply & "'")
        
        If RSe.EOF Then
             MsgBox "No record found", vbInformation
             Exit Sub
           
        Else
            
            With frmDelinquent
             .Label1.Caption = RSe.Fields("ContractNumber")
             .Label2.Caption = RSe.Fields!Name
            
             End With
            Set frmDelinquent.DataGrid1.DataSource = RSe
           With frmDelinquent.DataGrid1
           .Columns(3).Visible = False
           .Columns(4).Visible = False
           End With
            frmDelinquent.Show vbModal
        
        End If
         
        
End If

  
    Set cne = Nothing
    Set RSe = Nothing

End Sub

Private Sub mnucnwb_Click()
Dim cncategory5 As New ADODB.Connection
Dim rscategory5 As New ADODB.Recordset
Dim reply As String

reply = ""
        reply = InputBox("Enter Contract Number", "Search")
        
If reply = "" Then
Exit Sub

Else

    Call connect(cncategory5, App.path & "\myDB.mdb")
    Call SetRs(rscategory5, cncategory5, "SELECT * FROM debtrecord Where ContractNumber='" & reply & "'")
    
    If rscategory5.EOF Then
             MsgBox "No record found", vbInformation
             Exit Sub
           
        Else
            
            With frmcustomerwbalance
             .Label1.Caption = rscategory5.Fields("ContractNumber")
             .Label2.Caption = rscategory5.Fields!Name
            
             End With
            Set frmcustomerwbalance.DataGrid1.DataSource = rscategory5
         
            frmcustomerwbalance.Show vbModal
        
        End If
         
        
End If
        
      
    
        Set rscategory5 = Nothing
        Set cncategory5 = Nothing
  


End Sub

Private Sub mnuContractNumber_Click()
Dim cndebt As New ADODB.Connection
Dim rsdebt As New ADODB.Recordset
Dim reply As String
reply = ""
reply = InputBox("Enter Contract Number", "Add New Debt Record")
If reply = "" Then
Exit Sub

Else
Call connect(cndebt, App.path & "\myDB.mdb")
Call SetRs(rsdebt, cndebt, "Select * from custprofile Where contractnumber = '" & reply & "'")

If Not rsdebt.EOF Then
frmOrder.txtcn.Text = rsdebt.Fields("contractNumber")
frmOrder.txtname.Text = rsdebt.Fields("Name")
frmOrder.Show vbModal
Else

MsgBox "Contract Number Don't Exist", vbInformation
Exit Sub

End If


End If
Set cndebt = Nothing
Set rsdebt = Nothing
End Sub

Private Sub mnuCustomerName_Click()
Dim cndebt As New ADODB.Connection
Dim rsdebt As New ADODB.Recordset
Dim reply As String
reply = ""
reply = InputBox("Enter Custome Name", "Add New Debt Record")
If reply = "" Then
Exit Sub

Else
Call connect(cndebt, App.path & "\myDB.mdb")
Call SetRs(rsdebt, cndebt, "Select * from custprofile Where Name = '" & reply & "'")

If Not rsdebt.EOF Then
frmOrder.txtcn.Text = rsdebt.Fields("contractNumber")
frmOrder.txtname.Text = rsdebt.Fields("Name")
frmOrder.Show vbModal
Else

MsgBox "Customer Name Don't Exist", vbInformation
Exit Sub

End If


End If
Set cndebt = Nothing
Set rsdebt = Nothing
End Sub

'Private Sub mnuCWRB_Click()

'End Sub

Private Sub mnudebtitem_Click()
frmdebtitem.Show vbModal
End Sub

Private Sub mnuexit_Click()
Dim ans
ans = MsgBox("Do you really want to exit?", vbYesNo + vbInformation)
If ans = vbYes Then
End
Else
Exit Sub
End If
End Sub

Private Sub mnuRepai_Click()
frmCompactDB.Show vbModal
End Sub

Private Sub mnuRestore_Click()
frmRestoreDB.Show vbModal
End Sub

Private Sub Picture4_Click()
frmBackupDB.Show vbModal
End Sub

Private Sub mnuFullyPaid_Click()
frmFullypaid.Show vbModal
End Sub

Private Sub mnuincash_Click()
frmincash.Show vbModal

End Sub

Private Sub mnunamerar_Click()
On Error Resume Next

 Dim reply As String
        reply = ""
        reply = InputBox("Enter Name", "Search")
        
        If reply = "" Then
            Exit Sub
'
Else

        Call connect(cncategory3, App.path & "\myDB.mdb")
        Call SetRs(rscategory3, cncategory3, "SELECT * From Debtrecord Where name= '" & reply & "'")
        rscategory3.Requery
        
        If rscategory3.EOF Then
             MsgBox "No record found", vbInformation
             Exit Sub
           
        Else
            
            With frmDelinquent
             .Label1.Caption = rscategory3.Fields("ContractNumber")
             .Label2.Caption = rscategory3.Fields!Name
            
             End With
            Set frmDelinquent.DataGrid1.DataSource = rscategory3
           With frmDelinquent.DataGrid1
           .Columns(3).Visible = False
           .Columns(4).Visible = False
           End With
            frmDelinquent.Show vbModal
        End If
        
End If
 
    Set RSf = Nothing
    Set CNf = Nothing
    Set cncategory3 = Nothing
    Set rscategory = Nothing
    


End Sub

Private Sub mnuNewCustProfile_Click()
frmAddnewCust.Show vbModal
End Sub

Private Sub mnunewitem_Click()
frmAddstock.Show vbModal
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnunewprice_Click()
frmeditprice.Show vbModal
End Sub

Private Sub mnunname_Click()
Dim cncategory50 As New ADODB.Connection
Dim rscategory50 As New ADODB.Recordset
Dim reply As String

reply = ""
        reply = InputBox("Enter Customer Name", "Search")
        
If reply = "" Then
Exit Sub

Else

    Call connect(cncategory50, App.path & "\myDB.mdb")
    Call SetRs(rscategory50, cncategory50, "SELECT * FROM debtrecord Where Name='" & reply & "'")
    
    If rscategory50.EOF Then
             MsgBox "No record found", vbInformation
             Exit Sub
           
        Else
            
            With frmcustomerwbalance
             .Label1.Caption = rscategory50.Fields("ContractNumber")
             .Label2.Caption = rscategory50.Fields!Name
            
             End With
            Set frmcustomerwbalance.DataGrid1.DataSource = rscategory50
         
            frmcustomerwbalance.Show vbModal
        
        End If
         
        
End If
        
      
    
        Set rscategory5 = Nothing
        Set cncategory5 = Nothing
  


End Sub



Private Sub mnupaymentBCCN_Click()
Dim cndebt As New ADODB.Connection
Dim rsdebt As New ADODB.Recordset
Dim cndebt1 As New ADODB.Connection
Dim rsdebt1 As New ADODB.Recordset
Dim cndebt2 As New ADODB.Connection
Dim rsdebt2 As New ADODB.Recordset
Dim cncust As New ADODB.Connection
Dim rscust As New ADODB.Recordset

        Dim reply As String
        reply = ""
        reply = InputBox("Enter Contract Number", "Search")
        
        If reply = "" Then
            Exit Sub
        
        Else
            Call connect(cndebt, App.path & "\myDB.mdb")
            Call SetRs(rsdebt, cndebt, "Select * from Debtrecord Where contractnumber = '" & reply & "'")
            Set frmpayment.DataGrid1.DataSource = rsdebt
            
            
            Call connect(cndebt1, App.path & "\myDB.mdb")
            Call SetRs(rsdebt1, cndebt1, "Select sum(totalcost)as total from Debtrecord Where contractnumber = '" & reply & "'")
            With frmpayment
            .txttotaldebt.Text = Format(rsdebt1.Fields("Total"), " #,##0.00")
            End With
            
            Call connect(cndebt2, App.path & "\myDB.mdb")
            Call SetRs(rsdebt2, cndebt2, "Select sum(totalcost)as buo from Debtrecord Where contractnumber = '" & reply & "'")
            With frmpayment
            .Text1.Text = Format(rsdebt2.Fields("buo"))
            End With
        End If
        
        If rsdebt.EOF Then
            MsgBox "No Record Found" & vbCrLf & "Please Check it on the payment for item remaining balance", vbInformation
            Exit Sub
        Else
            Call connect(cncust, App.path & "\myDB.mdb")
            Call SetRs(rscust, cncust, "Select * from custprofile where contractnumber= '" & reply & "'")
            
            frmpayment.Text6.Text = rscust.Fields("Name")
            frmpayment.Text7.Text = rscust.Fields("address")
            frmpayment.Text8.Text = Date
            frmpayment.Text5.Text = reply
            frmpayment.Show vbModal
            
        End If
        
    Set cndebt = Nothing
    Set rsdebt = Nothing
    Set cndebt1 = Nothing
    Set rsdebt1 = Nothing
    Set cndebt2 = Nothing
    Set rsdebt2 = Nothing
    Set cncust = Nothing
    Set rscust = Nothing

End Sub

Private Sub mnupaymentBCN_Click()
Dim cndebt As New ADODB.Connection
Dim rsdebt As New ADODB.Recordset
Dim cndebt1 As New ADODB.Connection
Dim rsdebt1 As New ADODB.Recordset
Dim cndebt2 As New ADODB.Connection
Dim rsdebt2 As New ADODB.Recordset
Dim cncust As New ADODB.Connection
Dim rscust As New ADODB.Recordset

Dim reply As String
reply = ""
reply = InputBox("Enter Customer Name", "Search")

If reply = "" Then
    Exit Sub

Else
    Call connect(cndebt, App.path & "\myDB.mdb")
    Call SetRs(rsdebt, cndebt, "Select * from Debtrecord Where Name = '" & reply & "'")
    Set frmpayment.DataGrid1.DataSource = rsdebt
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call connect(cndebt1, App.path & "\myDB.mdb")
    Call SetRs(rsdebt1, cndebt1, "Select sum(totalcost)as total from Debtrecord Where Name = '" & reply & "'")
        With frmpayment
            .txttotaldebt.Text = Format(rsdebt1.Fields("Total"), " #,##0.00")
        End With
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call connect(cndebt2, App.path & "\myDB.mdb")
    Call SetRs(rsdebt2, cndebt2, "Select sum(totalcost)as buo from Debtrecord Where Name = '" & reply & "'")
        With frmpayment
            .Text1.Text = Format(rsdebt2.Fields("buo"))
        End With
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

            If rsdebt.EOF Then
                MsgBox "No Record Found", vbInformation
                Exit Sub
            Else
                Call connect(cncust, App.path & "\myDB.mdb")
                Call SetRs(rscust, cncust, "Select * from custprofile where Name='" & reply & "'")
                
                frmpayment.Text6.Text = rscust.Fields("Name")
                frmpayment.Text7.Text = rscust.Fields("address")
                frmpayment.Text8.Text = Date
                frmpayment.Text5.Text = rscust.Fields("ContractNumber")
                
                
                frmpayment.Label2.Caption = rscust.Fields("ContractNumber")
                frmpayment.Label3.Caption = rscust.Fields("Name")
                 frmpayment.Label4.Caption = Date
                frmpayment.Show vbModal
            
            End If
    
    
    Set cndebt = Nothing
    Set rsdebt = Nothing
    Set cndebt1 = Nothing
    Set rsdebt1 = Nothing
    Set cndebt2 = Nothing
    Set rsdebt2 = Nothing
    Set cncust = Nothing
    Set rscust = Nothing
End Sub


Private Sub mnuPFBBCCN_Click()





End Sub

Private Sub mnuReplenishequipment_Click()
frmReplenish.Show vbModal
End Sub

Private Sub mnuSIIC_Click()
frmSolditemincash.Show vbModal
End Sub

Private Sub mnuviewequipmentinventory_Click()
frmviewinventory.Show vbModal
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

Select Case ButtonMenu.Key
Case "Add New User":
mnuAdduser_Click
Case "Change UserName or Password":
mnuChange_Click
Case "Add New Item's"
mnunewitem_Click
Case "Replenish Stocks"
mnuReplenishequipment_Click
Case "View Inventory"
mnuviewequipmentinventory_Click
Case "AddNew StocksCategory"
mnuaddnewcateg_Click
End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 2:
        mnubackup_Click
   
    Case 4:
        mnuNewCustProfile_Click
  
        
    End Select
    
End Sub

Public Sub Developer1()
retvalue = GetSetting("A", "0", "Runcount")
Worm = Val(retvalue) + 1
SaveSetting "A", "0", "RunCount", Worm

If Worm > 600 Then 'put one number lower then it says....you can only run the program 200 times.
    MsgBox "This is the End of the trial run" & vbCrLf & "This Is only  Good for 500 Runs", 16, "Sorry"
    MsgBox "Email the developer of this system at" & vbCrLf & "venticjojo@yahoo.com" & vbCrLf & "Venticjojo05@hotmail.com" & vbCrLf & "or call 09186070112" & vbCrLf & "For the reactivation or full version of this system", 6, "System Developer"
    Unload Me
End If
MDIForm1.Toolbar1.Enabled = False
MDIForm1.mnuaddnewcateg.Enabled = False
MDIForm1.mnuaddnewcateg.Visible = False

End Sub
