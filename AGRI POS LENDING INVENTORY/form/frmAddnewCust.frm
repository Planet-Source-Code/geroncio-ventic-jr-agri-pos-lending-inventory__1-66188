VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAddnewCust 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Customer"
   ClientHeight    =   6375
   ClientLeft      =   4380
   ClientTop       =   2775
   ClientWidth     =   7620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7620
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   240
      ScaleHeight     =   5265
      ScaleWidth      =   7065
      TabIndex        =   1
      Top             =   960
      Width           =   7095
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   720
         TabIndex        =   12
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   3960
         TabIndex        =   11
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtContactNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cboGender 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   3000
         TabIndex        =   8
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox cboCivilStatus 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtCN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtbday 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55312385
         CurrentDate     =   38703
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address"
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
         Index           =   4
         Left            =   960
         TabIndex        =   20
         Top             =   1320
         Width           =   1650
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
         Left            =   960
         TabIndex        =   19
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number"
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
         Left            =   960
         TabIndex        =   18
         Top             =   1920
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
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
         Left            =   960
         TabIndex        =   17
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Left            =   960
         TabIndex        =   16
         Top             =   3360
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
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
         Left            =   960
         TabIndex        =   15
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
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
         Left            =   960
         TabIndex        =   14
         Top             =   3000
         Width           =   345
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
         Left            =   960
         TabIndex        =   13
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Add New Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmAddnewCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim count1 As String


Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

Dim Cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim reply As String

Call connect(Cn2, App.Path & "\myDb.mdb")
Call SetRs(rs2, Cn2, "SELECT * FROM CustProfile")

If txtname.Text = "" Then MsgBox "please Filled all the fields", vbInformation, "Invalid": Exit Sub
If txtAddress.Text = "" Then MsgBox "please Filled all the fields", vbInformation, "Invalid": Exit Sub
If txtAge.Text = "" Then MsgBox "please Filled all the fields", vbInformation, "Invalid": Exit Sub
If cboGender.Text = "" Then MsgBox "please Filled all the fields", vbInformation, "Invalid": Exit Sub
If cboCivilStatus.Text = "" Then MsgBox "please Filled all the fields", vbInformation, "Invalid": Exit Sub

'If Not rs2.EOF Then
  ' MsgBox " Name Already Exist", vbInformation
   ' Else
        rs2.addnew
        rs2.Fields("ContractNumber") = Me.txtcn.Text
        rs2.Fields("Name") = Me.txtname.Text
        rs2.Fields("Address") = Me.txtAddress.Text
        rs2.Fields("ContactNumber") = Me.txtContactNumber.Text
        'rs2.Fields("DateofBirth") = Me.MaskEdBox1.Text
        rs2.Fields("Dateofbirth") = txtbday.Text
        rs2.Fields("Age") = Me.txtAge.Text
        rs2.Fields("Gender") = Me.cboGender.Text
        rs2.Fields("CivilStatus") = Me.cboCivilStatus.Text
        rs2.Update
        rs2.Requery
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim cnxx As New ADODB.Connection
Dim rsxx As New ADODB.Recordset


    
'MsgBox " New Record Has been Added", vbInformation, "Saved"
'reply = MsgBox("Would you like to Add another record?", vbYesNo)
'If reply = vbYes Then
'txtName.Text = ""
'txtAddress.Text = ""
'txtContactNumber.Text = ""
'txtbday.Text = ""
'Me.txtAge.Text = ""
'Me.cboGender.Text = ""
'Me.cboCivilStatus.Text = ""
'Call End1
'txtName.SetFocus
'Else
'Unload Me
 '   End If
Set Cn2 = Nothing
Set rs2 = Nothing

frmOrder.txtcn.Text = Me.txtcn.Text
frmOrder.txtname.Text = txtname.Text
frmOrder.Show vbModal

Unload Me
End Sub

Private Sub DTPicker1_Change()
Dim a
txtbday.Text = DTPicker1.Value

a = txtbday.Text
txtAge.Text = DateDiff("yyyy", a, Now)
End Sub

Private Sub DTPicker1_Click()
On Error Resume Next
Dim a
txtbday.Text = DTPicker1.Value

a = txtbday.Text
txtAge.Text = DateDiff("yyyy", a, Now)

End Sub

Private Sub DTPicker1_LostFocus()
On Error Resume Next
Dim a
txtbday.Text = DTPicker1.Value

a = txtbday.Text
txtAge.Text = DateDiff("yyyy", a, Now)

End Sub

Private Sub Form_Load()
txtAddress.Text = ""
txtbday.Enabled = True
Me.Appearance = 0

With cboGender
.Clear
.AddItem "Male"
.AddItem "Female"
End With

With cboCivilStatus
.Clear
.AddItem "Single"
.AddItem "Married"
.AddItem "Widowed"
.AddItem "Separated"
End With

Dim CNcount As New ADODB.Connection
Dim RSCount As New ADODB.Recordset

Call connect(CNcount, App.Path & "\myDB.mdb")
Call SetRs(RSCount, CNcount, "Select * from CustProfile")

count1 = RSCount.RecordCount + 1
    Select Case Len(count1)
        Case 1: txtcn.Text = "2006" & RSCount.RecordCount + 1
        Case 2: txtcn.Text = "000" & RSCount.RecordCount + 1
        Case 3: txtcn.Text = "00" & RSCount.RecordCount + 1
        Case 4: txtcn.Text = "0" & RSCount.RecordCount + 1
        Case 5: txtcn.Text = RSCount.RecordCount + 1
    End Select
Set CNcount = Nothing
Set RSCount = Nothing





End Sub



Private Sub txtAddress_LostFocus()
txtAddress.Text = StrConv(txtAddress, vbProperCase)
End Sub

Private Sub txtbday_Change()
On Error Resume Next
txtAge.Text = DateDiff("yyyy", Me.txtbday, Now)
End Sub

Private Sub txtContactNumber_KeyPress(KeyAscii As Integer)
If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyInsert Then
KeyAscii = 0
End If
End Sub

Private Sub txtName_LostFocus()
txtname.Text = StrConv(txtname, vbProperCase)

End Sub

Public Sub End1()
Dim cncount1 As New ADODB.Connection
Dim rscount1 As New ADODB.Recordset
Call connect(cncount1, App.Path & "\myDB.mdb")
Call SetRs(rscount1, cncount1, "Select * from CustProfile")

count1 = rscount1.RecordCount + 1
    Select Case Len(count1)
        Case 1: txtcn.Text = "2006" & rscount1.RecordCount + 1
        Case 2: txtcn.Text = "000" & rscount1.RecordCount + 1
        Case 3: txtcn.Text = "00" & rscount1.RecordCount + 1
        Case 4: txtcn.Text = "0" & rscount1.RecordCount + 1
        Case 5: txtcn.Text = rscount1.RecordCount + 1
    End Select
Set cncount1 = Nothing
Set rscount1 = Nothing
End Sub
