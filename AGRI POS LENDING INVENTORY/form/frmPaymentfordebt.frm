VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPaymentforbalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment for Balance"
   ClientHeight    =   7065
   ClientLeft      =   2385
   ClientTop       =   2160
   ClientWidth     =   10740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10740
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   7080
      TabIndex        =   31
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   15300
      TabIndex        =   30
      Top             =   10200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   15240
      TabIndex        =   29
      Top             =   10320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1440
      TabIndex        =   28
      Top             =   6360
      Width           =   2175
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   5400
      ScaleHeight     =   5505
      ScaleWidth      =   5145
      TabIndex        =   26
      Top             =   480
      Width           =   5175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5055
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "ContractNumber"
            Caption         =   "CntrctNmbr"
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
            DataField       =   "Name"
            Caption         =   "Name"
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
         BeginProperty Column02 
            DataField       =   "totalamounttopay"
            Caption         =   "Amount To Pay"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Php""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "AmountPaid"
            Caption         =   "Amount Paid"
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
         BeginProperty Column04 
            DataField       =   "Balance"
            Caption         =   "Balance"
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
         BeginProperty Column05 
            DataField       =   "DatePaid"
            Caption         =   "Date"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5505
      ScaleWidth      =   4905
      TabIndex        =   5
      Top             =   480
      Width           =   4935
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3720
         Width           =   2295
      End
      Begin VB.TextBox txtpayment 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   18
         Top             =   3000
         Width           =   2295
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         ScaleHeight     =   465
         ScaleWidth      =   2265
         TabIndex        =   15
         Top             =   2400
         Width           =   2295
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0"
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   7
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   425
         Left            =   2160
         ScaleHeight     =   390
         ScaleWidth      =   2265
         TabIndex        =   12
         Top             =   1800
         Width           =   2295
         Begin VB.TextBox txttotaldebttopay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0"
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   5
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   425
         Left            =   2160
         ScaleHeight     =   390
         ScaleWidth      =   2385
         TabIndex        =   9
         Top             =   720
         Width           =   2415
         Begin VB.TextBox txttotaldebt 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Php""#,##0.00;(""Php""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13321
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Php"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   4
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   380
         Left            =   2160
         ScaleHeight     =   345
         ScaleWidth      =   945
         TabIndex        =   6
         Top             =   1320
         Width           =   975
         Begin VB.TextBox txtpercent 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   2
            Left            =   360
            TabIndex        =   8
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
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
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   9
         Left            =   720
         TabIndex        =   25
         Top             =   3720
         Width           =   1110
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Payment"
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
         Index           =   8
         Left            =   1200
         TabIndex        =   24
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Total Amount to Pay"
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
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Interest Amount"
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
         Left            =   480
         TabIndex        =   22
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Interest "
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
         Left            =   1200
         TabIndex        =   21
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Total Balance:"
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
         Left            =   600
         TabIndex        =   20
         Top             =   840
         Width           =   1260
      End
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   15240
      TabIndex        =   4
      Top             =   10200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   15225
      TabIndex        =   3
      Top             =   10200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   6360
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   15225
      TabIndex        =   1
      Top             =   10200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   15225
      TabIndex        =   0
      Top             =   10200
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "frmPaymentforbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
On Error Resume Next
cmdprint.Enabled = True
Dim RSpd As New ADODB.Recordset
Dim CNpd As New ADODB.Connection

Dim RSpd1 As New ADODB.Recordset
Dim CNpd1 As New ADODB.Connection

Dim RSpd3 As New ADODB.Recordset
Dim CNpd3 As New ADODB.Connection

Dim RSpd4 As New ADODB.Recordset
Dim CNpd4 As New ADODB.Connection


Dim sagot  As String


Call connect(CNpd3, App.Path & "\myDB.mdb")
Call SetRs(RSpd3, CNpd3, "Delete * from PartialPaid Where contractNumber like'" & Text5.Text & "'")
RSpd3.delete
 If RSpd3.EOF Then
            RSpd3.MoveLast
        End If
Set DataGrid1.DataSource = RSpd3


If Me.txtpayment.Text >= Me.Text2.Text Then
Call connect(CNpd, App.Path & "\myDB.mdb")
Call SetRs(RSpd, CNpd, "SELECT * from FullyPaid")
RSpd.AddNew
RSpd.Fields("Contractnumber") = Me.Text5.Text
RSpd.Fields("Name") = Text6.Text
RSpd.Fields("PercentInterest") = txtpercent.Text
RSpd.Fields("InterestAmount") = Me.txttotaldebttopay.Text
RSpd.Fields("AmountPaid") = txtpayment.Text
RSpd.Fields("DatePaid") = Date
RSpd.Update
RSpd.Requery



Else
Call connect(CNpd1, App.Path & "\myDB.mdb")
Call SetRs(RSpd1, CNpd1, "SELECT * from PartialPaid")
'Set DataGrid1.DataSource = RSpd1
RSpd1.AddNew
RSpd1.Fields("Contractnumber") = Me.Text5.Text
RSpd1.Fields("Name") = Text6.Text
RSpd1.Fields("totalamounttopay") = Text2.Text
'RSpd1.Fields("PercentInterest") = txtpercent.Text
RSpd1.Fields("Balance") = Me.Text4.Text
RSpd1.Fields("AmountPaid") = txtpayment.Text
RSpd1.Fields("DatePaid") = Date
RSpd1.Update
RSpd1.Requery
End If



'Me.txttotaldebt.Text = ""
'Me.txtpercent.Text = ""
'Me.txttotaldebttopay.Text = ""
'Me.Text2.Text = ""
'Me.txtpayment.Text = ""
'Me.Text4.Text = ""
Me.cmdok.Enabled = False



Set CNpd = Nothing
Set RSpd = Nothing

Set CNpd1 = Nothing
Set RSpd1 = Nothing

Set CNpd2 = Nothing
Set RSpd2 = Nothing

Set CNpd3 = Nothing
Set RSpd3 = Nothing


End Sub

Private Sub cmdPrint_Click()
Dim cnPnt As New ADODB.Connection
Dim rsPnt As New ADODB.Recordset

Dim cnPnt1 As New ADODB.Connection
Dim rsPnt1 As New ADODB.Recordset


If Me.txtpayment.Text >= Me.Text2.Text Then
Call connect(cnPnt, App.Path & "\myDB.mdb")
Call SetRs(rsPnt, cnPnt, "SELECT * from FullyPaid Where fullypaid.contractnumber ='" & Text5.Text & "' AND DatePaid ='" & Text8.Text & "'")

Set DataReport4.DataSource = rsPnt
Unload Me
DataReport4.Show

'Me.txttotaldebt.Text = ""
'Me.txtpercent.Text = ""
'Me.txttotaldebttopay.Text = ""
'Me.Text2.Text = ""
'Me.txtpayment.Text = ""
'Me.Text4.Text = ""
DataReport4.Show

Else

Call connect(cnPnt1, App.Path & "\myDB.mdb")
Call SetRs(rsPnt1, cnPnt1, "SELECT * from PartialPaid Where partialpaid.contractnumber ='" & Text5.Text & "' AND DatePaid='" & Text8.Text & "'")

Set DataReport5.DataSource = rsPnt1
Unload Me
'Me.txttotaldebt.Text = ""
'Me.txtpercent.Text = ""
'Me.txttotaldebttopay.Text = ""
'Me.Text2.Text = ""
'Me.txtpayment.Text = ""
'Me.Text4.Text = ""
DataReport5.Show
'End If
End If

Set rsPnt = Nothing
Set cnPnt = Nothing

Set rsPnt1 = Nothing
Set cnPnt1 = Nothing
End Sub

Private Sub Form_Load()
Me.cmdok.Enabled = True
Me.txtpayment.Text = ""
cmdprint.Enabled = False
Me.txtpercent.Text = ""
Me.txttotaldebttopay.Text = ""
Me.Text2.Text = ""
Text4.Text = ""
End Sub

Private Sub txtpayment_Change()


Text4.Text = Val(Text2.Text) - Val(txtpayment.Text)
If Val(txtpayment.Text) > Val(Text2.Text) Then
Me.Label1(9).Caption = "Change"
Label1(9).ForeColor = &H80000008
Else
Label1(9).Caption = "Balance"
Label1(9).ForeColor = &HFF&
End If
End Sub

Private Sub txtpayment_GotFocus()
If txtpercent.Text = "" Then
MsgBox "Percent Interest Can not be Null" & vbCrLf & "Please type Zero if does not have interest", vbInformation
txtpercent.SetFocus
End If
End Sub

Private Sub txtpayment_KeyPress(KeyAscii As Integer)
'If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

Private Sub txtpercent_Change()
Text2.Text = Format$("#,##0.00")
Me.txttotaldebttopay.Text = ((Val(Text1.Text) * Val(Me.txtpercent.Text)) / 100)
Text2.Text = Val(txttotaldebttopay) + Val(Text1.Text)
End Sub

Private Sub txtpercent_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0

Text2.Text = Format$("#,##0.00")

End Sub
