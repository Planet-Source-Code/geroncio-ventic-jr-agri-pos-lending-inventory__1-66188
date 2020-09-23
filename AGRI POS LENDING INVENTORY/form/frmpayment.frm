VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmpayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment"
   ClientHeight    =   7050
   ClientLeft      =   2610
   ClientTop       =   2580
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10845
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   13185
      TabIndex        =   31
      Top             =   8730
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   13200
      TabIndex        =   29
      Top             =   8730
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   4200
      TabIndex        =   28
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   13200
      TabIndex        =   27
      Top             =   8760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   13200
      TabIndex        =   26
      Top             =   8730
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   240
      ScaleHeight     =   5505
      ScaleWidth      =   4905
      TabIndex        =   7
      Top             =   360
      Width           =   4935
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   380
         Left            =   2160
         ScaleHeight     =   345
         ScaleWidth      =   945
         TabIndex        =   18
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
            TabIndex        =   0
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
            TabIndex        =   19
            Top             =   0
            Width           =   495
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
         TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
            Top             =   0
            Width           =   570
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
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         ScaleHeight     =   465
         ScaleWidth      =   2265
         TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
            Top             =   0
            Width           =   690
         End
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
         TabIndex        =   1
         Top             =   3000
         Width           =   2295
      End
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
         TabIndex        =   8
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Total Debt:"
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
         Left            =   960
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         TabIndex        =   24
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         TabIndex        =   23
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         TabIndex        =   22
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         TabIndex        =   21
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         TabIndex        =   20
         Top             =   3720
         Width           =   1110
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   5400
      ScaleHeight     =   5505
      ScaleWidth      =   5145
      TabIndex        =   6
      Top             =   360
      Width           =   5175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5055
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   15240
      TabIndex        =   5
      Top             =   9720
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   15240
      TabIndex        =   4
      Top             =   9720
      Width           =   150
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   9960
      TabIndex        =   37
      Top             =   0
      Width           =   15
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   8520
      TabIndex        =   36
      Top             =   0
      Width           =   15
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   7080
      TabIndex        =   35
      Top             =   0
      Width           =   15
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   0
      Width           =   15
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   3960
      TabIndex        =   33
      Top             =   0
      Width           =   15
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   2280
      TabIndex        =   32
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "frmpayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdclose_Click()
Unload Me

End Sub

Private Sub cmdok_Click()
On Error Resume Next
cmdPrint.Enabled = True
Dim RSpd As New ADODB.Recordset
Dim CNpd As New ADODB.Connection

Dim RSpd1 As New ADODB.Recordset
Dim CNpd1 As New ADODB.Connection

Dim RSpd2 As New ADODB.Recordset
Dim CNpd2 As New ADODB.Connection

Dim sagot  As String


    If (Val(txtpayment) >= Val(Text2)) Then
    
        Call connect(CNpd, App.path & "\myDB.mdb")
        Call SetRs(RSpd, CNpd, "SELECT * from FullyPaid")
            
            RSpd.addnew
            RSpd.Fields("Contractnumber") = Me.Text5.Text
            RSpd.Fields("Name") = Text6.Text
            RSpd.Fields("PercentInterest") = txtpercent.Text
            RSpd.Fields("InterestAmount") = Me.txttotaldebttopay.Text
            RSpd.Fields("AmountPaid") = txtpayment.Text
            RSpd.Fields("DatePaid") = Date
            RSpd.Update
            RSpd.Requery
            
            
        Call checkrs
        Call conek
         cnjen.Execute "Delete from DebtRecord where contractNumber = '" & Text5.Text & "'"
                '" & DataGrid1.Columns(0).Value
                rstransaction.Requery
     End If '
     'Else
     If (Val(txtpayment) < Val(Text2)) Then
     
     Call checkrs
        Call conek
         cnjen.Execute "Delete from DebtRecord where contractNumber = '" & Text5.Text & "'"
                '" & DataGrid1.Columns(0).Value
                rstransaction.Requery

     
        Call connect(CNpd1, App.path & "\myDB.mdb")
        Call SetRs(RSpd1, CNpd1, "SELECT * from debtrecord")
            
            RSpd1.addnew
            RSpd1.Fields("Contractnumber") = Label2.Caption
            RSpd1.Fields("Name") = Label3.Caption
            RSpd1.Fields!Item = "Cash"
            RSpd1.Fields!qunatity = "Cash"
            RSpd1.Fields!cost = Label5.Caption
            RSpd1.Fields!totalcost = Text4.Text
            RSpd1.Fields!Date = Label4.Caption
            RSpd1.Update
            RSpd1.Requery
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
               
        
        'Call connect(CNpd2, App.Path & "\myDB.mdb")
        'Call SetRs(RSpd2, CNpd2, "Delete * from DebtRecord Where contractNumber =" & Text5.Text & "")
         '   RSpd2.delete
          '   RSpd2.Requery
           '  If RSpd2.EOF Then
            '            RSpd2.MoveLast
            'End If
            'Set DataGrid1.DataSource = RSpd2
            Me.cmdok.Enabled = False
        

Set CNpd = Nothing
Set RSpd = Nothing

Set CNpd1 = Nothing
Set RSpd1 = Nothing

Set CNpd2 = Nothing
Set RSpd2 = Nothing

End Sub


Private Sub cmdPrint_Click()
'On Error Resume Next
Dim cnPnt As New ADODB.Connection
Dim rsPnt As New ADODB.Recordset

Dim cnPnt1 As New ADODB.Connection
Dim rsPnt1 As New ADODB.Recordset


        If (Val(txtpayment) >= Val(Text2)) Then
            Call connect(cnPnt, App.path & "\myDB.mdb")
            Call SetRs(rsPnt, cnPnt, "SELECT * from FullyPaid Where fullypaid.contractnumber ='" & Text5.Text & "' AND DatePaid ='" & Text8.Text & "'")
            
            Set DataReport1.DataSource = rsPnt
            Unload Me
            DataReport1.Show
        Else
            
             Call connect(cnPnt1, App.path & "\myDB.mdb")
            Call SetRs(rsPnt1, cnPnt1, "SELECT * from debtrecord Where debtrecord.contractnumber ='" & Label2.Caption & "'") ' AND DatePaid ='" & Label4.Caption & "'")
            Set DataReport2.DataSource = rsPnt1
           
            Unload Me
            
            DataReport2.Show
           
        End If
        
Set rsPnt = Nothing
Set cnPnt = Nothing

Set rsPnt1 = Nothing
Set cnPnt1 = Nothing
End Sub

Private Sub Form_Load()
        Me.cmdok.Enabled = True
        Me.txtpayment.Text = ""
        cmdPrint.Enabled = False
        Me.txtpercent.Text = ""
        Me.txttotaldebttopay.Text = ""
        Me.Text2.Text = ""
        Text4.Text = ""
End Sub

Private Sub txtpayment_Change()
Label7.Caption = txtpayment.Text

Text4.Text = Val(Text2) - Val(txtpayment)
Label6.Caption = Val(Text2) - Val(txtpayment)
    If Val(txtpayment) > Val(Text2) Then
        Me.Label1(9).Caption = "Change"
        Label1(9).ForeColor = &H80000008
        Label6.Caption = "0.00"
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
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0

End Sub

Private Sub txtpercent_Change()
    Text2.Text = Format$("#,##0.00")
    Me.txttotaldebttopay.Text = ((Val(Text1.Text) * Val(Me.txtpercent.Text)) / 100)
    Text2.Text = Val(txttotaldebttopay) + Val(Text1.Text)
    Label5.Caption = Val(txttotaldebttopay) + Val(Text1.Text)

End Sub

Private Sub txtpercent_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0

Text2.Text = Format$("#,##0.00")

End Sub

