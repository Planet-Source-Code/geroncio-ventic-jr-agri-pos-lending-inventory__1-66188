VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcustomerwbalance 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer's with Balance"
   ClientHeight    =   5190
   ClientLeft      =   3735
   ClientTop       =   2925
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7260
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   15
      Left            =   5160
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Enabled         =   0   'False
      Height          =   15
      Left            =   5160
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "frmcustomerwbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim cncategory5 As New ADODB.Connection
Dim rscategory5 As New ADODB.Recordset

Call connect(cncategory5, App.path & "\myDB.mdb")
Call SetRs(rscategory5, cncategory5, "SELECT * FROM PartialPaid where ContractNumber ='" & Label1.Caption & "'")
Set DataReport12.DataSource = rscategory5
Unload Me
DataReport12.Show

Set rscategory5 = Nothing
Set cncategory5 = Nothing
End Sub


