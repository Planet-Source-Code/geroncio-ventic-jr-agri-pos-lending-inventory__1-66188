VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFullypaid 
   Caption         =   "FullyPaid"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate Report"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1785
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   44957697
         CurrentDate     =   38723
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   44957697
         CurrentDate     =   38723
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3480
         TabIndex        =   3
         Top             =   120
         Width           =   315
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000006&
         BackStyle       =   1  'Opaque
         Height          =   45
         Left            =   240
         Top             =   1440
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmFullypaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim cngen As New ADODB.Connection
Dim rsgen As New ADODB.Recordset

Call connect(cngen, App.Path & "\myDB.mdb")
Call SetRs(rsgen, cngen, "Select * from FullyPAid where DatePaid BETWEEN '" & CDate(DTPicker1.Value) & "' AND  '" & CDate(DTPicker2.Value) & "'")
Set DataReport10.DataSource = rsgen
Unload Me
DataReport10.Show
Set cngen = Nothing
Set rsgen = Nothing

End Sub

