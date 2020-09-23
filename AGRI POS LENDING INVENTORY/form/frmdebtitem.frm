VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdebtitem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debt Item's"
   ClientHeight    =   3090
   ClientLeft      =   4710
   ClientTop       =   4845
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6255
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1785
      ScaleWidth      =   5745
      TabIndex        =   2
      Top             =   240
      Width           =   5775
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   38723
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   38723
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000006&
         BackStyle       =   1  'Opaque
         Height          =   45
         Left            =   240
         Top             =   1440
         Width           =   4935
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
         TabIndex        =   6
         Top             =   120
         Width           =   315
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
         TabIndex        =   5
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate Report"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "frmdebtitem"
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
Call SetRs(rsgen, cngen, "Select * from Debtrecord where Date BETWEEN '" & CDate(DTPicker1.Value) & "' AND  '" & CDate(DTPicker2.Value) & "'")
Set DataReport8.DataSource = rsgen
Unload Me
DataReport8.Show
Set cngen = Nothing
Set rsgen = Nothing

End Sub
