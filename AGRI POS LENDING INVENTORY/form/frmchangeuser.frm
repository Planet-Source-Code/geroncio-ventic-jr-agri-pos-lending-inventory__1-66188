VERSION 5.00
Begin VB.Form frmchangeuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Username or Password"
   ClientHeight    =   5520
   ClientLeft      =   3540
   ClientTop       =   2925
   ClientWidth     =   8655
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8655
   Begin VB.TextBox txtNewUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtNewPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      TabIndex        =   8
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "&Change"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtou 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   6600
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   6480
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   6240
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1920
      X2              =   5760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      Height          =   3255
      Left            =   1560
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " NOTE: You must never forgot your USERNAME and   your  PASSWORD,  in oder to access this system."
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmchangeuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdchange_Click()


End Sub



