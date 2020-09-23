VERSION 5.00
Begin VB.Form frmnewcategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Category"
   ClientHeight    =   2385
   ClientLeft      =   4710
   ClientTop       =   5040
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5505
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtcategory 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Category Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   1515
   End
End
Attribute VB_Name = "frmnewcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CNcateg As New ADODB.Connection
Dim RScateg As New ADODB.Recordset


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim ans As String

If txtcategory.Text = "" Then
MsgBox "Please Fill Category Field", vbInformation
Exit Sub
Else
txtcategory.Text = StrConv(txtcategory, vbProperCase)
Call connect(CNcateg, App.Path & "\mydb.mdb")
Call SetRs(RScateg, CNcateg, "SELECT * from stockscategory where StocksCategory= '" & Me.txtcategory.Text & "' ")

 If Not RScateg.EOF Then
        MsgBox "Category already exist", vbInformation
        Me.txtcategory.SelStart = 0
        Me.txtcategory.SelLength = Len(Me.txtcategory.Text)
        Me.txtcategory.SetFocus

    
Else
    
    RScateg.AddNew
    RScateg.Fields(0) = Me.txtcategory.Text
    RScateg.Update
    RScateg.Requery
    MsgBox "New Category of Stocks Has Been Added", vbInformation
    ans = MsgBox("Would You Like to Add Another Stocks Category?", vbInformation + vbYesNo)
    If ans = vbYes Then
    Exit Sub
    Else
    Unload Me
    End If
    
End If
    
 End If
 
   Set CNcateg = Nothing
    Set RScateg = Nothing
    
End Sub

Private Sub cmdOk_LostFocus()
txtcategory.Text = StrConv(txtcategory, vbProperCase)
End Sub

