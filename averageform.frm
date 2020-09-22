VERSION 5.00
Begin VB.Form averageform 
   Caption         =   "DDAverager v1.01"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exitbutton 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton clearbutton 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton averagebutton 
      Caption         =   "&Average"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox secondnumberfield 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox firstnumberfield 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter second number"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter first number"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "averageform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub averagebutton_Click()
Dim numberone As Integer
Dim numbertwo As Integer
Dim result As Integer
If firstnumberfield.Text = "" And secondnumberfield.Text = "" Then
Exit Sub
ElseIf firstnumberfield.Text = "" Or secondnumberfield.Text = "" Then
Exit Sub
End If
numberone = firstnumberfield.Text
numbertwo = secondnumberfield.Text
result = (numberone + numbertwo) / 2
MsgBox "The average of " & numberone & ", " & numbertwo & " is " & result
End Sub

Private Sub clearbutton_Click()
firstnumberfield.Text = ""
secondnumberfield.Text = ""
End Sub

Private Sub exitbutton_Click()
End
End Sub
