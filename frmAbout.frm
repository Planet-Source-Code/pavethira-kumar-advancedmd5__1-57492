VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
