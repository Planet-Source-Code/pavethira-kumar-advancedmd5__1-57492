VERSION 5.00
Begin VB.Form AdvancedMD5 
   Caption         =   "AdvancedMD5"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hash It!"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "        String / Password :"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "                                                                                                    Hash :"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   3135
   End
End
Attribute VB_Name = "AdvancedMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'create variables (for MD5)
Dim text As String
Dim d As String
Dim fill As String

text = Text1.text


md5 = CalculateMD5(text)


length = Len(md5)
d = ""



'xor hash method
'confusing huh????
'very complicated to make sure there is no way decrypting
For i = 1 To length
 Char$ = Mid(md5, i, 1)
 code = i + 1
 code2 = i * code
 salt = i * 2
result = (((Asc(Char$) Xor code) + ((code2 * code) + salt)) Xor code2)
logans = Abs(Fix(Fix(Cos(result)) * 255 + Sin(result)))
result = result + ((length And i) Or (length Or i)) + logans
d = d & result
Next i



hash = CalculateMD5(d)
hash = StrReverse(hash)

'make encrypted password more complicated
'like ABCD to aBCd

fill = ""
For i = 1 To Len(hash)
    alph = Mid(hash, i, 1)
    getrand = (i * 2 + salt) Mod i
        
        If getrand Mod 2 = 0 Then
        alph = LCase(alph)
        
        Else
        alph = UCase(alph)
   
        End If

fill = fill & alph


Next i




lasthash = CalculateMD5(fill)

Text2.text = UCase(lasthash)


End Sub

Private Sub Command2_Click()
frmAbout.Show
End Sub

