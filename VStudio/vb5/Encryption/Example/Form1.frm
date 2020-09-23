VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private test As Encryptor

Private Sub Form_Load()
Set test = New Encryptor

Dim Encrypted As String
Dim Decrypted As String
Encrypted = test.Encrypt("Hello My Name Is Jack !£$%^&*()", 9)
Form1.Print Encrypted
Decrypted = test.Decrypt(Encrypted, 9)
Form1.Print Decrypted
Encrypted = test.Encrypt("Jack", 128)
Form1.Print Encrypted
Decrypted = test.Decrypt(Encrypted, 128)
Form1.Print Decrypted
Dim Rot As Integer
Randomize Timer
Rot = Rnd * 127
Rot = Int(Rot) + 1
Encrypted = test.Encrypt("This is encrypted using a random number (" & Rot & ")", Rot)
Form1.Print Encrypted
Decrypted = test.Decrypt(Encrypted, Rot)
Form1.Print Decrypted

'Dim JetSon As String
'JetSon = "Jack"
'Dim Result As String
'Dim Code As Integer
'For i = 1 To Len(JetSon)
'Result = Mid$(JetSon, i, 1)
'Code = Asc(Result)
'MsgBox Result & "(" & Code & ")"
'Next i
End Sub


