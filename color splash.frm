VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "color splash.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   [D] Light Studio v.2
'
'   By Danish Mujeeb, comments------> d_mujeeb@hotmail.com
'
'   I apolagize for not commeting the code at all since I made the program
'   for myself. However it turned out to be quite interesting and I thought
'   that it should be shared. The idea of inverse square law has been applied
'   BUT only through trial and error, it only works for this program. However,
'   what it does is pretty cool :)


Private Sub Command1_Click()
Unload Form2
'Form1.Show 1
End Sub

Private Sub Command2_Click()
Unload Form2
'Form1.Show 1
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub
