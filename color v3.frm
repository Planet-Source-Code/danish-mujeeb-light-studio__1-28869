VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00000000&
   Caption         =   "[D] Light Studio"
   ClientHeight    =   8025
   ClientLeft      =   1860
   ClientTop       =   675
   ClientWidth     =   7725
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   Begin VB.CommandButton Command5 
      Height          =   135
      Left            =   240
      TabIndex        =   44
      Top             =   1995
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   43
      Text            =   "1"
      Top             =   7680
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawWidth       =   5
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   1575
      MousePointer    =   2  'Cross
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
   Begin VB.PictureBox quit_but_down 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7020
      Picture         =   "color v3.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   42
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox quit_but_up 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7005
      Picture         =   "color v3.frx":0748
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   41
      Top             =   6810
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox command3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6315
      Picture         =   "color v3.frx":0E90
      ScaleHeight     =   360
      ScaleWidth      =   975
      TabIndex        =   40
      Top             =   30
      Width           =   975
   End
   Begin VB.PictureBox full_but_up 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5910
      Picture         =   "color v3.frx":15D8
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   39
      Top             =   6810
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox full_but_down 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5925
      Picture         =   "color v3.frx":1D4C
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox command6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5340
      Picture         =   "color v3.frx":24C0
      ScaleHeight     =   360
      ScaleWidth      =   975
      TabIndex        =   37
      Top             =   30
      Width           =   975
   End
   Begin VB.PictureBox stp_but_down 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5175
      Picture         =   "color v3.frx":2C08
      ScaleHeight     =   375
      ScaleWidth      =   705
      TabIndex        =   36
      Top             =   7185
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox stp_but_up 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5175
      Picture         =   "color v3.frx":3350
      ScaleHeight     =   375
      ScaleWidth      =   690
      TabIndex        =   35
      Top             =   6810
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox full_screen 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3855
      Picture         =   "color v3.frx":3A98
      ScaleHeight     =   375
      ScaleWidth      =   1470
      TabIndex        =   34
      Top             =   30
      Width           =   1470
   End
   Begin VB.PictureBox rand_but_down 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4575
      Picture         =   "color v3.frx":420C
      ScaleHeight     =   375
      ScaleWidth      =   540
      TabIndex        =   33
      Top             =   7230
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox rand_but_up 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4575
      Picture         =   "color v3.frx":4960
      ScaleHeight     =   375
      ScaleWidth      =   510
      TabIndex        =   32
      Top             =   6795
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox command4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2880
      Picture         =   "color v3.frx":50B4
      ScaleHeight     =   360
      ScaleWidth      =   975
      TabIndex        =   31
      Top             =   30
      Width           =   975
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      Picture         =   "color v3.frx":5808
      ScaleHeight     =   285
      ScaleWidth      =   1470
      TabIndex        =   30
      Top             =   6660
      Width           =   1470
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      Picture         =   "color v3.frx":A4C0
      ScaleHeight     =   285
      ScaleWidth      =   1470
      TabIndex        =   29
      Top             =   6120
      Width           =   1470
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      Picture         =   "color v3.frx":F170
      ScaleHeight     =   285
      ScaleWidth      =   1470
      TabIndex        =   28
      Top             =   3840
      Width           =   1470
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      Picture         =   "color v3.frx":13E14
      ScaleHeight     =   285
      ScaleWidth      =   1470
      TabIndex        =   27
      Top             =   2925
      Width           =   1470
   End
   Begin VB.OptionButton op_cir 
      BackColor       =   &H00000000&
      Caption         =   "Circle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   26
      Top             =   3225
      Width           =   975
   End
   Begin VB.OptionButton op_sqr 
      BackColor       =   &H00000000&
      Caption         =   "Square"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   25
      Top             =   3525
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      Picture         =   "color v3.frx":18AD8
      ScaleHeight     =   285
      ScaleWidth      =   1470
      TabIndex        =   24
      Top             =   2160
      Width           =   1470
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      Picture         =   "color v3.frx":1D794
      ScaleHeight     =   285
      ScaleWidth      =   1470
      TabIndex        =   23
      Top             =   1320
      Width           =   1470
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   -30
      Picture         =   "color v3.frx":22450
      ScaleHeight     =   285
      ScaleWidth      =   1470
      TabIndex        =   22
      Top             =   435
      Width           =   1470
   End
   Begin VB.PictureBox edit_but_down 
      Height          =   375
      Left            =   3750
      Picture         =   "color v3.frx":270F0
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   21
      Top             =   7245
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox edit_but_up 
      Height          =   375
      Left            =   3750
      Picture         =   "color v3.frx":274E5
      ScaleHeight     =   315
      ScaleWidth      =   705
      TabIndex        =   20
      Top             =   6825
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox command7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1950
      Picture         =   "color v3.frx":278DA
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   19
      Top             =   45
      Width           =   975
   End
   Begin VB.PictureBox new_but_up 
      Height          =   375
      Left            =   2760
      Picture         =   "color v3.frx":27CCF
      ScaleHeight     =   315
      ScaleWidth      =   840
      TabIndex        =   18
      Top             =   6810
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox new_but_down 
      Height          =   375
      Left            =   2760
      Picture         =   "color v3.frx":280AC
      ScaleHeight     =   315
      ScaleWidth      =   870
      TabIndex        =   17
      Top             =   7230
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox command2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1005
      Picture         =   "color v3.frx":28489
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   16
      Top             =   45
      Width           =   975
   End
   Begin VB.PictureBox rend_but_down 
      Height          =   375
      Left            =   1815
      Picture         =   "color v3.frx":28866
      ScaleHeight     =   315
      ScaleWidth      =   840
      TabIndex        =   15
      Top             =   7215
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox rend_but_up 
      Height          =   375
      Left            =   1815
      Picture         =   "color v3.frx":28C5B
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   14
      Top             =   6795
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox command1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   60
      Picture         =   "color v3.frx":29050
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   13
      Top             =   45
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.PictureBox Picture5 
      Height          =   375
      Left            =   135
      ScaleHeight     =   315
      ScaleWidth      =   1245
      TabIndex        =   11
      Top             =   7200
      Width           =   1305
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   1425
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   1050
      ScaleHeight     =   405
      ScaleWidth      =   285
      TabIndex        =   7
      Top             =   810
      Width           =   285
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   585
      ScaleHeight     =   405
      ScaleWidth      =   285
      TabIndex        =   6
      Top             =   810
      Width           =   285
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   405
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   810
      Width           =   285
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   495
      Max             =   20
      Min             =   1
      TabIndex        =   3
      Top             =   2445
      Value           =   1
      Width           =   990
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   495
      Max             =   500
      Min             =   10
      TabIndex        =   1
      Top             =   1680
      Value           =   10
      Width           =   990
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   135
      TabIndex        =   10
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6420
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   2445
      Width           =   345
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   1680
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub make_pic(pic As PictureBox)
stp = 0
Picture1.Cls
Usat = HScroll1.Value
Usize = HScroll2.Value
Picture1.AutoRedraw = False
Command1.Enabled = False
Command2.Enabled = False
command4.Enabled = False
'Command5.Enabled = False
command7.Enabled = False
If HScroll2.Value = 1 Then
    
   
    
    For i = 1 To Picture1.Width
        For j = 1 To Picture1.Height
                rin = 0
                bin = 0
                gin = 0
            For k = 1 To n
            If stp = 1 Then stp = 0: Call ShowPoints: GoTo endingg
                
                If points(k, 3) = 1 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ 1) * HScroll1.Value
                    rin = rin + Intensity
                End If
                
                If points(k, 3) = 2 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ 1) * HScroll1.Value
                    gin = gin + Intensity
                End If
                
                If points(k, 3) = 3 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ 1) * HScroll1.Value
                    bin = bin + Intensity
                End If
                DoEvents
            Next k
            If rin > 255 Then rin = 255
            If gin > 255 Then gin = 255
            If bin > 255 Then bin = 255
            pic.PSet (i, j), RGB(rin, gin, bin)
        Next j
    Next i
End If


If HScroll2.Value > 1 Then
Picture1.DrawWidth = HScroll2
If c = 1 Then Picture1.Cls

    Command2.Enabled = False
    For i = 1 To Picture1.Width Step HScroll2.Value
        For j = 1 To Picture1.Height Step HScroll2.Value
                rin = 0
                bin = 0
                gin = 0
            For k = 1 To n
            If stp = 1 Then stp = 0: Call ShowPoints: GoTo endingg
                
                If points(k, 3) = 1 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ tst) * HScroll1.Value
                    rin = rin + Intensity
                End If
                
                If points(k, 3) = 2 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ tst) * HScroll1.Value
                    gin = gin + Intensity
                End If
                
                If points(k, 3) = 3 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ tst) * HScroll1.Value
                    bin = bin + Intensity
                End If
                DoEvents
            Next k
            
            If rin > 255 Then rin = 255
            If gin > 255 Then gin = 255
            If bin > 255 Then bin = 255
            
            If c = 2 Then pic.Line (i, j)-((i + HScroll2.Value), (j + HScroll2.Value)), RGB(rin, gin, bin), BF
            If c = 1 Then pic.PSet (i, j), RGB(rin, gin, bin)
            
        Next j
    Next i
End If

endingg:
Command1.Enabled = True
Command2.Enabled = True
command4.Enabled = True
'Command5.Enabled = True
command7.Enabled = True

Picture1.DrawWidth = 5
Picture1.AutoRedraw = True
End Sub

Private Sub command1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command1.Picture = rend_but_down.Picture
'Form3.Show
Call make_pic(Picture1)
End Sub

Private Sub command1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command1.Picture = rend_but_up.Picture
End Sub

Private Sub Command2_Click()
Picture1.Cls
List1.Clear
n = 0
End Sub

Private Sub command2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command2.Picture = new_but_down.Picture
End Sub

Private Sub command2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command2.Picture = new_but_up.Picture
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub command3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
command3.Picture = quit_but_down.Picture
End Sub

Private Sub command3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
command3.Picture = quit_but_up.Picture
End Sub

Private Sub Command4_Click()
n = 0
List1.Clear
Randomize Timer
n = (Int(10 * Rnd)) + 5
HScroll1.Value = Int(200 * Rnd) + 55
For i = 1 To n
    For j = 1 To 10000: Next
    
    Randomize Timer
    points(i, 1) = Int(Rnd * Picture1.ScaleWidth)
    points(i, 2) = Int(Rnd * Picture1.ScaleHeight)
    points(i, 3) = (Int(Rnd * 3)) + 1
    
    If points(i, 3) = 1 Then cc$ = "R, "
    If points(i, 3) = 2 Then cc$ = "G, "
    If points(i, 3) = 3 Then cc$ = "B, "
    
    List1.AddItem cc$ + Str$(points(i, 1)) + ", " + Str$(points(i, 2))
    
Next i

Call make_pic(Picture1)
End Sub

Private Sub command4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
command4.Picture = rand_but_down.Picture
End Sub

Private Sub command4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
command4.Picture = rand_but_up.Picture
End Sub

Private Sub Command6_Click()
stp = 1
End Sub

Private Sub command6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
command6.Picture = stp_but_down.Picture
End Sub

Private Sub command6_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
command6.Picture = stp_but_up.Picture
End Sub

Private Sub Command7_Click()
Call ShowPoints
End Sub

Private Sub command7_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
command7.Picture = edit_but_down.Picture
End Sub

Private Sub command7_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
command7.Picture = edit_but_up.Picture
End Sub

Private Sub Command8_Click()
If List1.Text <> "" Then
    k = (List1.ListIndex)
    For i = k + 1 To n - 1
        points(i, 1) = points(i + 1, 1)
        points(i, 2) = points(i + 1, 2)
        points(i, 3) = points(i + 1, 3)
    Next i
    n = n - 1
    Call ShowPoints
    Call UpdateList
End If
End Sub

Private Sub Form_GotFocus()
Picture1.AutoRedraw = True

End Sub

Private Sub Form_Load()
Form1.Left = (Screen.Width / 2) - (Form1.Width / 2)
Form1.Top = (Screen.Height / 2) - (Form1.Height / 2)
'1=red
'2=green
'3=blue
HScroll1.Value = 200
HScroll2.Value = 4
stp = 0
coll = 1
n = 0
c = 2
ful_scr = 0
tst = 1

max_dist = Sqr(((Picture1.ScaleWidth) ^ 2) + ((Picture1.ScaleHeight) ^ 2))
sc = max_dist / 10

'op_blue.Enabled
Label1.Caption = Str$(HScroll1.Value)
'MsgBox "WARNING: This program is only licensed to Danish Mujeeb" & vbLf & vbLf _
        & "Copyright: Danish Mujeeb" & vbLf & vbLf & "2000", vbInformation, "Hyper GRAFIX"

Form2.Show 1
Form3.Show

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Resize()
If Form1.ScaleWidth - Picture1.Left > 2 Then
If Form1.ScaleHeight - Picture1.Top > 2 Then
    Picture1.Width = (Form1.ScaleWidth - Picture1.Left)
    Picture1.Height = (Form1.ScaleHeight - Picture1.Top)
End If
End If



End Sub

Private Sub full_screen_Click()
o_form_left = Form1.Left
o_form_top = Form1.Top
o_form_width = Form1.Width
o_form_height = Form1.Height
o_pic_top = Picture1.Top
o_pic_left = Picture1.Left

Form1.BorderStyle = 0
Form1.Top = 0
Form1.Left = 0
Picture1.Top = 0
Picture1.Left = 0
Form1.Width = Screen.Width
Form1.Height = Screen.Height
ful_scr = 1
Call Command1_Click

End Sub

Private Sub full_screen_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
full_screen.Picture = full_but_down.Picture
End Sub

Private Sub full_screen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
full_screen.Picture = full_but_up.Picture
End Sub

Private Sub HScroll1_Change()
Label1.Caption = Str$(HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
Label1.Caption = Str$(HScroll1.Value)
End Sub

Private Sub HScroll2_Change()
Label4.Caption = Str$(HScroll2.Value)
End Sub

Private Sub HScroll2_Scroll()
Label4.Caption = Str$(HScroll2.Value)
End Sub

Private Sub List1_Click()
Picture1.Cls
Call ShowPoints
Picture1.DrawWidth = 1
If List1.Text <> "" Then
    i = List1.ListIndex + 1
    lx = points(i, 1)
    ly = points(i, 2)
    Picture1.Line (lx, 0)-(lx, Picture1.Height)
    Picture1.Line (0, ly)-(Picture1.Width, ly)
End If
End Sub

Private Sub op_cir_Click()
c = 1
End Sub

Private Sub op_sqr_Click()
c = 2
End Sub

Private Sub Picture1_DblClick()
If ful_scr = 1 Then
    Form1.Left = o_form_left
    Form1.Top = o_form_top
    Form1.Width = o_form_width
    Form1.Height = o_form_height
    Picture1.Top = o_pic_top
    Picture1.Left = o_pic_left
    ful_scr = 0
End If

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If ful_scr = 0 Then
    n = n + 1
    points(n, 1) = x
    points(n, 2) = Y
    points(n, 3) = coll
    If coll = 1 Then
        Picture1.PSet (x, Y), RGB(255, 0, 0)
        List1.AddItem "R, " + Str$(x) + ", " + Str$(Y)
    End If
    
    If coll = 2 Then
        Picture1.PSet (x, Y), RGB(0, 255, 0)
        List1.AddItem "G, " + Str$(x) + ", " + Str$(Y)
    End If
    
    If coll = 3 Then
        Picture1.PSet (x, Y), RGB(0, 0, 255)
        List1.AddItem "B, " + Str$(x) + ", " + Str$(Y)
    End If
End If
End Sub

Public Sub ShowPoints()
Picture1.DrawWidth = 5
Picture1.Cls
For i = 1 To n
    Xloc = points(i, 1)
    Yloc = points(i, 2)
    cc = points(i, 3)
    
    If cc = 1 Then Picture1.PSet (Xloc, Yloc), RGB(256, 0, 0)
    If cc = 2 Then Picture1.PSet (Xloc, Yloc), RGB(0, 256, 0)
    If cc = 3 Then Picture1.PSet (Xloc, Yloc), RGB(0, 0, 256)
Next i

End Sub
Public Sub UpdateList()
List1.Clear

For i = 1 To n
    If points(i, 3) = 1 Then
        List1.AddItem "R, " + Str$(points(i, 1)) + ", " + Str$(points(i, 2))
    End If

    If points(i, 3) = 2 Then
        List1.AddItem "G, " + Str$(points(i, 1)) + ", " + Str$(points(i, 2))
    End If
        
    If points(i, 3) = 3 Then
        List1.AddItem "B, " + Str$(points(i, 1)) + ", " + Str$(points(i, 2))
    End If
Next i
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label8.Caption = Str$(x) + ", " + Str$(Y)


i = x
j = Y

                rin = 0
                bin = 0
                gin = 0
            For k = 1 To n
                
                If points(k, 3) = 1 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ 1) * Usat
                    rin = rin + Intensity
                End If
                
                If points(k, 3) = 2 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ 1) * Usat
                    gin = gin + Intensity
                End If
                
                If points(k, 3) = 3 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ 1) * Usat
                    bin = bin + Intensity
                End If
                DoEvents
            Next k
            
            If rin > 255 Then rin = 255
            If gin > 255 Then gin = 255
            If bin > 255 Then bin = 255
            'Picture1.PSet (i, j), RGB(rin, gin, bin)
            Label9.Caption = Str$(Int(rin)) + "," + Str$(Int(gin)) + "," + Str$(Int(bin))
            Picture5.BackColor = RGB(rin, gin, bin)
            
            
End Sub

Private Sub Picture2_Click()
coll = 1
Picture2.BorderStyle = 1
Picture3.BorderStyle = 0
Picture4.BorderStyle = 0
End Sub

Private Sub Picture3_Click()
coll = 2
Picture2.BorderStyle = 0
Picture3.BorderStyle = 1
Picture4.BorderStyle = 0
End Sub

Private Sub Picture4_Click()
coll = 3
Picture2.BorderStyle = 0
Picture3.BorderStyle = 0
Picture4.BorderStyle = 1
End Sub


Private Sub Text1_Change()
tst = Val(Text1.Text)
End Sub


Private Sub Command5_Click()
Dim rin As Double
Dim bin As Double
Dim gin As Double
Open "H:\test.htm" For Output As #1
ew = 800 / (Picture1.Width / HScroll2.Value)
eh = 1 * ew
Print #1, "<html><head><title>just a test</title></head>"
Print #1, "<body bgcolor=black>"
Print #1, "<table border=0 cellpadding=0 cellspacing=1>"
For j = 1 To Picture1.Height Step HScroll2.Value
    Print #1, "<tr height=100>"
    For i = 1 To Picture1.Width Step HScroll2.Value
        
                rin = 0
                bin = 0
                gin = 0
            For k = 1 To n
            'If stp = 1 Then stp = 0: Call ShowPoints: GoTo endingg
                
                If points(k, 3) = 1 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ tst) * HScroll1.Value
                    rin = rin + Intensity
                End If
                
                If points(k, 3) = 2 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ tst) * HScroll1.Value
                    gin = gin + Intensity
                End If
                
                If points(k, 3) = 3 Then
                    X1 = i
                    Y1 = j
                    X2 = points(k, 1)
                    Y2 = points(k, 2)
                    dist = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
                    If dist <> 0 Then Intensity = (1 / (dist / sc) ^ tst) * HScroll1.Value
                    bin = bin + Intensity
                End If
                DoEvents
            Next k
            
            If rin > 255 Then rin = 255
            If gin > 255 Then gin = 255
            If bin > 255 Then bin = 255
            ccc$ = "#" + f(rin) + f(gin) + f(bin)
            Print #1, "<td bgcolor=" + ccc$ + " width=" + Str$(ew) + " height=" + Str$(eh) + "></td>"
            
            
        Next i
        Print #1, "</tr>"
    Next j
    Print #1, "</body></html>"
    MsgBox "done"
    Close #1
End Sub


Function f(x As Double) As String
    jj = Int((x / 255) * 5)
    If jj = 0 Then f = "00"
    If jj = 1 Then f = "33"
    If jj = 2 Then f = "66"
    If jj = 3 Then f = "99"
    If jj = 4 Then f = "cc"
    If jj = 5 Then f = "ff"
End Function


