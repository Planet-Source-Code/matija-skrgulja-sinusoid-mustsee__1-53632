VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   4995
   ClientTop       =   3300
   ClientWidth     =   9615
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   9615
   Begin VB.CheckBox Check2 
      Caption         =   "Sound freq  ( very slow )"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      Height          =   495
      Left            =   8280
      TabIndex        =   9
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Glue"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   7920
      Width           =   975
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   240
      Max             =   500
      Min             =   1
      TabIndex        =   6
      Top             =   8400
      Value           =   100
      Width           =   2655
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   500
      Min             =   1
      TabIndex        =   4
      Top             =   7440
      Value           =   100
      Width           =   2655
   End
   Begin VB.HScrollBar scroll1 
      Height          =   255
      Left            =   240
      Max             =   3577
      Min             =   1
      TabIndex        =   2
      Top             =   6600
      Value           =   90
      Width           =   2655
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   5895
      Left            =   0
      ScaleHeight     =   10.398
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   16.96
      TabIndex        =   1
      Top             =   0
      Width           =   9615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   3360
      Y1              =   8520
      Y2              =   8160
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   3360
      Y1              =   7560
      Y2              =   7920
   End
   Begin VB.Line Line2 
      X1              =   3000
      X2              =   3360
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line1 
      X1              =   3000
      X2              =   3360
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label3 
      Caption         =   "Freq 2 :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Freq 1 :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Amplitude :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Dim X, y, a, b, c
Dim Pause As Boolean
Dim Sound As Boolean


Private Sub Check1_Click()
If Check1.Value = 1 Then
    HScroll1.Value = HScroll2.Value
End If
End Sub

Private Sub Check2_Click()

If Check2.Value = 1 Then
    Sound = True
End If
If Check2.Value = 0 Then
    Sound = False
End If

End Sub

Private Sub Command1_Click()

Pause = False

Do

If Pause = True Then Exit Sub
    X = X + 0.01

    y = a * Cos(b * X) + a * Sin(c * X) + pic1.ScaleHeight / 2
    

    pic1.Left = pic1.Left - 1
    pic1.Width = pic1.Width + 1
    pic1.PSet (pic1.ScaleWidth - 1, pic1.ScaleHeight / 2), QBColor(8)
    pic1.PSet (pic1.ScaleWidth - 1, y), vbGreen
If Sound = True Then
    Beep y * 100, 1
End If
DoEvents
Loop
End Sub


Private Sub Command2_Click()

If Pause = False Then
    Pause = True
End If

End Sub

Private Sub Form_Load()
a = scroll1.Value / 100
b = HScroll1.Value / 100
c = HScroll2.Value / 100
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()

If Check1.Value = 1 Then
    HScroll2.Value = HScroll1.Value
End If

b = HScroll1.Value / 100
Label2.Caption = "Freq 1: " & b

End Sub

Private Sub HScroll1_Scroll()
If Check1.Value = 1 Then
    HScroll2.Value = HScroll1.Value
End If

b = HScroll1.Value / 100
Label2.Caption = "Freq 1: " & b

End Sub

Private Sub HScroll2_Change()
If Check1.Value = 1 Then
    HScroll1.Value = HScroll2.Value
End If

c = HScroll2.Value / 100
Label3.Caption = "Freq 2: " & c

End Sub

Private Sub HScroll2_Scroll()

If Check1.Value = 1 Then
    HScroll1.Value = HScroll2.Value
End If
c = HScroll2.Value / 100
Label3.Caption = "Freq 2: " & c
End Sub

Private Sub scroll1_Change()
a = scroll1.Value / 1000
Label1.Caption = "Amplitude: " & a
End Sub

Private Sub scroll1_Scroll()
a = scroll1.Value / 1000
Label1.Caption = "Amplitude: " & a
End Sub

