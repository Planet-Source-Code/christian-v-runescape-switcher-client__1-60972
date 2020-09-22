VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Snippets"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAlwOnTop 
      Caption         =   "Always On Top"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   840
   End
   Begin VB.TextBox txtSnip 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
'On Error Resume Next
'Open "c:\windows\lrsscsnip.txt" For Output As #1
'Print #1, txtSnip.Text
'Close #1
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkAlwOnTop_Click()
If chkAlwOnTop.Value = 1 Then
Call Always_On_Top(Me.hWnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, True)
Else
Call Always_On_Top(Me.hWnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, False)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim File

Call Always_On_Top(Me.hWnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, True)

Open "c:\windows\lrsscsnip.txt" For Input As #1
Do Until EOF(1)
Line Input #1, File
txtSnip.Text = txtSnip.Text & File & vbCrLf
Loop
Close #1
End Sub

Private Sub Form_Resize()
txtSnip.Width = Form2.Width - 345
txtSnip.Height = Form2.Height - 615 - 240
End Sub

Private Sub txtSnip_Change()
On Error GoTo e
Open "c:\windows\lrsscsnip.txt" For Output As #1
Print #1, txtSnip.Text
Close #1
e:
End Sub
