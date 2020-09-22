VERSION 5.00
Begin VB.UserControl cmdButton 
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ScaleHeight     =   3225
   ScaleWidth      =   2505
   Begin VB.Timer tmrCheck 
      Interval        =   1
      Left            =   1320
      Top             =   120
   End
   Begin VB.Image imgIO 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cmd"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   20
      Width           =   855
   End
   Begin VB.Image imgDown 
      Height          =   255
      Left            =   0
      Picture         =   "cmdButton.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgOver 
      Height          =   255
      Left            =   0
      Picture         =   "cmdButton.ctx":0BAE
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgNormal 
      Height          =   255
      Left            =   0
      Picture         =   "cmdButton.ctx":175C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "cmdButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Private Type POINTAPI
X As Long
Y As Long
End Type

Public Event Click()
Public Event MouseMove()

Private mpoiCursorPos As POINTAPI


Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
imgDown.Visible = True
RaiseEvent Click
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
tmrCheck.Enabled = True
imgOver.Visible = True
RaiseEvent MouseMove
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
imgDown.Visible = False
End Sub

Private Sub tmrCheck_Timer()
'If Not Enabled Then Exit Sub
Dim lonCStat As Long
Dim lonCurrhWnd As Long

lonCStat = GetCursorPos&(mpoiCursorPos)
lonCurrhWnd = WindowFromPoint(mpoiCursorPos.X, mpoiCursorPos.Y)
If lonCurrhWnd = UserControl.hwnd Then
Else
imgDown.Visible = False
imgOver.Visible = False
imgNormal.Visible = True
tmrCheck.Enabled = False
End If

End Sub

Private Sub UserControl_InitProperties()
Caption = Ambient.DisplayName
End Sub

Public Property Get Caption() As String
Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
lblCaption.Caption = sNewValue
UserControl.PropertyChanged "Caption"
End Property

Public Property Get FontName() As String
FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal sNewValue As String)
lblCaption.FontName = sNewValue
UserControl.PropertyChanged "FontName"
End Property

Private Sub UserControl_Resize()
lblCaption.Top = ((UserControl.Height - lblCaption.Height) / 2) + 20

imgOver.Width = UserControl.Width
imgOver.Height = UserControl.Height
imgNormal.Width = UserControl.Width
imgNormal.Height = UserControl.Height
imgDown.Width = UserControl.Width
imgDown.Height = UserControl.Height
imgIO.Width = UserControl.Width
imgIO.Height = UserControl.Height
lblCaption.Width = UserControl.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Caption", Caption, Ambient.DisplayName
    .WriteProperty "FontName", FontName, "Tahoma"
    .WriteProperty "Enabled", Enabled, True
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Caption = .ReadProperty("Caption", Ambient.DisplayName)
    FontName = .ReadProperty("FontName", "Tahoma")
    Enabled = .ReadProperty("Enabled", True)
End With
End Sub

Private Sub imgio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
imgDown.Visible = True
RaiseEvent Click
End Sub

Private Sub imgio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
tmrCheck.Enabled = True
imgOver.Visible = True
RaiseEvent MouseMove
End Sub

Private Sub imgio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
imgDown.Visible = False
End Sub

Public Property Get Enabled() As Boolean
Enabled = lblCaption.Enabled
End Property

Public Property Let Enabled(ByVal bNewValue As Boolean)
lblCaption.Enabled = bNewValue
End Property
