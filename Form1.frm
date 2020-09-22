VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click ME :)"
      Height          =   660
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1470
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   405
      Picture         =   "Form1.frx":0000
      Top             =   240
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Sub Command1_Click()
If Command1.Value = True Then
Image1.Visible = True
Command1.Visible = False
Me.BorderStyle = 0
    Dim lngRegion As Long
    Dim lngReturn As Long
    Dim lngFormWidth As Long
    Dim lngFormHeight As Long
    
    lngFormWidth = Me.Width / Screen.TwipsPerPixelX
    lngFormHeight = Me.Height / Screen.TwipsPerPixelY
    lngRegion = CreateEllipticRgn(0, 0, lngFormWidth, lngFormHeight)
    lngReturn = SetWindowRgn(Me.hWnd, lngRegion, True)
End If

End Sub

'Also you can place this code in the form load event
'    Dim lngRegion As Long
'    Dim lngReturn As Long
'    Dim lngFormWidth As Long
'    Dim lngFormHeight As Long
    
'    lngFormWidth = Me.Width / Screen.TwipsPerPixelX
'    lngFormHeight = Me.Height / Screen.TwipsPerPixelY
'    lngRegion = CreateEllipticRgn(0, 0, lngFormWidth, lngFormHeight)
'    lngReturn = SetWindowRgn(Me.hWnd, lngRegion, True)

Private Sub Form_Load()
Image1.Visible = False
End Sub
