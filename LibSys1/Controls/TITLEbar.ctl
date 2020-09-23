VERSION 5.00
Begin VB.UserControl TITLEbar 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   Begin VB.PictureBox imgClose 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   2970
      Picture         =   "TITLEbar.ctx":0000
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   1
      ToolTipText     =   "Close This Window"
      Top             =   30
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox imgClose 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3660
      Picture         =   "TITLEbar.ctx":028A
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   0
      ToolTipText     =   "Close This Window"
      Top             =   60
      Width           =   345
   End
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4485
      Top             =   15
   End
   Begin VB.Image imgShadow 
      Height          =   450
      Index           =   1
      Left            =   1755
      Picture         =   "TITLEbar.ctx":0532
      Top             =   2790
      Width           =   4500
   End
   Begin VB.Image imgBGRight 
      Height          =   450
      Index           =   1
      Left            =   3570
      Picture         =   "TITLEbar.ctx":07C4
      Top             =   2355
      Width           =   4500
   End
   Begin VB.Image imgBG 
      Height          =   345
      Index           =   1
      Left            =   1440
      Picture         =   "TITLEbar.ctx":0A56
      Stretch         =   -1  'True
      Top             =   2250
      Width           =   1590
   End
   Begin VB.Image imgBGLeft 
      Height          =   450
      Index           =   1
      Left            =   1080
      Picture         =   "TITLEbar.ctx":0CE8
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   60
   End
   Begin VB.Image imgBGRight 
      Height          =   450
      Index           =   0
      Left            =   -4080
      Picture         =   "TITLEbar.ctx":0F7A
      Top             =   -120
      Width           =   4500
   End
   Begin VB.Image imgBGLeft 
      Height          =   345
      Index           =   0
      Left            =   0
      Picture         =   "TITLEbar.ctx":120C
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgBG 
      Height          =   345
      Index           =   0
      Left            =   0
      OLEDropMode     =   2  'Automatic
      Picture         =   "TITLEbar.ctx":1401
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5310
   End
End
Attribute VB_Name = "TITLEbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long




'events
Public Event CloseMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event CloseMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event CLoseClick()



Dim MouseOnDOwn As Boolean
Dim iCurIMGIndex As Integer


'Default Property Values:
Const m_def_AutoFunction = True
'Property Variables:
Dim m_ShadowVisible As Boolean
Dim m_AutoFunction As Boolean
'Event Declarations:
Event DblClick() 'MappingInfo=imgBG,imgBG,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."


Private Sub FormDrag(frmName As Form) 'procedure to drag a no-titlebar form
    ReleaseCapture
    Call SendMessage(frmName.hwnd, &HA1, 2, 0&)
End Sub


Private Sub imgBG_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(1).Visible = False
    RaiseEvent CloseMouseDown(Button, Shift, X, Y)
    
    MouseOnDOwn = True
End Sub

Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(1).Visible = IIf(MouseOnDOwn, False, True)
    timerMouse.Enabled = True
End Sub

Private Sub imgClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim P As POINTAPI
    Dim R As RECT
    
    imgClose(1).Visible = True
    
    GetWindowRect imgClose(0).hwnd, R
    GetCursorPos P
    
    If Not (P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom) Then
        
        RaiseEvent CLoseClick
        
        If AutoFunction = True Then
            On Error Resume Next
            Unload UserControl.Parent
        End If
    End If
    
    
    RaiseEvent CloseMouseUp(Button, Shift, X, Y)
    MouseOnDOwn = False
    
End Sub



Private Sub timerMouse_Timer()
    Dim P As POINTAPI
    Dim R As RECT

    GetWindowRect imgClose(0).hwnd, R
    GetCursorPos P
    
    If P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom Then
        timerMouse.Enabled = False
        imgClose(1).Visible = False
    End If
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_Initialize()
    iCurIMGIndex = 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And AutoFunction = True Then
        FormDrag UserControl.Parent
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    UserControl.Height = 23 * Screen.TwipsPerPixelY
    
    imgBGLeft(iCurIMGIndex).Move 0, 0
    imgBGLeft(iCurIMGIndex).Visible = True
    
    imgBG(iCurIMGIndex).Move 2, 0, GetWidth - 2, GetHeight
    imgBG(iCurIMGIndex).Visible = True
    
    imgBGRight(iCurIMGIndex).Move GetWidth - 2, 0
    imgBGRight(iCurIMGIndex).Visible = True
    
    If m_ShadowVisible = True Then
        imgShadow(iCurIMGIndex).Move 1, 0
        imgShadow(iCurIMGIndex).Visible = True
    End If
    
    imgClose(0).Move GetWidth - 2 - imgClose(0).Width, 2
    imgClose(1).Move GetWidth - 2 - imgClose(1).Width, 2

End Sub


Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelY
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelX
End Function


Public Property Let Caption(ByVal New_Caption As String)
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    
    'If New_Caption = lblCaption.Caption Then Exit Property
    
   ' blCaption.Caption() = New_Caption
    
   ' If AutoFunction = True Then
        UserControl.Parent.Caption = New_Caption
    'End If
    
   ' PropertyChanged "Caption"
End Property

Private Sub imgBG_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
   ' FontSize = lblCaption.FontSize
End Property



Public Property Get AutoFunction() As Boolean
    AutoFunction = m_AutoFunction
End Property

Public Property Let AutoFunction(ByVal New_AutoFunction As Boolean)
    m_AutoFunction = New_AutoFunction
    PropertyChanged "AutoFunction"
End Property

Private Sub UserControl_InitProperties()
    m_AutoFunction = m_def_AutoFunction
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)

    m_AutoFunction = PropBag.ReadProperty("AutoFunction", m_def_AutoFunction)
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)

End Sub




Public Property Let ShadowVisible(ByVal New_ShadowVisible As Boolean)
    imgShadow(iCurIMGIndex).Visible = New_ShadowVisible
    m_ShadowVisible = New_ShadowVisible
    PropertyChanged "ShadowVisible"
End Property


Public Sub ParentFocus(Optional OnFocus As Boolean = True)
    If OnFocus = True Then
        iCurIMGIndex = 0
        imgBGLeft(1).Visible = False
        imgBG(1).Visible = False
        imgBGRight(1).Visible = False
        imgShadow(1).Visible = False
    Else
        iCurIMGIndex = 1
        imgBGLeft(0).Visible = False
        imgBG(0).Visible = False
        imgBGRight(0).Visible = False
        imgShadow(0).Visible = False
        
    End If
    
    Call UserControl_Resize
End Sub
