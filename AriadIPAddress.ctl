VERSION 5.00
Begin VB.UserControl AriadIPAddress 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "AriadIPAddress.ctx":0000
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ToolboxBitmap   =   "AriadIPAddress.ctx":0038
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   360
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "AriadIPAddress.ctx":0132
      Top             =   45
      Width           =   195
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   90
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "AriadIPAddress.ctx":0134
      Top             =   45
      Width           =   240
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   765
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "AriadIPAddress.ctx":0136
      Top             =   45
      Width           =   285
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   990
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "AriadIPAddress.ctx":0138
      Top             =   45
      Width           =   465
   End
   Begin VB.Image imgDot 
      Height          =   30
      Index           =   0
      Left            =   630
      Picture         =   "AriadIPAddress.ctx":013A
      Top             =   135
      Width           =   30
   End
   Begin VB.Image imgDot 
      Height          =   30
      Index           =   1
      Left            =   45
      Picture         =   "AriadIPAddress.ctx":018C
      Top             =   45
      Width           =   30
   End
   Begin VB.Image imgDot 
      Height          =   30
      Index           =   2
      Left            =   45
      Picture         =   "AriadIPAddress.ctx":01DE
      Top             =   45
      Width           =   30
   End
End
Attribute VB_Name = "AriadIPAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
DefInt A-Z

Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

Dim W As Single
Dim IPA$
Dim ThruIP  As Boolean
Dim m_UseDisabledBackColor As Boolean

Public Event Change(NewIP As String)
Attribute Change.VB_Description = "Occurs when the IP Address changes."
Attribute Change.VB_MemberFlags = "200"

Dim m_BackColor As OLE_COLOR
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim I
 I = Int(X / (W + 10))
 If I < 0 Then I = 0
 If I > 3 Then I = 3
 If Ambient.UserMode Then SetFocus txtIP(I).hWnd
End Sub


Private Sub txtIP_Change(Index As Integer)
 Dim I
 If IsValidIP() Then RaiseEvent Change(IPAddress)
 If Len(txtIP(Index)) = 3 And ThruIP = 0 Then
  I = Index + 1
  If I < 4 Then If Ambient.UserMode Then SetFocus txtIP(I).hWnd
 End If
End Sub

Private Sub txtIP_GotFocus(Index As Integer)
 txtIP(Index).SelStart = 0
 txtIP(Index).SelLength = 3
End Sub

Private Sub txtIP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 With txtIP(Index)
  If KeyCode = 37 And .SelStart = 0 Then
   If Index <> 0 Then
    With txtIP(Index - 1)
     .SetFocus
     .SelStart = Len(.Text)
     .SelLength = Len(.Text)
    End With
   End If
  ElseIf KeyCode = 39 And .SelStart = Len(.Text) Then
   If Index <> 3 Then
    With txtIP(Index + 1)
     .SetFocus
     .SelStart = 0
     .SelLength = 0
    End With
   End If
  End If
 End With
End Sub

Private Sub txtIP_KeyPress(Index As Integer, KeyAscii As Integer)
 Dim C$, T$
 Dim I
 C$ = Chr$(KeyAscii)
 If KeyAscii = 8 Then
 ElseIf C$ = "." Then
  I = Index + 1: If I = 4 Then I = 0
  If Ambient.UserMode Then SetFocus txtIP(I).hWnd
  KeyAscii = 0
 ElseIf C$ < "0" Or C$ > "9" Then
  KeyAscii = 0
  Beep
 End If
 If KeyAscii <> 0 And KeyAscii <> 8 Then
  If txtIP(Index).SelLength = 0 Then
   For I = 1 To Len(txtIP(Index))
    T$ = T$ + Mid$(txtIP(Index), I, 1)
   Next
   T$ = T$ + C$
   If Val(T$) > 255 Then
    KeyAscii = 0
    Beep
   End If
  End If
 End If
End Sub

Private Sub txtIP_LostFocus(Index As Integer)
 If txtIP(Index) = "" Then txtIP(Index) = "0"
 If Val(txtIP(Index)) > 255 Or Val(txtIP(Index)) < 0 Then
  MsgBox "Value must be between 0 and 255", 48
  If Ambient.UserMode Then SetFocus txtIP(Index).hWnd
  txtIP_GotFocus Index
 End If
End Sub

Private Sub UserControl_InitProperties()
 m_UseDisabledBackColor = -1
 BackColor = vbWindowBackground
 TextColor = vbWindowText
End Sub

Private Sub UserControl_Resize()
 Dim H As Single, MW As Single ', W As Single
 Dim I
 On Error GoTo ProcErr
  With UserControl
   MW = (TextWidth("XXX") + 10) * Screen.TwipsPerPixelY
   H = (TextHeight("X") + 6) * Screen.TwipsPerPixelX
   W = (((MW / Screen.TwipsPerPixelX) * 4) + 10) * Screen.TwipsPerPixelX
   If .Width < W Then Width = W
   If .Height < H Then Height = H
   W = (.ScaleWidth / 4) - 10
   For I = 0 To 3
    With txtIP(I)
     .Left = 5 + ((W + 10) * I)
     .Width = W
     .Height = TextHeight("X")
     .Top = (ScaleHeight - .Height) / 2
     If I < 3 Then
      imgDot(I).Left = .Left + .Width + 2
      imgDot(I).Top = .Top + .Height - 2
     End If
    End With
   Next
  End With
 On Error GoTo 0
Exit Sub

ProcErr:
 MsgBox Error$, 16
End Sub



Public Property Get IPAddress() As String
Attribute IPAddress.VB_Description = "Get/set the IP address displayed in the control."
Attribute IPAddress.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute IPAddress.VB_UserMemId = -518
Attribute IPAddress.VB_MemberFlags = "200"
 IPAddress = txtIP(0) + "." + txtIP(1) + "." + txtIP(2) + "." + txtIP(3)
End Property

Public Property Let IPAddress(ByVal IP As String)
 Dim P As Long, F As Long, V As Long
 Dim I
 Dim T$
 ThruIP = -1
  On Error GoTo ProcErr
   T$ = IP + "...."
   P = 1
   For I = 0 To 3
    F = InStr(P, T$, ".")
    If F <> 0 Then
     V = Val(Mid$(T$, P, F - P))
     If V < 0 Then V = 0
     If V > 255 Then V = 255
     txtIP(I) = V
     P = F + 1
    End If
   Next
   IP$ = IPAddress
   PropertyChanged "IPAddress"
   RaiseEvent Change(IP$)
  On Error GoTo 0
 ThruIP = 0
Exit Property

ProcErr:
 MsgBox Error$, 48
End Property

Public Property Get IPValue(Index As Integer) As Integer
Attribute IPValue.VB_Description = "Get/set individual sections of the controls'a IP Address."
Attribute IPValue.VB_ProcData.VB_Invoke_Property = ";Data"
 If Index > 3 Or Index < 0 Then
  MsgBox "Index must be between 0 and 3", 48
 Else
  IPValue = Val(txtIP(Index))
 End If
End Property

Public Property Let IPValue(Index As Integer, Value As Integer)
 Dim V As Long
 ThruIP = -1
  If Index > 3 Or Index < 0 Then
   MsgBox "Index must be between 0 and 3", 48
  Else
   V = Value
   If V < 0 Then V = 0
   If V > 255 Then V = 255
   txtIP(Index) = V
  End If
  PropertyChanged "IPValue"
  RaiseEvent Change(IPAddress)
 ThruIP = 0
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
 Dim I
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
 Set UserControl.Font = New_Font
 For I = 0 To 3
  Set txtIP(I).Font = New_Font
 Next
 UserControl_Resize
 PropertyChanged "Font"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
  Set Font = .ReadProperty("Font", Ambient.Font)
  IPAddress = .ReadProperty("IPAddress", "0.0.0.0")
  m_UseDisabledBackColor = .ReadProperty("UseDisabledBackColor", -1)
  Enabled = .ReadProperty("Enabled", -1)
  BackColor = .ReadProperty("BackColor", vbWindowBackground)
  TextColor = .ReadProperty("TextColor", vbWindowText)
 End With
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  .WriteProperty "Font", Font, Ambient.Font
  .WriteProperty "IPAddress", IPA$, "0.0.0.0"
  .WriteProperty "Enabled", Enabled, -1
  .WriteProperty "UseDisabledBackColor", m_UseDisabledBackColor, -1
  .WriteProperty "BackColor", BackColor, vbWindowBackground
  .WriteProperty "TextColor", TextColor, vbWindowText
 End With
End Sub


Public Sub About()
Attribute About.VB_Description = "Display copyright and version information."
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
 frmAriadAboutDialog.Show 1
End Sub

Public Function IsValidIP() As Boolean
Attribute IsValidIP.VB_Description = "Returns true if a specified IP Address is valid."
Attribute IsValidIP.VB_MemberFlags = "40"
 Dim I
 IsValidIP = -1
 For I = 0 To 3
  If Val(txtIP(I)) < 0 Or Val(txtIP(I)) > 255 Then IsValidIP = 0
 Next
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Get/set if the control is enabled."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal State As Boolean)
 Dim I
 Dim C As OLE_COLOR
 If State Then C = m_BackColor Else C = IIf(m_UseDisabledBackColor, vbButtonFace, m_BackColor)
 For I = 0 To 3
  txtIP(I).BackColor = C
  txtIP(I).Enabled = State
 Next
 UserControl.BackColor = C
 UserControl.Enabled = State
End Property

Public Property Get UseDisabledBackColor() As Boolean
Attribute UseDisabledBackColor.VB_Description = "Get/sets if the control uses the vbButtonFace colour when disable or not."
Attribute UseDisabledBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
 UseDisabledBackColor = m_UseDisabledBackColor
End Property

Public Property Let UseDisabledBackColor(ByVal NewVal As Boolean)
 m_UseDisabledBackColor = NewVal
 Enabled = Enabled
 PropertyChanged "UseDisabledBackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Set/get the background colour of the control."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
 BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewCol As OLE_COLOR)
 m_BackColor = NewCol
 Enabled = UserControl.Enabled
 PropertyChanged "BackColor"
End Property

Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Specifies the text colour of the control."
Attribute TextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
 TextColor = UserControl.ForeColor
End Property

Public Property Let TextColor(ByVal NewCol As OLE_COLOR)
 Dim I
 UserControl.ForeColor = NewCol
 For I = 0 To 3
  txtIP(I).ForeColor = NewCol
 Next
 PropertyChanged "TextColor"
End Property
