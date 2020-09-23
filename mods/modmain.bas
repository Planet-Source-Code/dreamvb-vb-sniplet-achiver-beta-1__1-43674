Attribute VB_Name = "modmain"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Consts used for Flat effects
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const BS_HOLLOW = 0

' Window Consts
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2


Public tvID As String
Public TipID As Long
Public TabSelect As String
Public DelOption As String


Public Function FlatTextBox(frm As Form)
Dim I As Long
    For I = 0 To frm.Controls.Count - 1
        Select Case TypeName(frm.Controls(I))
            Case "TextBox", "ListBox"
                FlatBorder frm.Controls(I).hwnd, True
        End Select
    Next
    I = 0
End Function
Public Function FlatBorder(ByVal hwnd As Long, MakeControlFlat As Boolean)
Dim TFlat As Long
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    If MakeControlFlat Then
        TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        TFlat = TFlat And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    End If
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Function
