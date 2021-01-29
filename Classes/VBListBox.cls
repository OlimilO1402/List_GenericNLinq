VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LB_CTLCODE         As Long = &H0&
'Private Const LB_Okay As Long = 0

Private Const LB_ADDSTRING       As Long = &H180&
Private Const LB_INSERTSTRING    As Long = &H181& '<-
Private Const LB_DELETESTRING    As Long = &H182& '<-
Private Const LB_SELITEMRANGEEX  As Long = &H183&

Private Const LB_SETCURSEL       As Long = &H186& '<-
Private Const LB_GETSEL          As Long = &H187&
Private Const LB_GETCURSEL       As Long = &H188& '<-
Private Const LB_GETTEXT         As Long = &H189& '<-
Private Const LB_GETTEXTLEN      As Long = &H18A& '<-
Private Const LB_GETCOUNT        As Long = &H18B& '<-
Private Const LB_SELECTSTRING    As Long = &H18C&
Private Const LB_DIR             As Long = &H18D&

Private Const LB_FINDSTRING      As Long = &H18F&
Private Const LB_GETSELCOUNT     As Long = &H190&

Private Const LB_SETTABSTOPS     As Long = &H192&
Private Const LB_GETHORIZONTALEXTENT As Long = &H193&

Private Const LB_ADDFILE         As Long = &H196&
Private Const LB_SETTOPINDEX     As Long = &H197&
Private Const LB_GETITEMRECT     As Long = &H198&
Private Const LB_GETITEMDATA     As Long = &H199&
Private Const LB_SETITEMDATA     As Long = &H19A&
Private Const LB_SELITEMRANGE    As Long = &H19B&

Private Const LB_SETITEMHEIGHT   As Long = &H1A0&
Private Const LB_GETITEMHEIGHT   As Long = &H1A1&
Private Const LB_FINDSTRINGEXACT As Long = &H1A2&
Private Const LB_GETLOCALE       As Long = &H1A6&
Private Const LB_ITEMFROMPOINT   As Long = &H1A9&


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private m_LB   As VB.ListBox
Private m_hWnd As Long

Friend Sub New_(aLB As ListBox)
    Set m_LB = aLB
    m_hWnd = m_LB.hwnd
End Sub

Public Property Get ListIndex() As Long
    ListIndex = SendMessage(m_hWnd, LB_GETCURSEL, 0, ByVal 0&)
End Property
Public Property Let ListIndex(ByVal Value As Long)
    SendMessage m_hWnd, LB_SETCURSEL, Value, ByVal 0&
End Property

'
'' ListIndex von ListBox ermitteln
'' Gibt den aktuell selektierten Index zurück
'Public Function LBGetListIndex(Obj As ListBox) As Long
'    'LBGetListIndex = SendMessage(Obj.hwnd, LB_GETCURSEL, Null, ByVal Null)
'    LBGetListIndex = SendMessage(Obj.hwnd, LB_GETCURSEL, 0&, ByVal 0&)
'End Function
'
'' ListIndex von ListBox setzen
'' setzt den ListIndex und markiert den Eintrag
'' Ein Click wird nicht ausgelöst!
'Public Sub LBSetListIndex(Obj As ListBox, NewIndex As Long)
'    Call SendMessage(Obj.hwnd, LB_SETCURSEL, NewIndex, ByVal 0&)
'End Sub
'

Public Property Get Visible() As Boolean
    Visible = m_LB.Visible
End Property
Public Property Let Visible(ByVal Value As Boolean)
    m_LB.Visible = Value
End Property

' Liest den Text in einer ListBox aus mit dem
' angegebenen Index aus und gibt diesen zurück
Public Function LBGetItemText(Obj As ListBox, Index As Long) As String
    Dim Buffer As String
    Dim TextLen As Long
    TextLen = SendMessage(Obj.hwnd, LB_GETTEXTLEN, Index, ByVal 0&)
    If TextLen > 0 Then
        Buffer = Space(TextLen)
        Call SendMessage(Obj.hwnd, LB_GETTEXT, Index, ByVal Buffer)
    End If
    LBGetItemText = Buffer
End Function

' Ersetzt den Item-Text an der Position 'Index' mit 'NewText'
' Der Eintrag wird markiert!
Public Sub LBSetItemText(Obj As ListBox, Index As Long, NewText As String)
    ' zunächst Eintrag löschen
    If SendMessage(Obj.hwnd, LB_DELETESTRING, Index, ByVal 0&) > 0 Then
        ' und neuen Eintrag an diese Position wieder einfügen
        Call SendMessage(Obj.hwnd, LB_INSERTSTRING, Index, ByVal NewText)
        ' und noch selectieren
        Call SendMessage(Obj.hwnd, LB_SETCURSEL, Index, ByVal 0)
    End If
End Sub

