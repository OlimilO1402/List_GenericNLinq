VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum EArt
    Golem  '= 0
    Elvea  '= 1
    Human  '= 2
    Ghost  '= 3
    Sauri  '= 4
End Enum
Private m_Name   As String
Private m_Art    As EArt
Private m_BirthD As Date

Private Sub Class_Initialize()
    m_Art = EArt.Human
End Sub

Friend Sub New_(aName As String, Optional aBirthD As Date, Optional aArt As EArt = Human)
    'normal c'tor
    m_Name = aName: m_BirthD = aBirthD: m_Art = aArt
End Sub
Friend Sub NewC(other As Person)
    'copy/clone-c'tor
    With other: New_ .Name, .Art, .BirthD: End With
End Sub
Friend Function Clone() As Person
    Set Clone = New Person: Clone.NewC Me
End Function

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Key() As String
    Key = m_Name
End Property

Public Property Get Art() As EArt
    Art = m_Art
End Property

Public Property Get BirthD() As Date
    BirthD = m_BirthD
End Property

Public Function Compare(ByVal other As Object) As Long
    'Dim p As Person: Set p = CPerson(other)
    Compare = StrComp(m_Name, other.Name)
    'Compare = DateDiff("d", m_BirthD, other.BirthD)
End Function

Public Function ToStr() As String
    ToStr = "Person{Name: " & m_Name & "; BirthD: " & m_BirthD & "; Art: " & EArt_ToStr(m_Art) & "}"
End Function

Public Function Parse(ByVal s As String)
    If UCase(Left(s, 6)) = "PERSON" Then
        s = Mid(s, 7)
        If Left(s, 1) = "{" Then s = Mid(s, 2)
        If Right(s, 1) = "}" Then s = Left(s, Len(s) - 1)
        Dim sa() As String: sa = Split(s, ";")
        Dim i As Long
        For i = 0 To UBound(sa)
            Dim sl() As String: sl = Split(Trim(sa(i)), ":")
            Select Case UCase(Trim(sl(0)))
            Case "NAME":   m_Name = Trim(sl(1))
            Case "ART":    m_Art = EArt_Parse(Trim(sl(1)))
            Case "BIRTHD": m_BirthD = Trim(sl(1))
            End Select
        Next
    End If
End Function

Function EArt_ToStr(aArt As EArt) As String
    Dim s As String
    Select Case aArt
    Case EArt.Golem: s = "Golem"
    Case EArt.Elvea: s = "Elvea"
    Case EArt.Human: s = "Human"
    Case EArt.Ghost: s = "Ghost"
    Case EArt.Sauri: s = "Sauri"
    End Select
    EArt_ToStr = s
End Function

Function EArt_Parse(ByVal s As String) As EArt
    Dim a As EArt
    Select Case UCase(s)
    Case "GOLEM": a = EArt.Golem
    Case "ELVEA": a = EArt.Elvea
    Case "HUMAN": a = EArt.Human
    Case "GHOST": a = EArt.Ghost
    Case "SAURI": a = EArt.Sauri
    End Select
    EArt_Parse = a
End Function
