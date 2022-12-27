VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lambdas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Convert_ToIntI32(ByVal s As String) As Long
    Convert_ToIntI32 = CLng(s)
End Function

Public Function NewPersonS(ByVal s As String) As Person
    Set NewPersonS = MNew.Person(s)
End Function

Public Function F3(ByVal n As Integer) As Boolean
    F3 = n < 10
End Function

Public Function F4(ByVal n As Integer) As Integer
    '
End Function

Public Function var_Mod_10_eq_0(var As Long) As Boolean
    var_Mod_10_eq_0 = (var Mod 10) = 0
End Function
    
Public Function var_beginswith_A(var As String) As Boolean
    var_beginswith_A = Left(var, 1) = "A"
End Function

Public Function GridSettingsTypeToStr(V) As String
    GridSettingsTypeToStr = MUdt.GridSettingsTypeToStr(V)
End Function
