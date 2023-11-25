VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub Clear():                                 End Sub
Public Function Clone() As Stack:                   End Function
Public Function Contains(Obj As Object) As Boolean: End Function
Public Sub CopyTo(Arr()):                           End Sub
Public Property Get Count() As Long:                End Property
Public Function Peek():                             End Function
Public Function Pop():                              End Function
Public Sub Push(Obj):                               End Sub
