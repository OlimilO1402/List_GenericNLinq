Attribute VB_Name = "MNew"
Option Explicit
Private Const LB_GETCURSEL As Long = &H188&
Private Const LB_SETCURSEL As Long = &H186&
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef src As Any, ByVal BytLen As Long)

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function lambdas() As lambdas
    Static lams As New lambdas
    Set lambdas = lams
End Function

Public Function PersonRnd() As Person
    Set PersonRnd = New Person: PersonRnd.New_ RndName, RndBirthD, RndArt
End Function
Public Function Person(aName As String, Optional aBirthD As Date, Optional aArt As EArt) As Person
    Set Person = New Person: Person.New_ aName, aBirthD, aArt
End Function

Public Function CPerson(ByVal aObj As Object) As Person
    Set CPerson = aObj
End Function

Public Function CLStr(aList As List) As String
    If Not aList.DataType = vbWChar Then Exit Function
    Dim L As Long: L = aList.Count: CLStr = Space$(L)
    RtlMoveMemory ByVal StrPtr(CLStr), ByVal aList.DataPtr, 2 * L
End Function

Function RndName() As String
    Dim s As String: s = Chr(65 + Rnd * 25)
    Dim i As Long
    For i = 1 To Rnd * 5 + 5
        s = s & Chr(97 + Rnd * 25)
    Next
    RndName = s
End Function

Function RndBirthD() As Date
    Dim y As Integer: y = 1919 + Rnd * 100
    Dim m As Integer: m = 1 + Rnd * 11
    Dim d As Integer
    Select Case m
    Case 1, 3, 5, 7, 8, 10, 12
        d = Rnd * 31
    Case 2
        d = Rnd * 28
    Case Else
        d = Rnd * 30
    End Select
    RndBirthD = DateSerial(y, m, d)
End Function

Function RndArt() As EArt
    RndArt = Rnd * 4
End Function

Function SplitL(s As String, Optional Delimiter, Optional ByVal Limit As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As List
    'Set SplitL = MNew.List(vbWChar, Split(s, Delimiter, Limit, Compare)) 'nein nicht so!
    Set SplitL = MNew.List(vbString, Split(s, Delimiter, Limit, Compare))
End Function

Function ArrayS(ParamArray ArgList() As Variant) As String()
    ReDim a(0 To UBound(ArgList)) As String
    Dim i As Long
    For i = 0 To UBound(a)
        a(i) = CStr(ArgList(i))
    Next
    ArrayS = a
End Function

Function ArrayContains(Arr(), var) As Boolean
    Dim V
    For Each V In Arr
        If Not VBA.Information.IsEmpty(V) And Not VBA.Information.IsMissing(V) Then
            If V = var Then ArrayContains = True: Exit Function
        End If
    Next
End Function

Sub EDataType_ToCombo(aCmb As ComboBox, ParamArray exclude())
    'ReDim exc(0)
    Dim exc()
    'If IsMissing(exclude) Then
    '    ReDim exc(0)
    'Else
        exc = exclude
    'End If
    With aCmb
        '.Clear
        Dim vt As VbVarType, s As String, c As Long
        For vt = vbInteger To vbUserDefinedType 'vbWChar
            s = IIf(ArrayContains(exc, vt), "", EDataType_ToStr(vt))
            If Len(s) Then
                .AddItem s
                .ItemData(c) = vt
                c = c + 1
            End If
        Next
        vt = vbArray
        s = IIf(ArrayContains(exc, vt), "", EDataType_ToStr(vt))
        If Len(s) Then .AddItem s
    End With
End Sub

Function EDataType_ToStr(ByVal vt As EDataType) As String
    Dim s As String
    Select Case vt
    Case EDataType.vbInteger:    s = "Integer"      ' vbInteger    =  2
    Case EDataType.vbLong:       s = "Long"         ' vbLong       =  3
    Case EDataType.vbSingle:     s = "Single"       ' vbSingle     =  4
    Case EDataType.vbDouble:     s = "Double"       ' vbDouble     =  5
    Case EDataType.vbCurrency:   s = "Currency"     ' vbCurrency   =  6
    
    
    Case EDataType.vbDate:       s = "Date"         ' vbDate       = 7
    Case EDataType.vbString:     s = "String"       ' vbString     = 8
    Case EDataType.vbObject:     s = "Object"       ' vbObject     = 9
    Case EDataType.vbBoolean:    s = "Boolean"      ' vbBoolean    = 11
    Case EDataType.vbVariant:    s = "Variant"      ' vbVariant    = 12
    Case EDataType.vbDataObject: s = "DataObject"   ' vbDataObject = 13
    Case EDataType.vbDecimal:    s = "Decimal"      ' vbDecimal    = 14
    Case EDataType.vbSByte:      s = "SByte"        ' vbSByte      = 16 (&H10)
    Case EDataType.vbByte:       s = "Byte"         ' vbByte       = 17 (&H11)
    Case EDataType.vbUInteger:   s = "UInteger"     ' vbUInteger   = 18
    Case EDataType.vbULong:      s = "ULong"        ' vbULong      = 19
    Case EDataType.vbLongLong:   s = "LongLong"     ' vbLongLong   = 20
    Case EDataType.vbULongLong:  s = "ULongLong"    ' vbULongLong  = 21
    Case EDataType.vbWChar:      s = "WChar"        ' vbWChar      = 32 (&H27)
    Case EDataType.vbUserDefinedType: s = "UserDefinedType" 'vbUserDefinedType = 36 (&H24)
    Case Else:
    End Select
    EDataType_ToStr = s
End Function

Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function

' ListIndex von ListBox ermitteln
' Gibt den aktuell selektierten Index zurück
Public Function LBGetListIndex(obj As ListBox) As Long
    LBGetListIndex = SendMessage(obj.hwnd, LB_GETCURSEL, 0, ByVal 0&)
End Function

' ListIndex von ListBox setzen
' setzt den ListIndex und markiert den Eintrag
' Ein Click wird nicht ausgelöst!
Public Sub LBSetListIndex(obj As ListBox, NewIndex As Long)
    Call SendMessage(obj.hwnd, LB_SETCURSEL, NewIndex, ByVal 0&)
End Sub

' The GridSettingsType is just for testing the ability to list structure-types in ax-dlls
Public Function GridSettingsTypeRnd() As GridSettingsType
    Randomize
    With GridSettingsTypeRnd
        .HeaderBold = IIf(Rnd - 0.5 < 0, False, True)
        .ShadeAltCols = IIf(Rnd - 0.5 < 0, False, True)
        .ShadeAltRows = IIf(Rnd - 0.5 < 0, False, True)
        .AllowColDragging = IIf(Rnd - 0.5 < 0, False, True)
        .AllowColSorting = IIf(Rnd - 0.5 < 0, False, True)
        .AllowDragAndSort = IIf(Rnd - 0.5 < 0, False, True)
        .GridStyle = IIf(Rnd - 0.5 < 0, 0, 1)
        .GridType = IIf(Rnd - 0.5 < 0, 0, 1)
    End With
End Function
