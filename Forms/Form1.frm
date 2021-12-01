VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Generic List And Linq"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   2640
      TabIndex        =   33
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnTestvbWChar 
      Caption         =   "Test vbWChar"
      Height          =   375
      Left            =   1800
      TabIndex        =   31
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton BtnTestSelect2 
      Caption         =   "Test Select2"
      Height          =   375
      Left            =   1800
      TabIndex        =   30
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton BtnTestSelect1 
      Caption         =   "Test Select1"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox TxtCount 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1440
      TabIndex        =   27
      ToolTipText     =   "The initial capacity of the inner array"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   6135
      Left            =   3600
      TabIndex        =   0
      Top             =   960
      Width           =   7215
   End
   Begin VB.CommandButton BtnBack 
      Caption         =   "< Back"
      Height          =   375
      Left            =   9600
      TabIndex        =   26
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   5010
      Left            =   3600
      TabIndex        =   25
      Top             =   960
      Width           =   7215
   End
   Begin VB.CommandButton BtnClone 
      Caption         =   "Clone >"
      Height          =   375
      Left            =   9600
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnSortDown 
      Caption         =   "Sort v"
      Height          =   375
      Left            =   8400
      TabIndex        =   23
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnSortUp 
      Caption         =   "Sort ^"
      Height          =   375
      Left            =   8400
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnMoveDown 
      Caption         =   "Move v"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnClearAll 
      Caption         =   "ClearAll"
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnMoveUp 
      Caption         =   "Move ^"
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtGrowSize 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1440
      TabIndex        =   12
      ToolTipText     =   "Growing either by factor or by chunksize, or both."
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox TxtGrowRate 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "Growing either by factor or by chunksize, or both."
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox CmbIsHashed 
      Height          =   345
      Left            =   1440
      TabIndex        =   8
      Text            =   "True/False"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox TxtCapacity 
      Alignment       =   1  'Rechts
      Height          =   345
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "The initial capacity of the inner array"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton BtnTestWhere 
      Caption         =   "Test Where"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton BtnCreate 
      Caption         =   "Create Random List"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox CmbDataType 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Test functions similar to Linq:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Count:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "           "
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "GrowSize:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "GrowRate:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Capacity:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "IsHashed:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Datatype:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_List As List
Dim m_ListClone As List

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription, vbInformation Or vbOKOnly
End Sub

Private Sub Form_Load()
    'Set m_List = New List
    EDataType_ToCombo CmbDataType, Empty
    CmbDataType.ListIndex = 0
    Boolean_ToCombo CmbIsHashed
    UpdateView
    Me.BtnBack.Enabled = False
    EnableCtrls False
End Sub
Sub EnableCtrls(bEnabled As Boolean)
    Me.BtnAdd.Enabled = bEnabled
    Me.BtnClearAll.Enabled = bEnabled
    Me.BtnClone.Enabled = bEnabled
    Me.BtnDelete.Enabled = bEnabled
    Me.BtnEdit.Enabled = bEnabled
    Me.BtnInsert.Enabled = bEnabled
    Me.BtnMoveDown.Enabled = bEnabled
    Me.BtnMoveUp.Enabled = bEnabled
    Me.BtnSearch.Enabled = bEnabled
    Me.BtnSortDown.Enabled = bEnabled
    Me.BtnSortUp.Enabled = bEnabled
    'Me.BtnBack.Enabled = bEnabled
End Sub
Sub Boolean_ToCombo(aCmb As ComboBox)
    aCmb.Clear
    aCmb.AddItem ""
    aCmb.AddItem "True"  ' "Wahr"
    aCmb.AddItem "False" ' "Falsch"
    'aCmb.Text = "Wahr/Falsch"
End Sub

Private Sub BtnCreate_Click()
    If CmbDataType.ListIndex < 0 Then
        MsgBox "Please select datatype first"
        Exit Sub
    End If
    Dim Of_Type  As EDataType:  Of_Type = CmbDataType.ItemData(CmbDataType.ListIndex)
    Dim Count    As Long:         Count = Lng_Parse(TxtCount.Text)
    Dim IsHashed As Boolean:   IsHashed = Bol_Parse(CmbIsHashed.Text)
    Dim Capacity As Long:      Capacity = Lng_Parse(TxtCapacity.Text)
    Dim GrowRate As Single:    GrowRate = Sng_Parse(TxtGrowRate.Text)
    Dim GrowSize As Long:      GrowSize = Lng_Parse(TxtGrowSize.Text)
    If IsHashed Then
        If (Capacity = 32 And GrowRate = 2 And GrowSize = 0) Then
            'Default-Werte
            Set m_List = MNew.List(Of_Type, , IsHashed)
        ElseIf (Capacity > 0 And Capacity <> 32) Then
            Set m_List = MNew.List(Of_Type, , IsHashed, Capacity)
        End If
        If GrowSize > 0 And GrowSize <> 2 Then
            Set m_List = MNew.List(Of_Type, , IsHashed, , , GrowSize)
        Else
            Set m_List = MNew.List(Of_Type, , IsHashed, Capacity, GrowRate, GrowSize)
        End If
    Else
        If (Capacity = 32 And GrowRate = 2 And GrowSize = 0) Then
            Set m_List = MNew.List(Of_Type, , IsHashed)
        Else
            Set m_List = MNew.List(Of_Type, , IsHashed, Capacity, GrowRate, GrowSize)
        End If
    End If
    EnableCtrls True
    AddRandom Count
    List1.Clear
    
    If Count < 50000 Then
        UpdateView
    Else
        UpdateView True
        If MsgBox("This may take a while, to fill the listbox, do you really want to proceed?", vbOKCancel) = vbOK Then
            UpdateView
        End If
    End If
End Sub
 
Sub AddRandom(nCount As Long)
    Randomize
    Dim dt As EDataType: dt = m_List.DataType
    Dim i As Long, u As Long: u = IIf(nCount, nCount - 1, 10 + Rnd * 40) 'zw 10 und 50 Elemente
    Dim Arr
    Select Case dt
    Case vbInteger:  ReDim Arr(0 To u) As Integer
                     For i = 0 To u: Arr(i) = Rnd * 65535 - 32768: Next
    Case vbLong:     ReDim Arr(0 To u) As Long
                     For i = 0 To u: Arr(i) = (Rnd - 0.5) * 2 * 2147483647: Next
    Case vbSingle:   ReDim Arr(0 To u) As Single
                     For i = 0 To u: Arr(i) = (Rnd - 0.5) * 2 * 1000000: Next
    Case vbDouble:   ReDim Arr(0 To u) As Double
                     For i = 0 To u: Arr(i) = (Rnd - 0.5) * 2 * 100000000000#: Next
    Case vbCurrency: ReDim Arr(0 To u) As Currency
                     For i = 0 To u: Arr(i) = (Rnd - 0.5) * 2 * 100000000000#: Next
    Case vbDate:     ReDim Arr(0 To u) As Date
                     For i = 0 To u: Arr(i) = Now - CDate(Rnd * 100): Next
    Case vbBoolean:  ReDim Arr(0 To u) As Boolean
                     For i = 0 To u: Arr(i) = CBool(Max((Rnd - 0.5), 0)): Next
    Case vbByte:     ReDim Arr(0 To u) As Byte
                     For i = 0 To u: Arr(i) = (Rnd) * 255: Next
    Case vbDecimal:  ReDim Arr(0 To u)
                     For i = 0 To u: Arr(i) = CDec(Rnd * CDec(16777216)): Next
    Case vbString:   ReDim Arr(0 To u) As String
                     For i = 0 To u: Arr(i) = MNew.RndName: Next
    Case vbWChar:    Arr = "Dies ist ein Teststring"
    Case vbUserDefinedType:
                     ReDim Arr(0 To u) As GridSettingsType
                     For i = 0 To u: Arr(i) = MNew.GridSettingsTypeRnd: Next
    Case vbObject:   ReDim Arr(0 To u) As Person
                     For i = 0 To u: Set Arr(i) = MNew.PersonRnd: Next
    Case vbVariant:  Arr = Array(1, "eins", 123456789, 123.456798, PersonRnd, Now)
    End Select
    m_List.AddRange Arr
End Sub

Sub UpdateView(Optional OnlyInfo As Boolean = False)
    If m_List Is Nothing Then
        ButtonsEnabled
        Exit Sub
    End If
    If m_List.IsEmpty Then
        ButtonsEnabled
        BtnAdd.Enabled = True
        'Exit Sub
    End If
    Dim s As String
    With m_List
        s = s & .GetType & vbCrLf
        s = s & "Count:    " & .Count & vbCrLf
        s = s & "Capacity: " & .Capacity & vbCrLf
        s = s & "GrowRate: " & .GrowRate & vbCrLf
        s = s & "GrowSize: " & .GrowSize & vbCrLf
        s = s & "IsHashed: " & .IsHashed & vbCrLf
        s = s & "SAPtr:    " & .SAPtr & vbCrLf
        s = s & "DataPtr:  " & .DataPtr & vbCrLf
        s = s & "ByteLen:  " & .ByteLength & vbCrLf
        s = s & "UBound:   " & .UUBound & vbCrLf
        's = s & "DataType: " & .DataType
    End With
    Label6.Caption = s
    DoEvents
    If OnlyInfo Then Exit Sub
    
    List1.Visible = False
    List2.Visible = False
    
    With m_List
        Dim i As Long: i = List1.ListIndex
        List1.Clear
        .ToListbox List1
        If i < List1.ListCount Then
            List1.ListIndex = i
        End If
    End With
    If Not m_ListClone Is Nothing Then
        List2.Clear
        m_ListClone.ToListbox List2
    End If
    List1.Visible = True
    List2.Visible = True
End Sub

Private Function ButtonsEnabled(Optional Enabled)
'    If Not IsMissing(Enabled) Then
        ToggleEnabled BtnAdd, Enabled
        ToggleEnabled BtnDelete, Enabled
        ToggleEnabled BtnEdit, Enabled
        ToggleEnabled BtnSearch, Enabled
        ToggleEnabled BtnInsert, Enabled
        ToggleEnabled BtnClearAll, Enabled
        ToggleEnabled BtnMoveUp, Enabled
        ToggleEnabled BtnMoveDown, Enabled
        ToggleEnabled BtnSortUp, Enabled
        ToggleEnabled BtnSortDown, Enabled
        ToggleEnabled BtnClone, Enabled
        'ToggleEnabled BtnBack, Enabled
'    Else
'        BtnAdd.Enabled = Enabled
'        BtnDelete.Enabled = Enabled
'        BtnEdit.Enabled = Enabled
'        BtnSearch.Enabled = Enabled
'        BtnInsert.Enabled = Enabled
'        BtnClearAll.Enabled = Enabled
'        BtnMoveUp.Enabled = Enabled
'        BtnMoveDown.Enabled = Enabled
'        BtnSortUp.Enabled = Enabled
'        BtnSortDown.Enabled = Enabled
'        BtnClone.Enabled = Enabled
'        BtnBack.Enabled = Enabled
'    End If
End Function
Private Sub ToggleEnabled(Btn As CommandButton, Optional Enabled)
    Btn.Enabled = IIf(IsMissing(Enabled), Not Btn.Enabled, Enabled)
End Sub

Private Function Bol_Parse(ByVal s As String) As Boolean
Try: On Error GoTo Catch
    Bol_Parse = CBool(s)
Catch:
End Function
Private Function Lng_Parse(ByVal s As String) As Long
Try: On Error GoTo Catch
    Lng_Parse = CLng(s)
Catch:
End Function
Private Function Sng_Parse(ByVal s As String) As Single
Try: On Error GoTo Catch
    s = Replace(s, ",", ".")
    Sng_Parse = CSng(Val(s))
Catch:
End Function
Private Function Dbl_Parse(ByVal s As String) As Single
Try: On Error GoTo Catch
    s = Replace(s, ",", ".")
    Dbl_Parse = Val(s)
Catch:
End Function

' #################### '  Buttons  ' #################### '
Private Sub BtnAdd_Click()
Try: On Error GoTo Catch
    'adding a new element to the list
    Dim s As String: s = InputBox("Add new element: ", "Add")
    If Len(s) Then
        'acccording to datatype add the element
        m_List.Add s
        UpdateView
    End If
    Exit Sub
Catch:
    MsgBox Err.Description
End Sub
Private Sub BtnDelete_Click()
    'deleting the element at the current ListBox position
    Dim i As Long:   i = List1.ListIndex
    If i < 0 Then
        MsgBox "Please select one element first"
        Exit Sub
    End If
    Dim s As String: s = List1.List(i)
    If MsgBox("Do you really want to delete the element " & vbCrLf & s & vbCrLf & "at the position: " & i, vbOKCancel) = vbOK Then
        m_List.Remove i
        UpdateView
    End If
End Sub
Private Sub BtnEdit_Click()
Try: On Error GoTo Catch
    Dim i As Long:   i = LBGetListIndex(List1)
    Dim s As String
    Dim v
    Select Case m_List.DataType
    Case vbObject:  s = m_List.Item(i).ToStr
    Case vbVariant: 'v = m_List.Item(i)
                    If VarType(m_List.Item(i)) = vbObject Then
                        Set v = m_List.Item(i)
                        s = v.ToStr
                    End If
    Case Else:      s = m_List.Item(i)
    End Select
    s = InputBox("Index: " & i, "Element editieren: ", s)
    If s = vbNullString Then Exit Sub
    m_List.Item(i) = s
    
    'List1.List(i) = s
    
    UpdateView
    Exit Sub
Catch:
    MsgBox TypeName(Me) & "::BtnEdit_Click: " & vbCrLf & Err.Description
End Sub
Private Sub BtnSearch_Click()
    Dim s As String: s = InputBox("Element suchen: ")
    If s = vbNullString Then Exit Sub
    Dim i As Long: i = m_List.IndexOf(s)
    If i >= 0 Then LBSetListIndex List1, i
End Sub
Private Sub BtnInsert_Click()
Try: On Error GoTo Catch
    Dim i As Long:   i = List1.ListIndex
    If i < 0 Then
        MsgBox "zuerst Stelle markieren wo eingefügt werden soll"
        Exit Sub
    End If
    Dim s As String: s = m_List.ItemToStr(i)
    s = InputBox("New element: ", "Insert", s)
    If s = vbNullString Then Exit Sub
    Dim newEl
    If m_List.DataType = vbObject Then
        Set newEl = New Person: newEl.Parse s
    Else
        newEl = s
    End If
    m_List.Insert i, newEl
    UpdateView
    Exit Sub
Catch:
    MsgBox Err.Description
End Sub
Private Sub BtnClearAll_Click()
    m_List.Clear
    UpdateView
End Sub
Private Sub BtnMoveUp_Click()
    Dim i As Long: i = List1.ListIndex
    If Not (0 < i And i < m_List.Count) Then Exit Sub
    'm_List.Swap i, i - 1
    m_List.MoveUp i
    UpdateView
    List1.ListIndex = i - 1
End Sub
Private Sub BtnMoveDown_Click()
    Dim i As Long: i = List1.ListIndex
    'If i < 0 Then Exit Sub
    'If i >= (m_List.Count - 1) Then Exit Sub
    If Not (0 <= i And i < m_List.Count - 1) Then Exit Sub
    'm_List.Swap i, i + 1
    m_List.MoveDown i
    UpdateView
    List1.ListIndex = i + 1
End Sub
Private Sub BtnSortUp_Click()
    If m_List Is Nothing Then Exit Sub
    If m_List.Count = 0 Then Exit Sub
    m_List.Sort
    UpdateView
End Sub
Private Sub BtnSortDown_Click()
    If m_List Is Nothing Then Exit Sub
    If m_List.Count = 0 Then Exit Sub
    m_List.SortRev
    UpdateView
End Sub
Private Sub BtnClone_Click()
    'we are only cloning the list, for nailing the sequence. not the elemens in the list
    Set m_ListClone = m_List.Clone
    List2Show True
    BtnBack.Enabled = Not BtnBack.Enabled
    Me.Width = (Me.Width - Me.ScaleWidth) + List2.Left + List2.Width + 8 * Screen.TwipsPerPixelX
    UpdateView
End Sub

'Private Sub TxtCapacity_LostFocus()
'    If m_List Is Nothing Then Exit Sub
'    Dim s As String: s = TxtCapacity.Text
'    If IsNumeric(s) Then
'        Dim c As Long: c = Lng_Parse(s)
'        m_List.Capacity = c
'        UpdateView
'    Else
'        MsgBox "Please give numeric value for Capacity: " & s
'    End If
'End Sub
'
'Private Sub TxtGrowRate_LostFocus()
'    If m_List Is Nothing Then Exit Sub
'    Dim s As String: s = TxtGrowRate.Text
'    If IsNumeric(s) Then
'        Dim g As Single: g = Sng_Parse(s)
'        m_List.GrowRate = g
'        UpdateView
'    Else
'        MsgBox "Please give numeric value for GrowRate: " & s
'    End If
'End Sub

Sub List2Show(bShow As Boolean)
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim l As Single: l = List1.Left
    Dim t As Single: t = List1.Top
    Dim W As Single: W = List1.Width
    Dim H As Single: H = List1.Height
    If bShow Then
        l = l + brdr + W
        List2.ZOrder 0
    Else
        List1.ZOrder 0
    End If
    If W > 0 And H > 0 Then List2.Move l, t, W, H
End Sub
Private Sub BtnBack_Click()
    Set m_ListClone = Nothing
    List2Show False
    BtnBack.Enabled = Not BtnBack.Enabled
    Me.Width = (Me.Width - Me.ScaleWidth) + List2.Left + List2.Width + 8 * Screen.TwipsPerPixelX
End Sub

Private Sub Form_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim l As Single: l = List1.Left
    Dim t As Single: t = List1.Top
    Dim W As Single: W = List1.Width
    Dim H As Single: H = Me.ScaleHeight - List1.Top - brdr
    If W > 0 And H > 0 Then
        List1.Move l, t, W, H
        List2.Move l, t, W, H
    End If
    If BtnBack.Enabled Then
        l = l + W + brdr
        'W = L + W + 8 * Screen.TwipsPerPixelX
        If W > 0 And H > 0 Then List2.Move l, t, W, H
    End If
End Sub

Private Sub List1_DblClick()
    BtnEdit_Click
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Dim i As Long: i = List1.ListIndex
        If i < 0 Then Exit Sub
        m_List.Remove i
        List1.Clear
        m_List.ToListbox List1
        If List1.ListCount <= i Then i = List1.ListCount - 1
        List1.ListIndex = i
    End If
End Sub












Private Sub BtnTestWhere_Click()
'6 * Filtern:
'   Where, Distinct, Take, TakeWhile, Skip, SkipWhile
'
'2 * Projizieren:
'   Select, SelectMany
'
'2 * Verknüpfen:
'   Join, GroupJoin

    'Dim li As List
    Dim i As Long
    
    MsgBox "Filling a List with 100 integers from 1 to 100"
    Set m_List = MNew.List(vbInteger)
    For i = 1 To 100:        m_List.Add i:     Next
    List1.Clear:    m_List.ToListbox List1
    UpdateView

    Set m_List = m_List.Where(lambdas, "var_Mod_10_eq_0")
    MsgBox "Linq: List1(0-100).Where(var Mod 10 = 0).Count=" & m_List.Count
    List1.Clear:    m_List.ToListbox List1
    UpdateView
    
    Randomize
    
    MsgBox "Filling a List with 500 random strings"
    Set m_List = MNew.List(vbString)
    For i = 1 To 500: m_List.Add MNew.RndName: Next
    List1.Clear:    m_List.ToListbox List1
    UpdateView
    
    Set m_List = m_List.Where(lambdas, "var_beginswith_A")
    MsgBox "Linq: List2=List1(101 Rnd str).Where(var beginswith ""A"")" & vbCrLf & "List2.Count=" & m_List.Count & "; List2.Item(0)=" & m_List.Item(0)
    List1.Clear:    m_List.ToListbox List1
    UpdateView
    
    
End Sub

Private Sub BtnTestvbWChar_Click()
    Dim li As List: Set li = MNew.List(vbWChar, "ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    List1.Clear: li.ToListbox List1
    MsgBox CLStr(li)
End Sub


Sub TestStuff()
    'TestListObj
    'TestListStr
    ReDim arr1(0 To 10) As Long
    ReDim arr2(0 To 2) As Long
    
    arr1(0) = 123
    arr1(1) = 456
    arr1(2) = 789
    
    arr2 = arr1
    
    'Debug.Print UBound(arr2)
    
    Dim c As Integer
    
    'c = AscW("A")
    
    'MsgBox ChrW(c)
End Sub
Sub TestFillInteger()
'Test Size, GrowFactor und ChunkSize

    'nur Typ ohne Parameter:
    TestFill MNew.List(vbInteger, , Capacity:=4, GrowRate:=0, GrowSize:=10)
    
    'nur Size:
    TestFill MNew.List(vbInteger, , , 4)
    TestFill MNew.List(vbInteger, , , 1)
    TestFill MNew.List(vbInteger, , , 5)
    
    'Size und GrowFact
    TestFill MNew.List(vbInteger, , , 4, 0)
    TestFill MNew.List(vbInteger, , , 4, 1)
    TestFill MNew.List(vbInteger, , , 4, 1.5)
    TestFill MNew.List(vbInteger, , , 4, 2)
    TestFill MNew.List(vbInteger, , , 4, 4)
    
    'Size, GrowFact und ChunkSize
    TestFill MNew.List(vbInteger, , , 4, 0, 0)
    TestFill MNew.List(vbInteger, , , 4, 0, 1)
    TestFill MNew.List(vbInteger, , , 4, 0, 2)
    
    TestFill MNew.List(vbInteger, , , 4, 1, 0)
    TestFill MNew.List(vbInteger, , , 4, 1, 1)
    TestFill MNew.List(vbInteger, , , 4, 1, 2)
    
    TestFill MNew.List(vbInteger, , , 4, 2, 0)
    TestFill MNew.List(vbInteger, , , 4, 2, 1)
    TestFill MNew.List(vbInteger, , , 4, 2, 2)

End Sub
Sub TestFill(li As List)
    Dim i As Long
    For i = 0 To 20
        li.Add i
        Debug.Print "i: " & i & "; Count: " & li.Count & "; Ubound: " & li.UUBound
    Next
End Sub




Private Sub BtnTestSelect1_Click()
    Dim s As Object: Set s = lambdas
    'Code1:
    '======
    'string s1 = "1;2;3;4;5;6;7;8;9;10;11;12";
    'int[] ia = tointarray(s1, ';');
    'simple Solution:
    'string[] sa = value.Split(sep);
    'int[] ia = new int[sa.Length];
    'for (int i = 0; i < ia.Length; ++i)
    '{
    '    int j;
    '    string s = sa[i];
    '    if (int.TryParse(s, out j))
    '    {
    '        ia[i] = j;
    '    }
    '}
    '
    'Linq1:
    '======
    'string s1 = "1;2;3;4;5;6;7;8;9;10;11;12";
    'int[] ia = s1.Split(';').Select(s => Convert.ToInt32(s)).ToArray();
    'in VBC:
    Dim s1 As String: s1 = "1;2;3;4;5;6;7;8;9;10;11;12"
    MsgBox "s1 = ""1;2;3;4;5;6;7;8;9;10;11;12"""
    List1.Clear: List1.AddItem s1
    
    Dim ia() As Long: ia = SplitL(s1, ";").SSelect(s, "Convert_ToIntI32", vbLong).ToArray
    MsgBox "ia() As Long = SplitL(s1, "";"").SSelect(s, ""Convert_ToIntI32"", vbLong).ToArray"
    
    MsgBox "List = MNew.List(vbLong, ia())"
    Set m_List = MNew.List(EDataType.vbLong, ia)
    UpdateView
    
    'List1.Clear
    
    'hmm geht das auch wenn man direkt Conversion als Lambda verwendet?
    'Dim iu As Object: Set iu = VBA.Conversion   'Typen unverträglich
    'Dim iu As IUnknown: Set iu = VBA.Conversion 'Typen unverträglich
    'Dim iu: Set iu = VBA.Conversion             'Typen unverträglich
    'Dim iu: iu = VBA.Conversion                 'Typen unverträglich
    'nö geht leider nicht
    
    
End Sub

Private Sub BtnTestSelect2_Click()
    Dim s As Object: Set s = lambdas
    'Code2:
    '======
    'string[] names = {"Peter", "Paul", "Mary"};
    'Person[] people;
    '/*  I could do this but I'm wondering if there's a better way. */
    'List<Person> persons = new List<Person>();
    'foreach(string name in names)
    '{
    '    persons.Add(new Person(name));
    '}
    'people = persons.ToArray();
    '
    'Linq2:
    '======
    'string[] names = {"Peter", "Paul", "Mary"};
    'Person[] people = names.Select(s => new Person(s)).ToArray();
    'in VBC:
    Dim names()  As String:  names = ArrayS("Peter", "Paul", "Mary")
    MsgBox "names()  As String:  names = ArrayS(""Peter"", ""Paul"", ""Mary"")"
    Set m_List = MNew.List(vbString, names)
    UpdateView
    
    MsgBox "people() As Person: people = MNew.List(vbString, names).SSelect(s, ""NewPersonS"", vbObject).ToArray"
    Dim people() As Person: people = MNew.List(vbString, names).SSelect(s, "NewPersonS", vbObject).ToArray
    Set m_List = MNew.List(vbObject, people)
    UpdateView
    'Debug.Print UBound(people)
    'Debug.Print people(0).Name
    
End Sub

Sub TestLINQ3()

    'var values = new[] { 1, 2, 3, 4, 5, 6, 7, 8 };
    'var average = values.Skip(2).Take(5).Average();
    
    'var myList = new double[] {1,2,3}
    'var avg = myList.Where(i => i > 1 && i < 2).Avg();
    'double[] values = new[] { 1.0, 2.0, 3.14, 2.71, 9.1 };
    'double average = values.Where(x => x > 2.0 && x < 4.0).Average();
    

'double avg = array
'    .Skip (startIndex)
'    .Take (endIndex - startIndex + 1)
'    .Average();


End Sub


Sub TestListStr()
    Dim m_Persons As List
    Set m_Persons = MNew.List(vbString, True)
    Dim i As Long
    For i = 1 To 100
        m_Persons.Add RndName
    Next
    m_Persons.Sort
    For i = 0 To m_Persons.Count - 1
        List1.AddItem m_Persons.Item(i).ToStr
    Next
End Sub
Sub TestListObj()
    Dim m_Persons As List
    Set m_Persons = MNew.List(vbObject, True)
    Dim i As Long
    For i = 1 To 100
        m_Persons.Add MNew.Person(RndName, RndBirthD, RndArt)
    Next
    m_Persons.Sort
    For i = 0 To m_Persons.Count - 1
        List1.AddItem CPerson(m_Persons.Item(i)).ToStr
    Next
End Sub

Sub TestList1()
    Set m_List = New List
    m_List.New_ vbLong, True, 1000
    
    m_List.Add 1000
    m_List.Add 1001
    m_List.Add 1002
    
    m_List.Add 1001
    
    'debug.Print mylist.Count
End Sub
