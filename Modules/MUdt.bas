Attribute VB_Name = "MUdt"
Option Explicit
'herein just 3 helper-functions for the udt for convenient-reasons

Function GridSettingsType_ToStr(t As GridSettingsType) As String
    Dim s As String: s = TypeName(t) & "{"
    With t
        s = s & "AllowColDragging := " & .AllowColDragging '& vbCrLf
        s = s & ", AllowColSorting := " & .AllowColSorting '& vbCrLf
        s = s & ", AllowDragAndSort := " & .AllowDragAndSort '& vbCrLf
        s = s & ", GridStyle := " & .GridStyle '& vbCrLf
        s = s & ", GridType := " & .GridType '& vbCrLf
        s = s & ", HeaderBold := " & .HeaderBold '& vbCrLf
        s = s & ", ShadeAltCols := " & .ShadeAltCols '& vbCrLf
        s = s & ", ShadeAltRows := " & .ShadeAltRows '& vbCrLf
    End With
    GridSettingsType_ToStr = s & "}"
End Function

Function GridSettingsType_TryParse(ByVal s As String, t_out As GridSettingsType) As Boolean
Try: On Error GoTo Catch
    If InStr(1, s, "GridSettingsType") = 0 Then Exit Function
    t_out = GridSettingsType_Parse(s)
    GridSettingsType_TryParse = True
Catch:
End Function

Function GridSettingsType_Parse(ByVal s As String) As GridSettingsType
    's = Trim(s)
    'Dim sa1() As String: sa1 = Split(s, "=")
    'Dim sv As String: sv = sa1(1)
    Dim t As GridSettingsType
    If InStr(1, s, "GridSettingsType") Then
        Dim sa() As String: sa = Split(s, "{")
        Dim sv As String: sv = sa(1)
        If Right(sv, 1) = "}" Then sv = Mid(sv, 1, Len(sv) - 1)
        Dim svs() As String: svs = Split(sv, ",")
        Dim va() As String
        If UBound(svs) = 7 Then
            Dim i As Long
            For i = 0 To UBound(svs)
                va = Split(svs(i), ":=")
                If UBound(va) > 0 Then
                    Select Case LCase(Trim(va(0)))
                    Case "allowcoldragging": t.AllowColDragging = Trim(va(1))
                    Case "allowcolsorting":  t.AllowColSorting = Trim(va(1))
                    Case "allowdragandsort": t.AllowDragAndSort = Trim(va(1))
                    Case "gridstyle":        t.GridStyle = Trim(va(1))
                    Case "gridtype":         t.GridType = Trim(va(1))
                    Case "headerbold":       t.HeaderBold = Trim(va(1))
                    Case "shadealtcols":     t.ShadeAltCols = Trim(va(1))
                    Case "shadealtrows":     t.ShadeAltRows = Trim(va(1))
                    End Select
                End If
            Next
        End If
    End If
    GridSettingsType_Parse = t
End Function

