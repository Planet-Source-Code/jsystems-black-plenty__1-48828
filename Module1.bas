Attribute VB_Name = "Module1"
Public EUF As Single
Public ii As Single
Public INV As Single
Public OUT As Single
Public EVI As Boolean
Public ADOS As Single

Public Sub CRDB()
Dim fs As String
fs = App.Path & "\data.h"
If Dir(fs) = "" Then
    Dim DB As Database
    Set DB = CreateDatabase(fs, dbLangGeneral)
    DB.Execute "create table T1(Datum char(7), Leiras memo, Osszeg single)"
    DB.Execute "create table T2(Datum char(7), Leiras memo, Osszeg single)"
    DB.Execute "create table T3(Datum char(7), Leiras memo, Osszeg single)"
    DB.Execute "create table T4(Datum char(7), Leiras memo, Osszeg single)"
    Set DB = Nothing
End If
End Sub

Public Sub MakeInfo(ByVal DataC As String)
Screen.MousePointer = vbHourglass
Dim info1 As String
Dim info2 As String
Dim info3 As String
Dim info4 As String
Dim info5 As String

On Error Resume Next
frmMain.Data1.RecordSource = "T1"
frmMain.Data1.Refresh
Dim a1 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a1 = a1 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info1 = a1
INV = a1
frmMain.Data1.RecordSource = "T2"
frmMain.Data1.Refresh
Dim a2 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a2 = a2 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info2 = a2
OUT = a2
info3 = a1 - a2

frmMain.Data1.RecordSource = "T3"
frmMain.Data1.Refresh
Dim a3 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a3 = a3 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info4 = a3
ADOS = a3


frmMain.Data1.RecordSource = "T4"
frmMain.Data1.Refresh
Dim a4 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a4 = a4 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info5 = a4

frmMain.Label20.Caption = info1
frmMain.Label21.Caption = info2
frmMain.Label22.Caption = info3
frmMain.Label23.Caption = info4
frmMain.Label24.Caption = info5

'###make graph ide
frmMain.MSC.chartType = VtChChartType2dBar
frmMain.MSC.ColumnCount = 4
frmMain.MSC.RowCount = 1
frmMain.MSC.RowLabel = ""
frmMain.MSC.Column = 1
frmMain.MSC.ColumnLabel = "Bejovet"

frmMain.MSC.Row = 1
frmMain.MSC.Data = a1
frmMain.MSC.Column = 2
frmMain.MSC.ColumnLabel = "Kiadas"

frmMain.MSC.Row = 1
frmMain.MSC.Data = a2
frmMain.MSC.Column = 3
frmMain.MSC.ColumnLabel = "Adossag"

frmMain.MSC.Row = 1
frmMain.MSC.Data = a3
frmMain.MSC.Column = 4
frmMain.MSC.ColumnLabel = "Nekem tartoznak"

frmMain.MSC.Row = 1
frmMain.MSC.Data = a4
Screen.MousePointer = vbNormal
End Sub

Public Sub MakeEvi(ByVal DataC As String)
Screen.MousePointer = vbHourglass
Dim info1 As String
Dim info2 As String
Dim info3 As String
Dim info4 As String
Dim info5 As String

On Error Resume Next
frmMain.Data1.RecordSource = "T1"
frmMain.Data1.Refresh
Dim a1 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Right(Trim(frmMain.Data1.Recordset(0).Value), 4) = DataC Then
        a1 = a1 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info1 = a1

frmMain.Data1.RecordSource = "T2"
frmMain.Data1.Refresh
Dim a2 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Right(Trim(frmMain.Data1.Recordset(0).Value), 4) = DataC Then
        a2 = a2 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info2 = a2
info3 = a1 - a2

frmMain.Data1.RecordSource = "T3"
frmMain.Data1.Refresh
Dim a3 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Right(Trim(frmMain.Data1.Recordset(0).Value), 4) = DataC Then
        a3 = a3 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info4 = a3

frmMain.Data1.RecordSource = "T4"
frmMain.Data1.Refresh
Dim a4 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Right(Trim(frmMain.Data1.Recordset(0).Value), 4) = DataC Then
        a4 = a4 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info5 = a4

frmMain.Label20.Caption = info1
frmMain.Label21.Caption = info2
frmMain.Label22.Caption = info3
frmMain.Label23.Caption = info4
frmMain.Label24.Caption = info5

'###make graph ide
frmMain.MSC.chartType = VtChChartType2dLine
frmMain.MSC.ColumnCount = 3
frmMain.MSC.RowCount = 12

MakeInfoS "01/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 1
frmMain.MSC.RowLabel = "Jan"
frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 1
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 1
frmMain.MSC.Data = ADOS

MakeInfoS "02/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 2
frmMain.MSC.RowLabel = "Feb"
frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 2
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 2
frmMain.MSC.Data = ADOS


MakeInfoS "03/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 3
frmMain.MSC.RowLabel = "Marc"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 3
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 3
frmMain.MSC.Data = ADOS


MakeInfoS "04/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 4
frmMain.MSC.RowLabel = "Apr"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 4
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 4
frmMain.MSC.Data = ADOS


MakeInfoS "05/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 5
frmMain.MSC.RowLabel = "Maj"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 5
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 5
frmMain.MSC.Data = ADOS


MakeInfoS "06/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 6
frmMain.MSC.RowLabel = "Jun"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 6
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 6
frmMain.MSC.Data = ADOS


MakeInfoS "07/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 7
frmMain.MSC.RowLabel = "Jul"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 7
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 7
frmMain.MSC.Data = ADOS


MakeInfoS "08/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 8
frmMain.MSC.RowLabel = "Aug"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 8
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 8
frmMain.MSC.Data = ADOS


MakeInfoS "09/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 9
frmMain.MSC.RowLabel = "Szept"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 9
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 9
frmMain.MSC.Data = ADOS


MakeInfoS "10/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 10
frmMain.MSC.RowLabel = "Okt"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 10
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 10
frmMain.MSC.Data = ADOS


MakeInfoS "11/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 11
frmMain.MSC.RowLabel = "Nov"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 11
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 11
frmMain.MSC.Data = ADOS


MakeInfoS "12/" & DataC
frmMain.MSC.Column = 1
frmMain.MSC.Row = 12
frmMain.MSC.RowLabel = "Dec"

frmMain.MSC.Data = INV
frmMain.MSC.Column = 2
frmMain.MSC.Row = 12
frmMain.MSC.Data = OUT
frmMain.MSC.Column = 3
frmMain.MSC.Row = 12
frmMain.MSC.Data = ADOS


Screen.MousePointer = vbNormal
End Sub

Private Sub MakeInfoS(ByVal DataC As String)
Dim info1 As String
Dim info2 As String
Dim info3 As String
Dim info4 As String
Dim info5 As String

On Error Resume Next
frmMain.Data1.RecordSource = "T1"
frmMain.Data1.Refresh
Dim a1 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a1 = a1 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info1 = a1
INV = a1
frmMain.Data1.RecordSource = "T2"
frmMain.Data1.Refresh
Dim a2 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a2 = a2 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info2 = a2
OUT = a2
info3 = a1 - a2

frmMain.Data1.RecordSource = "T3"
frmMain.Data1.Refresh
Dim a3 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a3 = a3 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info4 = a3
ADOS = a3
frmMain.Data1.RecordSource = "T4"
frmMain.Data1.Refresh
Dim a4 As Single
frmMain.Data1.Recordset.MoveFirst
Do While Not frmMain.Data1.Recordset.EOF
    If Trim(frmMain.Data1.Recordset(0).Value) = DataC Then
        a4 = a4 + Val(Trim(frmMain.Data1.Recordset(2).Value))
    End If
frmMain.Data1.Recordset.MoveNext
Loop
info5 = a4


End Sub
