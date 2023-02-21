
'============================================================================
' Name        : OpenFile.vb
' Author      : Abdurrahman Nurhakim
' Version     : 1.0
' Copyright   : Your copyright notice
' Description : Read Data from Excel 
'============================================================================

Imports Modbus.Device
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Collections.ObjectModel
Imports System.Runtime.Extensions
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Security.Policy

Public Class Form1

    Public Property collom As Integer
    Public Property row As Integer
    Public Property _second As String
    Public Property _minute As String
    Public Property _hour As String
    Public Property _year As String
    Public Property _month As String
    Public Property _date As String
    Public Property row_DP As Integer
    Public Property row_SP As Integer
    Public Property row_RTD As Integer
    Public Property statusDP As Boolean
    Public Property statusSP As Boolean
    Public Property statusRTD As Boolean
    Public Property unit As String
    Public Property num As String
    Public Property RTD As String
    Public Property SP As String
    Public Property DP As String
    Public Property filePath As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        openFileDialog1.Title = "Select an Excel file"

        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            filePath = openFileDialog1.FileName
        End If
    End Sub

    Public Function ParseDateToInt(dateString As String) As Integer()
        Dim Hasil(6) As Integer
        Dim format As String = "dd/MM/yyyy HH:mm:ss"
        Dim dateValue As DateTime
        dateValue = DateTime.ParseExact(dateString, format, CultureInfo.InvariantCulture)
        Hasil(0) = dateValue.Second
        Hasil(1) = dateValue.Minute
        Hasil(2) = dateValue.Hour
        Hasil(3) = dateValue.Year
        Hasil(4) = dateValue.Month
        Hasil(5) = dateValue.Day
        Return Hasil
    End Function

    Public Sub ReadUnitSensorSP(Input As String, CollomNumb As Integer)
        Dim Hasil As Boolean
        Dim Sbuff As String = "qwertyuiopasdfghasdfghjk"
        Dim myChar As Char() = (Input + Sbuff).ToCharArray()
        Dim Buff As String
        Dim size As Integer = Len(Input)

        For i = 0 To 16
            Buff += CStr(myChar(i))
        Next

        If Buff = "Upstream Pressure" Then
            statusSP = True
            row_SP = CollomNumb
        End If
    End Sub

    Public Sub ReadUnitSensorRTD(Input As String, CollomNumb As Integer)
        Dim Hasil As Boolean
        Dim Sbuff As String = "qwertyuiopasdfghasdfghjk"
        Dim myChar As Char() = (Input + Sbuff).ToCharArray()
        Dim Buff As String
        Dim size As Integer = Len(Input)

        For i = 0 To 18
            Buff += CStr(myChar(i))
        Next

        If Buff = "Process Temperature" Then
            statusRTD = True
            row_RTD = CollomNumb
        End If
    End Sub

    Public Sub ReadUnitSensorDP(Input As String, CollomNumb As Integer)
        Dim Hasil As Boolean
        Dim Sbuff As String = "qwertyuiopasdfghasdfghjk"
        Dim myChar As Char() = (Input + Sbuff).ToCharArray()
        Dim Buff As String
        Dim size As Integer = Len(Input)

        For i = 0 To 20
            Buff += CStr(myChar(i))
        Next

        If Buff = "Differensial Pressure" Then
            statusDP = True
            row_DP = CollomNumb
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        row = 0
        collom = 0
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Add("number1", "Number")
        DataGridView1.Columns.Add("number2", "DATE")
        DataGridView1.Columns.Add("number3", "DP")
        DataGridView1.Columns.Add("number4", "SP")
        DataGridView1.Columns.Add("number5", "RTD")
        DataGridView1.Columns.Add("number6", "UNIT")
        statusDP = False
        statusSP = False
        statusRTD = False
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(filePath)
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets("Sheet1")
        Dim range As Excel.Range = xlWorkSheet.UsedRange
        Dim numRows As Integer = range.Rows.Count
        Dim numCols As Integer = range.Columns.Count

        collom += 1
        unit = CStr(collom)
        ReadUnitSensorDP(range.Cells(1, collom).Value.ToString(), collom)
        ReadUnitSensorSP(range.Cells(1, collom).Value.ToString(), collom)
        ReadUnitSensorRTD(range.Cells(1, collom).Value.ToString(), collom)

        If statusDP = True And statusSP = True And statusRTD = True Then
            collom = 0
            If DataGridView1.Columns.Count = 0 Then
                DataGridView1.Columns.Add("number1", "Number")
                DataGridView1.Columns.Add("number2", "DATE")
                DataGridView1.Columns.Add("number3", "DP")
                DataGridView1.Columns.Add("number4", "SP")
                DataGridView1.Columns.Add("number5", "RTD")
            End If

            If row < numRows - 1 Then
                row += 1
            ElseIf row = numRows - 1 Then
                row = numRows - 1
            Else
                row = numRows - 1
            End If

            ' range.Cells(numRows, 1).Value.ToString()
            num = range.Cells(row + 1, 1).Value.ToString() 'number
            DP = range.Cells(row + 1, row_DP).Value.ToString()
            RTD = range.Cells(row + 1, row_RTD).Value.ToString()
            SP = range.Cells(row + 1, row_SP).Value.ToString()
            _date = range.Cells(row + 1, 2).Value.ToString()

            Dim buff(6) As String
            buff(0) = num
            buff(2) = DP
            buff(3) = SP
            buff(4) = RTD
            buff(1) = _date

            DataGridView1.Rows.Add(buff)
        End If

        range = Nothing

        xlApp.DisplayAlerts = False 'disable service alerts
        xlWorkSheet = Nothing
        xlWorkBook.Close(filePath) 'close axcel with service
        xlWorkBook = Nothing
        xlApp.Quit()
        xlApp = Nothing

        If row = numRows - 1 Then
            collom = 0
            statusDP = False
            statusSP = False
            statusRTD = False
            Timer1.Stop()
        End If
    End Sub

End Class