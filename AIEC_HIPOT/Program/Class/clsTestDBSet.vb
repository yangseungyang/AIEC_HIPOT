Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Ivi.Visa.Interop
'---------------------------------------------------------------
'   clsTestItemSet
'   Copyright(c)  2016 Yokogawa Electric Corporation.
'
'   Comment:  
'   Modify : Ver 0.00 2016/11/01 新規作成
'---------------------------------------------------------------
Public Class clsTestDBSet

    Public TestMemories As Object = Nothing
    Public Test_Item As Object = Nothing

    Public Test_No(50) As String
    Public Test_Testmode(50) As String
    Public Test_TestVoltage(50) As String
    Public Test_LowerJudgement(50) As String
    Public Test_LowerCurr(50) As String
    Public Test_UpperCurr(50) As String
    Public Test_TimStat(50) As String
    Public Test_TestTime(50) As String
    Public Test_Freq(50) As String
    Public Test_VoltageRiseTime(50) As String
    Public Test_ContactCheck(50) As String
    Public Test_Scanner101(50) As String
    Public Test_Scanner102(50) As String
    Public Test_Scanner103(50) As String
    Public Test_Scanner104(50) As String


    Dim mAppPath As String = Mid(Application.StartupPath, 1, InStr(Application.StartupPath, "bin") - 1)

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
    End Sub

    Public Sub TesDBSet()

        Dim i As Integer

        Dim xlApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim xlBook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Try
            xlApp = CType(VB.CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            xlBook = xlApp.Workbooks.Open(mAppPath & "\TestDB\TestDB_AIEC.xls")
            xlSheet = CType(xlBook.Worksheets("Test"), Microsoft.Office.Interop.Excel.Worksheet)
            TestMemories = xlSheet.Range("A2:O100").Value
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet)

            xlApp.Quit()
            xlSheet = Nothing
            xlBook = Nothing
            xlApp = Nothing

            System.GC.Collect()

            For i = 1 To 50
                Test_No(i) = TestMemories(i, 1)
                Test_Testmode(i) = TestMemories(i, 2)
                Test_TestVoltage(i) = TestMemories(i, 3)
                Test_LowerJudgement(i) = TestMemories(i, 4)
                Test_LowerCurr(i) = TestMemories(i, 5)
                Test_UpperCurr(i) = TestMemories(i, 6)
                Test_TimStat(i) = TestMemories(i, 7)
                Test_TestTime(i) = TestMemories(i, 8)
                Test_Freq(i) = TestMemories(i, 9)
                Test_VoltageRiseTime(i) = TestMemories(i, 10)
                Test_ContactCheck(i) = TestMemories(i, 11)
                Test_Scanner101(i) = TestMemories(i, 12)
                Test_Scanner102(i) = TestMemories(i, 13)
                Test_Scanner103(i) = TestMemories(i, 14)
                Test_Scanner104(i) = TestMemories(i, 15)
                System.Windows.Forms.Application.DoEvents()
                If Test_No(i) = "#" Then
                    Exit For
                End If
            Next

            mSecWait(0.2)
        Catch ex As System.Exception

            mSecWait(0.2)
            Throw New System.ApplicationException("DB Read NG!!")

        End Try
    End Sub


End Class
