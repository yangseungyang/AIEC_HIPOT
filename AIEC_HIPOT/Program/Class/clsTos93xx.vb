'---------------------------------------------------------------
'   cls5080A
'   Copyright(c)  2012 Yokogawa Electric Corporation.
'
'   Comment: Class of 5080A 
'   Modify : Ver 0.00 2015.07.xx Initial Release for InternShip by Yong
'   Caution: This class use only for internship. Because this class only contain performance of Output.
'---------------------------------------------------------------
Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports VB = Microsoft.VisualBasic
Imports Ivi.Visa.Interop
Imports EtherComN.ClsBcom

Public Class clsTos93xx
    Private myPhaseLock As Boolean
    Private myOper As Boolean
    Private myVoltage As Double
    Private RM As New Ivi.Visa.Interop.ResourceManager       'USB Port 연결 참조
    Private TOS93xx As New Ivi.Visa.Interop.FormattedIO488   'USB Port 연결 참조

    Dim mErrorComment As String = ""
    Public Sub New(ByVal TOS93xxNo As String)
        MyBase.New()
        TOS93xx.IO = CType(RM.Open("USB0::0x0B3E::0x104F::" & TOS93xxNo & "::INSTR"), Ivi.Visa.Interop.IMessage)    '제품의 MAC ADDRESS 설정
        TOS93xx.WriteString("SYST:COMM:RLST REM")                                                                   'Remote상태로 설정
        myPhaseLock = False
        myOper = False
        myVoltage = 0
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Public Sub Preset()
        TOS93xx.WriteString("ABOR:TEST")                                                            'Stop
        TOS93xx.WriteString("*CLS")                                                                 'Clear
    End Sub
    Public Sub TosStop()
        TOS93xx.WriteString("ABOR:TEST")
    End Sub
    Public Function PortSendReceive(ByVal cmd As String) As String
        Dim recData As String
        Dim Indata As String = ""

        Try
            TOS93xx.WriteString(cmd)
            recData = TOS93xx.ReadString()
            Return recData
        Catch ex As Exception
            Throw New ApplicationException("[ERROR] SendReceive: " & ex.ToString)
        End Try
    End Function
    Public Sub Start()
        Try
            Me.SendOnly("INIT:TEST")
        Catch ex As System.Exception
            Throw New ApplicationException("[ERROR] SendReceive: " & ex.ToString)
        End Try
    End Sub
    Public Function SetTosStatus(ByVal Row As Integer) As String

        Dim i As Integer
        Dim ErrComment = ""
        Dim Testmode As String = frmMain.myDataSet.Test_Testmode(Row).ToString
        '----------------------------- Ver0.01 KDH DEL Start -----------------------------------
        'TestItemCheck() As String = {"FUNC?", Testmode & ":VOLT?", "SENS:" & Testmode & ":JUDG:LOW:STAT?", "SENS:" & Testmode & ":JUDG?", Testmode & ":VOLT:TIM:STAT?", Testmode & ":VOLT:TIM?", Testmode & ":VOLT:FREQ?", Testmode & ":VOLT:SWE:TIM?",
        '    "ROUT:" & Testmode & ":TERM:CCH?", "ROUT:" & Testmode & ":TERM? 101", "ROUT:" & Testmode & ":TERM? 102", "ROUT:" & Testmode & ":TERM? 103", "ROUT:" & Testmode & ":TERM? 104"}
        'Dim TestItemData(12) As String
        'Dim TestItemCheckSpec(12) As String
        '----------------------------- Ver0.01 KDH DEL End -----------------------------------
        '----------------------------- Ver0.01 KDH ADD Start -----------------------------------
        Dim TestItemCheck() As String = {"FUNC?", Testmode & ":VOLT?", "SENS:" & Testmode & ":JUDG:LOW:STAT?", "SENS:" & Testmode & ":JUDG?", Testmode & ":VOLT:TIM:STAT?", Testmode & ":VOLT:TIM?", Testmode & ":VOLT:FREQ?", Testmode & ":VOLT:SWE:TIM?",
            "ROUT:" & Testmode & ":TERM:CCH?", "ROUT:" & Testmode & ":TERM? 101", "ROUT:" & Testmode & ":TERM? 102", "ROUT:" & Testmode & ":TERM? 103", "ROUT:" & Testmode & ":TERM? 104", "SENS:" & Testmode & ":JUDG:LOW?"}
        Dim TestItemData(13) As String
        Dim TestItemCheckSpec(13) As String
        '----------------------------- Ver0.01 KDH ADD End -----------------------------------

        Try
            Me.SendOnly("FUNC " & Testmode)
            Me.SendOnly(Testmode & ":VOLT " & frmMain.myDataSet.Test_TestVoltage(Row).ToString)                       'TEST Voltage
            Me.SendOnly("SENS:" & Testmode & ":JUDG:LOW:STAT " & frmMain.myDataSet.Test_LowerJudgement(Row).ToString)      'Lower Pass OFF
            Me.SendOnly("SENS:" & Testmode & ":JUDG:LOW " & frmMain.myDataSet.Test_LowerCurr(Row).ToString)                'Lower Current Setting (mA)
            Me.SendOnly("SENS:" & Testmode & ":JUDG " & frmMain.myDataSet.Test_UpperCurr(Row).ToString)                    'Upper Current Setting (mA)
            Me.SendOnly("SENS:" & Testmode & ":VOLT:TIM:STAT " & frmMain.myDataSet.Test_UpperCurr(Row).ToString)           'Time Status         
            Me.SendOnly(Testmode & ":VOLT:TIM " & frmMain.myDataSet.Test_TestTime(Row).ToString)                      'Test Time (s)
            If Testmode = "ACW" Then
                Me.SendOnly(Testmode & ":VOLT:FREQ " & frmMain.myDataSet.Test_Freq(Row).ToString)                         'Test Freq(Hz)
            End If
            Me.SendOnly(Testmode & ":VOLT:SWE:TIM " & frmMain.myDataSet.Test_VoltageRiseTime(Row).ToString)           'Voltage Rise Time
            Me.SendOnly("ROUT:" & Testmode & ":TERM:CCH " & frmMain.myDataSet.Test_ContactCheck(Row).ToString)             'Contact Check Function (ON/OFF)
            Me.SendOnly("ROUT:" & Testmode & ":TERM 101," & frmMain.myDataSet.Test_Scanner101(Row).ToString)
            Me.SendOnly("ROUT:" & Testmode & ":TERM 102," & frmMain.myDataSet.Test_Scanner102(Row).ToString)
            Me.SendOnly("ROUT:" & Testmode & ":TERM 103," & frmMain.myDataSet.Test_Scanner103(Row).ToString)
            Me.SendOnly("ROUT:" & Testmode & ":TERM 104," & frmMain.myDataSet.Test_Scanner104(Row).ToString)
            '----------------------------- Ver0.02 YSH ADD Start -----------------------------------
            ' 신규 설비는 TOS9301 + TOS9320 ×3 (101~104 / 201~204 / 301~304) 이고 FAM3 프로그램과 같이 쓴다.
            ' A8 은 101~104 만 쓰므로, FAM3 검사 후 남아 있을 수 있는 201~304 를 OPEN 으로 만든다.
            ' 근거 : md Rev 3 6-7항 / 구조설계 A8-10 / FE1-AIEC002_1
            For Each sCh As String In {"201", "202", "203", "204", "301", "302", "303", "304"}
                Me.SendOnly("ROUT:" & Testmode & ":TERM " & sCh & ",OPEN")
            Next
            '----------------------------- Ver0.02 YSH ADD End -------------------------------------

            TestItemCheckSpec(0) = frmMain.myDataSet.Test_Testmode(Row)
            TestItemCheckSpec(1) = frmMain.myDataSet.Test_TestVoltage(Row)
            TestItemCheckSpec(2) = frmMain.myDataSet.Test_LowerJudgement(Row)
            TestItemCheckSpec(3) = frmMain.myDataSet.Test_UpperCurr(Row)
            TestItemCheckSpec(4) = frmMain.myDataSet.Test_TimStat(Row)
            TestItemCheckSpec(5) = frmMain.myDataSet.Test_TestTime(Row)
            TestItemCheckSpec(6) = frmMain.myDataSet.Test_Freq(Row)
            TestItemCheckSpec(7) = frmMain.myDataSet.Test_VoltageRiseTime(Row)
            TestItemCheckSpec(8) = frmMain.myDataSet.Test_ContactCheck(Row)
            TestItemCheckSpec(9) = frmMain.myDataSet.Test_Scanner101(Row)
            TestItemCheckSpec(10) = frmMain.myDataSet.Test_Scanner102(Row)
            TestItemCheckSpec(11) = frmMain.myDataSet.Test_Scanner103(Row)
            TestItemCheckSpec(12) = frmMain.myDataSet.Test_Scanner104(Row)
            TestItemCheckSpec(13) = frmMain.myDataSet.Test_LowerCurr(Row)           'Ver 0.01 KDH ADD

            frmMain.TEST_SPEC = (CDbl(TestItemCheckSpec(13)) * 1000).ToString & " ~ " & (CDbl(TestItemCheckSpec(3)) * 1000).ToString & "mA"          'Ver 0.01 KDH ADD

            'For i =0 To 12         'Ver 0.01 KDH DEL
            For i = 0 To 13         'Ver 0.01 KDH ADD
                If Testmode = "DCW" And i = 6 Then
                    Continue For
                End If
                TestItemData(i) = Me.PortSendReceive(TestItemCheck(i))
                Select Case i
                    'Case 1, 2, 3, 4, 5, 6, 7, 8            'Ver 0.01 KDH DEL 
                    Case 1, 2, 3, 4, 5, 6, 7, 8, 13         'Ver 0.01 KDH ADD
                        If CDbl(TestItemCheckSpec(i)) <> CDbl(TestItemData(i)) Then
                            ErrComment = VB.vbCrLf & "TOS Setting Error" & VB.vbCrLf & "Command : " & TestItemCheck(i) & VB.vbCrLf & "Row : " & Row & VB.vbCrLf & "DB Data : " & TestItemCheckSpec(i) & VB.vbCrLf & "TOS Status : " & CDbl(TestItemData(i))
                            Return ErrComment
                        End If
                    Case 0, 9, 10, 11, 12
                        If TestItemCheckSpec(i).Length = 4 Then
                            If VB.Right(TestItemCheckSpec(i), 4) <> VB.Left(TestItemData(i), 4) Then
                                ErrComment = "TOS Setting Error" & VB.vbCrLf & "Command : " & TestItemCheck(i) & VB.vbCrLf & "Row : " & Row & VB.vbCrLf & "DB Data : " & TestItemCheckSpec(i) & VB.vbCrLf & "TOS Status : " & TestItemData(i)
                                Return ErrComment
                            End If
                        ElseIf TestItemCheckSpec(i).Length = 3 Then
                            If VB.Right(TestItemCheckSpec(i), 3) <> VB.Left(TestItemData(i), 3) Then
                                ErrComment = "TOS Setting Error" & VB.vbCrLf & "Command : " & TestItemCheck(i) & VB.vbCrLf & "Row : " & Row & VB.vbCrLf & "DB Data : " & TestItemCheckSpec(i) & VB.vbCrLf & "TOS Status : " & TestItemData(i)
                                Return ErrComment
                            End If
                        End If
                    Case Else
                        frmMain.Err_Close("DB Data Error")
                        Return "Set Error"
                End Select

            Next

        Catch ex As System.Exception
            Throw New System.Exception(ex.ToString)
        End Try
        Return "Set"
    End Function
    Public Sub SendOnly(ByVal Data As String)

        Try
            TOS93xx.WriteString(Data)
        Catch ex As Exception
            Throw New ApplicationException("[ERROR] Send: " & ex.ToString)
        End Try

    End Sub
    Public Function GetCOND() As Double
        Dim ReceiveString As String = ""
        Try
            Me.PortSendReceive("STAT:OPER:TEST:COND?")
            Return CDbl(ReceiveString)
        Catch ex As System.Exception
            Throw New ApplicationException(String.Format("{0} : GetCOND に失敗しました。" & ex.ToString))
        End Try
    End Function
End Class
