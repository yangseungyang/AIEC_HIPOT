Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouterLink
'Function   : 바코드 분배 프로그램(Barcode_Router) 연동 (A8)
'Framwork   : VB2022 Framework4.8
'Memo       : ----------------------------- Ver0.02 YSH ADD (클래스 전체 신규) -----------------------------
'             흐름 : Router → 가상 COM 으로 "바코드 + CR" 수신 → 이 프로그램 창을 앞으로 → BarcodeReceived 이벤트
'                    (frmMain 이 기존 TextBox_BarCode_keydown(Enter) 과 같은 처리 : 8자리 = 작업자 갱신, 그 외 = 검사)
'                    → 처리가 끝나면 "END + CR" 을 Router 로 회신
'             - 수신 COM : EquipmentSetting.mdb > tblEquipmentSetting > BarcodeReaderPort / BarcodeReaderSetting
'                          (기존 컬럼, FAM3 와 같은 형식 예 : COM9 / 9600-N-8-1)
'                          공란이면 COM 을 열지 않고 기존 HID(키보드) 입력만 동작한다 (기존 동작 유지)
'             - HID 로 받은 건은 END 를 보내지 않는다 (Router 를 거치지 않은 입력)
'             - END 를 보내기 전에 AllowSetForegroundWindow 로 Router 가 다시 맨 앞으로 올 수 있게 권한을 넘긴다
'             - 포트 설정 / 송신은 FAM3 와 같은 cls101_ComRS232C 를 쓴다
'             frmMain 의 연동 코드가 길어져 이 클래스로 분리 (2026-10-26 지시)
'Issue      : 2026/10/26  Ver0.02    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  6-2항, 7-7 ~ 7-10항 / A8_FAM3_Router_구조설계_v1.1.xlsx A8-02 ~ A8-07
'============================================================================================================================
Imports System.Runtime.InteropServices

Public Class clsRouterLink

    Public Event BarcodeReceived(ByVal sBarcode As String)     'Router 로 받은 바코드 한 건 (UI 스레드에서 발생)

    Private Const ROUTER_END As String = "END"                  'Router 로 보내는 검사 종료 신호 (구조설계 5_통신규약 P-04)
    Private Const ASFW_ANY As Integer = -1

    <DllImport("user32.dll")>
    Private Shared Function AllowSetForegroundWindow(ByVal dwProcessId As Integer) As Boolean
    End Function

    Private mForm As frmMain
    Private mAppPath As String
    Private WithEvents mPort As New System.IO.Ports.SerialPort  'Designer 를 건드리지 않기 위해 코드에서 만든다
    Private cBcd As cls101_ComRS232C
    Private mOpen As Boolean = False            'COM 을 열었으면 True
    Private mBusy As Boolean = False            '받은 건을 처리 중이면 True (처리 중 재진입 방지)
    Private mBuf As String = ""                 'CR 이 올 때까지 받은 문자를 모아 둔다
    Private ReadOnly mLock As New Object
    Private Delegate Sub LineDelegate(ByVal sBarcode As String)

    Public ReadOnly Property IsOpen() As Boolean
        Get
            Return mOpen
        End Get
    End Property

    Public Sub New(ByVal form As frmMain, ByVal appPath As String)
        mForm = form
        mAppPath = appPath
    End Sub

    Public Sub Open()
        '==========================================================================================
        'Function : Router 수신 COM 설정 / 열기 (구조설계 A8-04)
        '==========================================================================================
        Dim cADO As cls010_MDB_Operation = Nothing
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = mAppPath & "Files\EquipmentSetting.mdb"
            cADO.TB_NAME = "tblEquipmentSetting"
            cADO.DB_SQL = "SELECT BarcodeReaderPort , BarcodeReaderSetting FROM tblEquipmentSetting"
            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_RECORD_COUNT = 0 OrElse sData(0).Trim = "" Then
                mForm.SysLogWrite(True, "Router COM : not set (HID input only)")
                Exit Sub
            End If

            Dim BarSetting() As String = Split(sData(1).Trim, "-")

            cBcd = New cls101_ComRS232C
            cBcd.TERMINAL_CHAR = vbCr
            cBcd.RS_COMM = mPort
            cBcd.PORT = sData(0).Trim
            cBcd.SETTING = BarSetting(0) & "," & BarSetting(1) & "," & BarSetting(2) & "," & BarSetting(3)
            cBcd.PortSetup()

            mOpen = True
            mForm.SysLogWrite(True, "Router COM Open OK : " & sData(0).Trim & " " & sData(1).Trim)

        Catch ex As Exception
            mOpen = False
            mForm.SysLogWrite(True, "Router COM Open NG : " & ex.Message)
            MessageBox.Show("Router COM Open NG" & vbCrLf & ex.Message & vbCrLf & "HID 입력만 사용합니다.", "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Finally
            cADO = Nothing

        End Try

    End Sub

    Public Sub Close()
        '==========================================================================================
        'Function : Router 수신 COM 닫기 (구조설계 A8-07)
        '==========================================================================================
        Try
            If mOpen Then cBcd.PortClose()
        Catch
        End Try
        mOpen = False
    End Sub

    Private Sub mPort_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles mPort.DataReceived
        '==========================================================================================
        'Function : Router 수신 (수신 스레드) - CR 까지 모아서 한 건이 되면 UI 스레드로 넘긴다
        'Remarks  : Router 는 END 를 받을 때까지 다음 바코드를 보내지 않으므로 한 번에 한 건만 처리한다
        '==========================================================================================
        Dim sLine As String = ""

        Try
            SyncLock mLock
                mBuf &= mPort.ReadExisting()
                Dim p As Integer = mBuf.IndexOf(vbCr)
                If p < 0 Then Exit Sub
                sLine = mBuf.Substring(0, p).Trim()
                mBuf = ""
            End SyncLock

            If sLine = "" Then Exit Sub
            mForm.BeginInvoke(New LineDelegate(AddressOf HandleLine), sLine)

        Catch ex As Exception
            Debug.Print("Router Receive NG : " & ex.Message)

        End Try

    End Sub

    Private Sub HandleLine(ByVal sBarcode As String)
        '==========================================================================================
        'Function : Router 수신 처리 (UI 스레드) - 창을 앞으로 → BarcodeReceived → END 회신 (구조설계 A8-05, A8-06)
        '==========================================================================================
        If mBusy Then
            mForm.SysLogWrite(True, "Router Receive skip (busy) : " & sBarcode)
            Exit Sub
        End If

        mBusy = True
        Try
            mForm.SysLogWrite(True, "Router Receive : " & sBarcode)
            Call ShowFront()
            RaiseEvent BarcodeReceived(sBarcode)

        Catch ex As Exception
            mForm.SysLogWrite(True, "Router Receive NG : " & ex.Message)
            MessageBox.Show(ex.Message, "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Error)
            mForm.TextBox_BarCode.Text = ""

        Finally
            Call SendEnd()
            mBusy = False

        End Try

    End Sub

    Private Sub ShowFront()
        '==========================================================================================
        'Function : 검사 프로그램 창을 맨 앞으로 (숨김 / 최소화 해제) (md 7-7항)
        '==========================================================================================
        If mForm.WindowState = FormWindowState.Minimized Then mForm.WindowState = FormWindowState.Normal
        mForm.Show()
        mForm.TopMost = True
        mForm.TopMost = False
        mForm.Activate()
        mForm.TextBox_BarCode.Focus()
    End Sub

    Private Sub SendEnd()
        '==========================================================================================
        'Function : 검사 종료 신호 "END + CR" 을 Router 로 회신 (md 7-9항 / 구조설계 P-04, P-08)
        '==========================================================================================
        If Not mOpen Then Exit Sub

        Try
            Call AllowSetForegroundWindow(ASFW_ANY)
            cBcd.SendCmd = ROUTER_END
            cBcd.FuncSendCmd()
            mForm.SysLogWrite(True, "Router END Send OK")

        Catch ex As Exception
            mOpen = False                   'FuncSendCmd 는 오류 시 포트를 닫는다
            mForm.SysLogWrite(True, "Router END Send NG : " & ex.Message)

        End Try

    End Sub

End Class
