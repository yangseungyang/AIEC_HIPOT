Option Strict On : Option Explicit On
'==========================================================
'Form Name  : Main_Form
'Function   : AIEC HIPOT
'Framwork   : VB2022 Framework4.8
'Reference  : 
'Memo       : 
'Issue      : 2025/02/13  Rev 0.00    Daehee.Kim       New Release

'==========================================================

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
Imports System.Net.Sockets
Imports System.Net
Imports VB = Microsoft.VisualBasic
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement


Public Class frmMain

    'Program Ver ==========================================
    'Private Const Version As String = "0.00"   '2025.02.13 Ver 0.00 KDH
    Private Const Version As String = "0.01"    '2025.08.11 Ver 0.01 KDH

    Dim mAppPath As String = Mid(Application.StartupPath, 1, InStr(Application.StartupPath, "bin") - 1)
    Private EtherComN As New EtherComN.Dll

    Dim mErrorComment As String
    Dim DB_Replace As Boolean

    Dim sInsModel As String = ""
    Dim sInSerial As String = ""
    Dim SetVFlag As String = ""

    Dim Systemname As String
    Dim Systemversion As String

    Dim sJisEnc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")

    'Class ================================================
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cSrv As cls110_ServerConnect
    Public myDataSet As clsTestDBSet
    Private myTos As clsTos93xx
    Dim cPsm As cls100_PLASMA_Order
    Dim dfrm As frm500_MessagePopUp
    Private myPlc As clsHiPot
    Dim cLnk As cls080_F3LinkToLE
    Dim cSys As cls102_SignalTower

    Private Const SerialSelfTest As String = "SelfTest"
    'Local Property =======================================
    Dim mSystemLog As String = ""
    Dim Manual As String = "N"
    Public SelfTest As String = "N"
    Private COND As String = ""
    Private myTOS93xxNo As String = "DN000939" 'USB Serial
    Public selftest_Fin As Boolean = False
    Public selftest_NG As Boolean = False
    'Const Process_Code As String = "FAM3_010"          'Ver 0.01 KDH DEL
    Dim RDATE_S As String = ""
    Dim RTIME_S As String = ""
    Private I As Integer
    Dim mInspectionResult As String = ""
    'Dim mInspectionTestItem As String = ""             'Ver 0.01 KDH DEL
    Dim mInspectionTestItem As String = "耐電圧検査"             'Ver 0.01 KDH ADD
    Dim mInspectionSpec As String = ""
    Dim mInspectionMeasure As String = ""
    Dim mPLASMABarcode As String
    Private myLbTestItem() As System.Windows.Forms.Label

    '----------------------------- Ver 0.01 KDH ADD Start --------------------------------
    Public Property TEST_MEASURE() As String
        Get
            TEST_MEASURE = mInspectionMeasure
        End Get
        Set(ByVal value As String)
            mInspectionMeasure = value
        End Set
    End Property

    Public Property TEST_SPEC() As String
        Get
            TEST_SPEC = mInspectionSpec
        End Get
        Set(ByVal value As String)
            mInspectionSpec = value
        End Set
    End Property
    '----------------------------- Ver 0.01 KDH ADD End --------------------------------

    Public ReadOnly Property TOS93xxNo() As String
        Get
            TOS93xxNo = myTOS93xxNo
        End Get
    End Property
    Public ReadOnly Property SYSTEM_LOG() As String
        Get
            SYSTEM_LOG = mSystemLog
        End Get
    End Property
    Public Property TEST_RESULT() As String
        Get
            TEST_RESULT = mInspectionResult
        End Get
        Set(ByVal value As String)
            mInspectionResult = value
        End Set
    End Property
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim dbName As String = ""
        Dim tbName As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim rField As String

        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        Dim scmd As String = ""
        Dim rCmd As String = ""

        'Application check ===============================
        If UBound(GetProcessesByName(GetCurrentProcess.ProcessName)) > 0 Then
            MessageBox.Show("이 프로그램은 이미 시작하고 있습니다", "Application start NG", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End If

        'Make of System Log =======================================================================
        MakeSystemLog()
        If mErrorComment <> "" Then
            MessageBox.Show(mErrorComment, "Make System Log", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End If

        'Time, Temp,Humi Check =======================================================================
        Dim msg As String
        msg = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm")
        frmDateCK.TextBox2.Text = msg
        frmDateCK.ShowDialog()
        If frmDateCK.DateCKResult = False Then MyBase.Close()

        'Program Name & Version Check ===============================
        ProgramVerCheck()
        SysLogWrite(True, "Program Version Check OK")

        'OperatorName,ID Check ===============================
        Try
            cADO = Nothing

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = mAppPath & "Files\EquipmentSetting.mdb"
            tbName = "tblDefNameEntry"
            cADO.TB_NAME = tbName
            rField = "OperatorName,Operatorid"

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT = 0 Then
                mErrorComment = "Unable to retrieve Operator information."
                Throw New ApplicationException(mErrorComment)
            End If

            TextBox_op_name.Text = sData(0)
            TextBox_op_id.Text = sData(1)

            SysLogWrite(True, "OpratorName, ID Read OK")
        Catch
            SysLogWrite(True, "OpratorName, ID Read NG")
            SysLogWrite(True, "OpratorName, ID Read NG")
            MessageBox.Show(mErrorComment, "OperatorName, ID read Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        Finally
            cADO = Nothing
        End Try

        'MSCODE.mdb Copy & Check ===============================
        Call Button_MS_Update_Click(sender, e)
        SysLogWrite(True, "MSCODE.mdb Update OK")

        'TestDB Read_Test_Item/ACW ===============================
        myTos = New clsTos93xx(myTOS93xxNo)

        myDataSet = New clsTestDBSet
        myDataSet.TesDBSet()
        SysLogWrite(True, "TestDB Read OK")

        myTos.Preset()
        SysLogWrite(True, "TOS Preset OK")

        COND = myTos.PortSendReceive("STAT:OPER:TEST:COND?")

        'Test Item Set =====================================
        ComboBox_TestItem_Listing()
        ComboBox2_selfItem_Listing()

        'self test check =====================================
        dbName = mAppPath & "Result\Result_Data\Selftest_Result.mdb"
        tbName = "tblRES"

        cADO = New cls010_MDB_Operation

        cADO.DB_PATH = dbName
        cADO.TB_NAME = tbName

        sSQL = "SELECT " & "*" & " FROM " & tbName & " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"
        cADO.DB_SQL = sSQL

        Call cADO.DB_READ_JET(fName, sData)

        If cADO.DB_DATA_COUNT = 0 Then
            ComboBox2_selfItem.Text = "Continue"
        Else
            For I = 1 To sData.Length - 2
                If sData(I) <> "Pass" Then
                    ComboBox2_selfItem.Text = "Continue"
                    Exit For
                Else
                    ComboBox2_selfItem.Text = "Pass"
                End If
            Next

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            If ComboBox2_selfItem.Text = "Pass" Then

                sSQL = "UPDATE " & tbName & " SET "
                sSQL &= "TestMode" & " = '" & "Normal" & "'"
                sSQL &= " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"

                cADO.DB_SQL = sSQL
                Call cADO.DB_EXEC_JET()
                selftest_Fin = True
            End If
        End If

        'Reset =====================================
        PLC_RUN("192.168.1.2", "01" & "BWRY" & "201" & ",016," & "0000000000000000")

    End Sub
    Private Sub frmMain_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        cSys = New cls102_SignalTower
        Call cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
        cSys = Nothing
        TextBox_BarCode.Focus()
    End Sub

    Private Sub frmMain_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        PLC_RUN("192.168.1.2", "01" & "BWRY" & "201" & ",016," & "0000000000000000")

    End Sub

    Private Sub MakeSystemLog()
        '==============================================================================================================
        ' Function : System log
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim res As String() = {}

        'Make Folder ==============================================================================
        Dim SearchArea As String = mAppPath & "System_Log"
        Dim SearchFolder As String = Format(DateTime.Now, "yyyy-MM")
        Dim SearchFile As String = Format(DateTime.Now, "yyyy-MM-dd")

        'Search date folder -----------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea
            cFop.FOLDER_NAME = SearchFolder

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        Dim sFolder As String = String.Empty
        For Each stFilePath As String In res
            sFolder &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then make folder -------------------------------------
        If sFolder = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.FOLDER_PATH = SearchArea & "\" & SearchFolder

                Call cFop.FOLDER_MAKE()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()

            Finally
                cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Search Log file ------------------------------------------- 
        mSystemLog = SearchArea & "\" & SearchFolder & "\" & Format(DateTime.Now, "yyyy-MM-dd") & ".txt"

        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\" & SearchFolder & "\"
            cFop.FILE_NAME = Format(DateTime.Now, "yyyy-MM-dd") & ".txt"
            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        Dim sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Log Write --------------------------------------------------
        If sFile = String.Empty Then
            Call Me.SysLogWrite(False, mSystemLog)
        End If

        Call Me.SysLogWrite(False, "")
        Call Me.SysLogWrite(True, "■　システム起動 ==========================================================")

    End Sub

    Public Sub SysLogWrite(ByVal tStmp As Boolean, ByVal msg As String)
        '==============================================================================================================
        ' Function : System Log write
        ' Data     : Time stapY/N , Message
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Try
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = mSystemLog

            Call cTxt.WRITE_TEXT_FILE(tStmp, msg)

        Finally
            cTxt = Nothing

        End Try

    End Sub

    Private Sub ProgramVerCheck()
        '==========================================================================================
        'Function : Check program version between Date in VB.NET and Systemsetting.mdb
        'Remarks  :
        '==========================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim rField As String = ""
        Dim kField As String = ""

        'Read of System setting ==================================================================
        Try
            dbName = mAppPath & "Files\SystemSetting.mdb"
            tbName = "tblSystemSetting"

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            rField = "SystemName,SystemVersion"
            kField = "SystemVersion"

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & kField & " in "
            sSQL = sSQL & "(SELECT " & "Max(" & kField & ") as maxver" & " FROM " & tbName & ")"
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)

            If Version <> sData(1) Then
                mErrorComment = "It's not the latest program."
                Throw New ApplicationException(mErrorComment)
            End If

            'Save the Value
            Me.Text = sData(0) & "  Ver : " & sData(1)
            Systemname = sData(0)
            Systemversion = sData(1)

        Catch ex As Exception
            If mErrorComment = "" Then
                mErrorComment = "Check the 'Systemsetting.mdb' "
            End If
            MessageBox.Show(mErrorComment & vbCrLf & "생산기술 호출바랍니다." & vbCrLf & "VB Version : " & Version & vbCrLf & "DB Version : " & sData(0), "Program Check NG")
            End
        Finally
            cADO = Nothing
        End Try
    End Sub

    Private Sub Button_MS_Update_Click(sender As Object, e As EventArgs) Handles Button_MS_Update.Click
        '==============================================================================================================
        ' Function : Get Server file
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim sSQL As String = ""
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0
        Dim mServerDrive As String = ""
        Dim mWarningComment As String = ""
        Dim mNormalComment As String = ""

        DB_Replace = True
        TextBox_MS.Text = "미 갱 신"
        mErrorComment = ""
        'Server Connect ===========================================================================
        Try
            cSrv = New cls110_ServerConnect

            cSrv.FORM_NAME = Me
            Call cSrv.ServerConnect()

            mServerDrive = cSrv.SERVER_DRIVE

        Catch ex As Exception
            mWarningComment = cSrv.ERR_COMMENT
            If mWarningComment = "" Then mWarningComment = ex.ToString()

            mServerDrive = ""

        Finally
            cSrv = Nothing

        End Try
        If mServerDrive = "" Then mWarningComment = "サーバーに接続できません。"

        If mWarningComment <> "" Then
            DB_Replace = False
            Call Me.SysLogWrite(True, "サーバー接続：" & mWarningComment)
            MessageBox.Show(mWarningComment, "MSCODE.mdb", MessageBoxButtons.OK, MessageBoxIcon.Error)

            TextBox_MS.Text = "미 갱 신"
            Exit Sub
        End If

        'Get of Common information ================================================================

        Dim oFileName() As String = {}
        Dim mFoldName() As String = {}
        Dim mFileName() As String = {}

        Dim mFileCount As Integer = 0

        ReDim oFileName(mFileCount)
        ReDim mFoldName(mFileCount)
        ReDim mFileName(mFileCount)

        'Path -------------------------------------------------------

        'oFileName(0) = Mid(mServerDrive, 1, 2) & "\AIEC_HIPOT\MSCODE.mdb "             'Ver 0.01 KDH DEL
        oFileName(0) = Mid(mServerDrive, 1, 2) & "\001 Run system\Common File\aiec\Common Data\MSCODE.mdb"             'Ver 0.01 KDH ADD
        mFoldName(0) = mAppPath & "Files\"
        mFileName(0) = "MSCODE.mdb"

        'Copy -------------------------------------------------------
        For i As Integer = 0 To mFileCount

            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = oFileName(i)
                cFop.FOLDER_PATH = mFoldName(i)
                cFop.FILE_NAME = mFileName(i)

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()
                Call Me.SysLogWrite(True, "共通ファイル取得：" & mErrorComment)

                mNormalComment = ""
                DB_Replace = False
                TextBox_MS.Text = "미 갱 신"
                Exit Sub
            Finally
                cFop = Nothing

            End Try

            If mErrorComment = "" Then
                DB_Replace = True
                TextBox_MS.Text = "갱신완료"
                Call Me.SysLogWrite(True, "共通ファイル取得：" & mFileName(i))
                ComboBox1.Focus()
            End If
        Next

        'Server Disconnect ========================================================================
        Try
            cSrv = New cls110_ServerConnect

            Call cSrv.ServerDisconnect()

        Catch ex As Exception
            DB_Replace = False
            mErrorComment = cSrv.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()
            Call Me.SysLogWrite(True, "サーバー切断失敗")
        Finally
            cSrv = Nothing

        End Try

        If mErrorComment <> "" Then
            Call Me.SysLogWrite(True, "サーバー切断：" & mErrorComment)
            DB_Replace = False
            TextBox_MS.Text = "미 갱 신"
        End If
    End Sub
    Private Sub ComboBox_TestItem_Listing()
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim tbName As String
        Dim rField As String
        Dim count As Integer

        Me.ComboBox_TestItem.DropDownStyle = ComboBoxStyle.DropDownList

        cADO = New cls010_MDB_Operation
        cADO.DB_PATH = mAppPath & "TestDB\TestDB_Model.mdb"
        tbName = "tblTest_Select"
        cADO.TB_NAME = tbName
        rField = "耐電圧型式"

        sSQL = "SELECT " & rField & " FROM " & tbName
        cADO.DB_SQL = sSQL

        Call cADO.DB_READ_JET(fName, sData)

        If cADO.DB_DATA_COUNT = 0 Then
            mErrorComment = "File確認：" & vbCrLf & mAppPath & "TestDB\TestDB_Model.mdb"

        End If

        count = 0
        Do While sData(count) <> "検査なし"
            Me.ComboBox_TestItem.Items.Add(sData(count))
            count += 1
        Loop

    End Sub
    Private Sub ComboBox2_selfItem_Listing()
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim tbName As String
        Dim rField As String
        Dim count As Integer

        Me.ComboBox2_selfItem.DropDownStyle = ComboBoxStyle.DropDownList

        cADO = New cls010_MDB_Operation
        cADO.DB_PATH = mAppPath & "TestDB\TestDB_Model2.mdb"
        tbName = "tblTest_Select"
        cADO.TB_NAME = tbName
        rField = "耐電圧型式"

        sSQL = "SELECT " & rField & " FROM " & tbName
        cADO.DB_SQL = sSQL

        Call cADO.DB_READ_JET(fName, sData)

        If cADO.DB_DATA_COUNT = 0 Then
            mErrorComment = "File確認：" & vbCrLf & mAppPath & "TestDB\TestDB_Model2.mdb"

        End If

        count = 0
        Do While sData(count) <> "検査なし"
            Me.ComboBox2_selfItem.Items.Add(sData(count))
            count += 1
        Loop

    End Sub

    Private Sub Button_ResultClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button_ResultClear.Click

        Me.TextBox_MSCode.Text = ""
        Me.TextBox_SerialNo.Text = ""
        Me.TextBox_OrderNo.Text = ""
        Me.TextBox_XjNo.Text = ""
        Me.TextBox_StartTime.Text = ""
        Me.TextBox_TestTime.Text = ""
        Me.TextBox_StartDate.Text = ""
        Me.TextBox_StationNo.Text = ""

        Me.ComboBox_TestItem.Text = ""
        Me.TextBox_BarCode.Text = ""
        Me.TextBox_Log.Text = "Log"

        '---------------------- Ver 0.01 KDH ADD Start -------------------
        cSys = New cls102_SignalTower
        Call cSys.SignalTower("YELLOW", "OFF", "OFF", "OFF")
        cSys = Nothing
        '---------------------- Ver 0.01 KDH ADD End ---------------------

        Me.TextBox_BarCode.Focus()

    End Sub
    Private Sub TextBox_BarCode_keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox_BarCode.KeyDown
        If e.KeyCode = Keys.Return Then
            '作業者情報更新
            If TextBox_BarCode.Text.Length = 8 Then 'And InStr(TextBox_BarCode.Text.ToString, "-") <> 0
                Call Op_Excange(TextBox_BarCode.Text)
                TextBox_BarCode.Text = ""
                Me.TextBox_BarCode.Focus()
                Exit Sub
            End If

            'Test EXCUTE ===============================
            If TextBox_BarCode.Text <> "" Then
                Me.Execution(sender, e)
            End If

            ' Click Clear Button ===============================
            Me.Button_ResultClear_Click(sender, e)

        End If

    End Sub

    Private Function Op_Excange(ByVal OpID As String) As Boolean
        Dim mPlasmaResult As String = ""
        Dim mErrorCode As String = ""
        Dim sSQL As String = ""

        Dim tbName As String

        Try
            sSQL = OpID & "-E"
            mPlasmaResult = EtherComN.ComHost("10.7.15.254", "10.7.15.140", 2031, "EMP-" & sSQL)

            Debug.Print(mPlasmaResult)

            Select Case Mid(mPlasmaResult, 1, 8)
                Case "ERROR:A0"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "生産情報がありません。このバ-コ-ドは無効です。"
                Case "ERROR:A1"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "生産情報がｷｬﾝｾルされているか。無効の着工札です。"
                Case "ERROR:A2" : Case "ERROR:A3" : Case "ERROR:A4"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "プログラム エラーです。"
                Case "ERROR:A5"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "ＯＲＡＣＬＥエラーです。"
                Case "ERROR:13"
                    mErrorCode = mPlasmaResult
                    mErrorComment = "タイムアウトエラーです。"
                Case ""
                    mErrorCode = mPlasmaResult
                    mErrorComment = "データがありません。"
            End Select
            If mErrorCode <> "" Then Throw New ApplicationException(mErrorComment)

            Op_Excange = True

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)

        End Try

        TextBox_op_name.Text = mPlasmaResult
        TextBox_op_id.Text = OpID

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = mAppPath & "Files\EquipmentSetting.mdb"
            tbName = "tblDefNameEntry"
            cADO.TB_NAME = "tblDefNameEntry"

            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "OperatorName = '" & TextBox_op_name.Text & "' , "
            sSQL &= "Operatorid = '" & OpID & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)

            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

    End Function

    Private Sub Execution(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Manual_Name As String = ""
        Dim rDTkey As String = ""
        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""
        Dim rField As String = ""
        Dim sData() As String = {""}
        Dim fName() As String = {""}

        'Dim EF1_Flag As String = ""            'Ver 0.01 KDH DEL
        Dim AteNum As String = ""               'Ver 0.01 KDH ADD
        Dim Test_Stop_Flag As String = ""
        Dim Test_Stop_MSG As String = ""

        Dim Test_Model As String = ""

        cLnk = New cls080_F3LinkToLE
        cPsm = New cls100_PLASMA_Order

        RDATE_S = Format(Now, "yyyy-MM-dd")
        RTIME_S = Format(Now, "H:mm:ss")

        Me.TextBox_StartDate.Text = RDATE_S
        Me.TextBox_StartTime.Text = RTIME_S

        'PLASMAデータ読込み ==============================================
        Try
            Dim recv As String = ""
            Dim Index_No As String = TextBox_BarCode.Text

            cPsm.BARCODE_TEXT = TextBox_BarCode.Text
            rDTkey = TextBox_BarCode.Text
            Call cPsm.GET_ORDER_INFORMAION()

            Me.TextBox_MSCode.Text = cPsm.MSCODE
            Me.TextBox_SerialNo.Text = cPsm.SERIAL
            Me.TextBox_OrderNo.Text = cPsm.ORDER
            Me.TextBox_XjNo.Text = cPsm.XJNO
            sInsModel = cPsm.MSCODE

            'Temp Dataの更新(Write of data)
            Call WriteData("RunMode", "AUT")
            Call WriteData("MSCODE", cPsm.MSCODE)
            Call WriteData("SerialNo", cPsm.SERIAL)

            'Traceability Data Save ==================================================
            Traceability_Data_Update(TextBox_BarCode.Text)

        Catch oExcept As Exception
            mErrorComment = cPsm.ERR_COMMENT : If mErrorComment = "" Then oExcept.ToString()
            Err_Close("PLASMA DATA READ NG " & mErrorComment)
            Exit Sub
        Finally
            cLnk = Nothing
            cPsm = Nothing
        End Try

        ' ==============================================================
        '入力間違え検出
        If TextBox_BarCode.Text = "" And TextBox_MSCode.Text = "" Then
            Err_Close("There is no barcode data and manual input data.")
            Exit Sub
        End If
        If Manual = "N" And DB_Replace = False Then
            Err_Close("MSCODE.mdb Update NG!! Please try to update the MSCODE.mdb. ")
            Exit Sub
        End If

        Me.TextBox_MSCode.Text = sInsModel

        ' 検査中止フラグ確認 ==============================================================
        'MSCODE.mdbファイルの読み込み
        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = mAppPath & "Files\MSCODE.mdb"
            tbName = "tblMSCODE"
            cADO.TB_NAME = tbName
            'rField = "EF1_D0042,検査中止フラグ,検査中止メッセージ,耐電圧型式, 耐電圧メッセージ"             'Ver 0.01 KDH DEL
            rField = "AteNum, 検査中止フラグ,検査中止メッセージ,耐電圧型式"             'Ver 0.01 KDH ADD

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model = '" & sInsModel & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT = 0 Then
                mErrorComment = "MSCode.mdb could not retrieve the data"
                MsgBox(mErrorComment, MsgBoxStyle.OkOnly, "MSCODEの検査中止Flag")
                Exit Sub
            End If

            'MSCODE.mdb Search Result
            'EF1_Flag = sData(0) : Test_Stop_Flag = sData(1) : Test_Stop_MSG = sData(2)         'Ver 0.01 KDH DEL
            AteNum = sData(0) : Test_Stop_Flag = sData(1) : Test_Stop_MSG = sData(2)         'Ver 0.01 KDH ADD
            Test_Model = sData(3)

            Dim StationNo As String
            'StationNo = sData(4).Replace("号機", "")             'Ver 0.01 KDH DEL
            StationNo = String.Join(", ", sData(0).Select(Function(c) c & "号機"))             'Ver 0.01 KDH ADD
            TextBox_StationNo.Text = StationNo

            '耐電圧型式確認
            If Test_Model = "別" Or Test_Model = "" Then
                MsgBox("耐電圧検査対象ではありません。" & vbCrLf & "Please check the model", MsgBoxStyle.OkOnly, "耐電圧型式")
                Exit Sub
            End If

            '検査中止フラグ確認
            If Manual = "N" And Test_Stop_Flag = "True" Then
                MsgBox("検査中止Flagがあります" & vbCrLf & Test_Stop_MSG, MsgBoxStyle.OkOnly, "検査中止Flag")
                Exit Sub
                '--------------------- Ver 0.01 KDH DEL Start ---------------------
                'ElseIf Manual <> "N" And EF1_Flag = "False" Then
                '    MsgBox("EF1_D0042に無いモデル名です。", MsgBoxStyle.OkOnly, "EF1_D0042_Flag")
                '    Exit Sub
                '--------------------- Ver 0.01 KDH DEL End ---------------------
            End If

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Err_Close(mErrorComment)
            Exit Sub
        Finally
            cADO = Nothing

        End Try

        If Test_Model <> "検査なし" Then
            For I = 0 To ComboBox_TestItem.Items.Count - 2
                If Test_Model = ComboBox_TestItem.Items(I).ToString Then
                    ComboBox_TestItem.SelectedItem = Test_Model
                    Exit For
                Else
                    ComboBox_TestItem.SelectedItem = Nothing
                End If
            Next
        Else
            Product_Info_Main(TextBox_BarCode.Text)
            Make_ResultLog_Main()
        End If

        If ComboBox_TestItem.SelectedItem Is Nothing Then
            'MsgBox("設備に登録されていないMSCODEです。")         'Ver 0.01 KDH DEL
            MsgBox("設備に登録されていない検査です。")          'Ver 0.01 KDH ADD
            Exit Sub
        End If
        ErrorData_Upload()

    End Sub
    Private Sub Make_ResultLog_Main()
        '==============================================================================================================
        ' Function : System log
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim res As String() = {}
        Dim m As System.IO.StreamWriter

        'Make Folder ==============================================================================
        Dim SearchArea As String = mAppPath & "\Result\Result_Log"
        Dim SearchFolder As String = Format(DateTime.Now, "yyyy-MM")
        Dim SearchFile As String

        'Search date folder -----------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea
            cFop.FOLDER_NAME = SearchFolder

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        Dim sFolder As String = String.Empty
        For Each stFilePath As String In res
            sFolder &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then make folder -------------------------------------
        If sFolder = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.FOLDER_PATH = SearchArea & "\" & SearchFolder

                Call cFop.FOLDER_MAKE()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString()

            Finally
                cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Search Log file ------------------------------------------- 
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\" & SearchFolder & "\"
            SearchFile = Format(DateTime.Now, "yyyyMMdd_HHmmss") & "_" & TextBox_SerialNo.Text & ".txt"
            cFop.FILE_NAME = SearchFile
            Call cFop.FILE_SEARCH(res)

            Try
                m = New System.IO.StreamWriter(SearchArea & "\" & SearchFolder & "\" & SearchFile, True, sJisEnc)
                m.WriteLine(TextBox_Log.Text)

            Catch ex As Exception
                mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                Debug.Print(mErrorComment)

                Throw New ApplicationException(mErrorComment)

            Finally
                m.Close()

            End Try

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString()

        Finally
            cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

    End Sub

    Private Sub Test_LogMsgChanged_Main(ByVal tStmp As Boolean, ByVal msg As String)

        'Log
        Dim tNow As String
        If tStmp = True Then
            tNow = "[" & Format(DateTime.Now, "HH:mm:ss") & "] "
        Else
            tNow = ""
        End If


        TextBox_Log.Text &= VB.vbCrLf & tNow & msg
        TextBox_Log.SelectionStart = TextBox_Log.Text.Length
        TextBox_Log.Focus()
        TextBox_Log.ScrollToCaret()

        System.Windows.Forms.Application.DoEvents() 'Redraw

    End Sub
    Private Sub Product_Info_Main(ByVal Barcode As String)
        Test_LogMsgChanged_Main(True, "")
        Test_LogMsgChanged_Main(False, "System = " & Text)
        Test_LogMsgChanged_Main(False, "***************")
        Test_LogMsgChanged_Main(False, System.DateTime.Now.ToString("yyyy-MM-dd HH:mm"))
        Test_LogMsgChanged_Main(False, "Temperature = " & TextBox_temp.Text)
        Test_LogMsgChanged_Main(False, "Humidity = " & TextBox_humi.Text)
        Test_LogMsgChanged_Main(False, "Operator = " & TextBox_op_name.Text)
        Test_LogMsgChanged_Main(False, "BAR = " & Barcode)
        Test_LogMsgChanged_Main(False, "MSCode = " & TextBox_MSCode.Text)
        Test_LogMsgChanged_Main(False, "Serial = " & TextBox_SerialNo.Text)
        Test_LogMsgChanged_Main(False, "OrderNo = " & TextBox_OrderNo.Text)
        Test_LogMsgChanged_Main(False, "検査なし")
        Test_LogMsgChanged_Main(False, "PASS")
        Test_LogMsgChanged_Main(False, "***************")

        Me.TextBox_BarCode.Focus()
    End Sub

    Private Sub WriteData(ByVal fName As String, ByVal sData As String)
        '==========================================================================================
        ' Function : Write data(Temporary.mdb)
        ' Data     : Field name , data
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""

        dbName = mAppPath & "Result\Result_Data\Temporary.mdb"
        tbName = "tblResult"

        'Write of data ----------------------------------------------
        sSQL = "UPDATE " & tbName & " SET "
        sSQL &= fName & " = '" & sData & "'"

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try

    End Sub
    Public Sub Err_Close(ByVal msg As String)
        '==========================================================================================
        ' Function : Write of message
        ' Data     : Time stamp Y/N , Message
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Try
            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "PPT\NGWin.ppt"
            dfrm.Button2.Visible = False
            dfrm.Button1.Left = CInt(dfrm.Width / 2) - CInt(dfrm.Button1.Width / 2)
            dfrm.LED_COLOR = "RED" : dfrm.BUZZ1 = "ON" : dfrm.BUZZ2 = "OFF" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                If dfrm.DIALOG_RESULT <> "" Then
                    Exit Do
                End If
                DoEvents()
            Loop

        Catch ex As Exception
            MsgBox("에러 PPT 출력이상, 생산설계 담당자에게 문의 하세요.")
        Finally
            dfrm.Close()
            dfrm = Nothing

            'WRITE_TEXT_FILE(False, "Test Fail : " & msg)
            Call WriteData("Result", "NG")
            Call WriteData("ResultComment", msg)

            'Msg_BOX 確認
            MsgBox(msg, MsgBoxStyle.OkOnly, "NG!!!!")

            'Temp Dataの更新(Write of data)
            Call WriteData("MSCODE", "")
            Call WriteData("SerialNo", "")

            Manual = "N"
        End Try

    End Sub
    Private Sub Traceability_Data_Update(ByVal barcode As String)
        '==============================================================================================================
        ' Function : Traceability_Data save
        ' Data     : Temperary Colum
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""
        Dim rField As String = ""

        Try
            cADO = New cls010_MDB_Operation

            dbName = mAppPath & "Result\Result_Data\Traceability_Temp.mdb"
            tbName = "tblResult"

            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= "BAR = '" & barcode & "' , "
            sSQL &= "SERIAL = '" & cPsm.SERIAL & "' , "
            sSQL &= "PROD_NO = '" & cPsm.PRDNO & "' , "
            sSQL &= "SAP_LINKAGE_NO = '" & cPsm.LINKG & "' , "
            sSQL &= "SAP_COMP_NO = '" & cPsm.COMP & "' , "
            sSQL &= "CRP_GR_NO = '" & cPsm.CORPG & "' , "
            sSQL &= "START_NO = '" & cPsm.START & "' , "
            sSQL &= "MS_CODE = '" & cPsm.MSCODE & "' , "
            sSQL &= "FINISH_SCHDULE_D = '" & cPsm.FINSC & "' , "
            sSQL &= "INST_FINISH_D = '" & cPsm.INSFN & "' , "
            sSQL &= "XJ_NO = '" & cPsm.XJNO & "' , "
            'sSQL &= "PROCESS_CODE = '" & Process_Code & "' , "         'Ver 0.01 KDH DEL
            sSQL &= "RDATE_S = '" & RDATE_S & "' , "
            sSQL &= "RTIME_S = '" & RTIME_S & "' , "
            sSQL &= "WORKERID = '" & TextBox_op_id.Text & "' , "
            sSQL &= "INDEX_NO = '" & cPsm.Index & "' , "
            sSQL &= "BYORD_START_NO = '' , "
            sSQL &= "PC_Name = '" & System.Net.Dns.GetHostName & "' , "
            sSQL &= "SOFT = '" & Systemname & "' , "
            sSQL &= "VER = '" & Systemversion & "'"

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then
                mErrorComment = ex.ToString
            End If
            Err_Close(mErrorComment)
        Finally
            cADO = Nothing
        End Try
    End Sub

    Private Sub Button_TestEnd_Click(sender As Object, e As EventArgs) Handles Button_TestEnd.Click
        PLC_RUN("192.168.1.2", "01" & "BWRY" & "201" & ",016," & "0000000000000000")
        System.Windows.Forms.Application.Exit()

    End Sub


    Private Sub ComboBox_TestItem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox_TestItem.SelectedIndexChanged
        Dim dbName As String = mAppPath & "Result\Result_Data\Selftest_Result.mdb"
        Dim tbname As String = "tblRES"
        Dim sSQL As String = ""
        Dim rField As String = ""
        Dim sData() As String = {""}
        Dim fName() As String = {""}
        Dim Test_Model As String = ""
        Dim Manual_Check As Integer
        Dim str As String = sInsModel

        If ComboBox_TestItem.Text = "" Then Exit Sub

        myPlc = New clsHiPot

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbname

            sSQL = "SELECT " & "*" & " FROM " & tbname & " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT = 0 Then
                MsgBox("시업점검이 완료되지 않았습니다. 시업점검을 실시해 주세요.")
                Exit Sub
            End If

            For I = 1 To sData.Length - 2
                If sData(I) <> "Pass" Or sData(sData.Length - 1) <> "Normal" Then
                    MsgBox("시업점검이 완료되지 않았습니다. 시업점검을 실시해 주세요.")
                    Exit Sub
                End If
            Next

            If TextBox_BarCode.Text = "" Then Manual = "Y"

            If Manual = "Y" Then
                Manual_Check = MsgBox("수동검사를 진행하시겠습니까?", MsgBoxStyle.YesNo, "수동검사 확인")
                If Manual_Check = 6 Then      'Yes Return
                Else
                    Exit Sub
                End If
            End If

            'PPT ==============================================================

            Try
                dfrm = New frm500_MessagePopUp
                dfrm.DISP_TITLENAME = "Inspection "
                If TextBox_BarCode.Text = "検査なし" Then
                    dfrm.DISP_FILENAME = mAppPath & "PPT\Connector\x.ppt"
                Else
                    If sInsModel <> "" Then
                        Dim pptmodel As String = str.Substring(0, 11)
                        dfrm.DISP_FILENAME = mAppPath & "PPT\Connector\" & pptmodel & ".ppt"
                    Else
                        dfrm.DISP_FILENAME = mAppPath & "PPT\Connector\" & ComboBox_TestItem.Text & ".ppt"
                    End If
                End If
                dfrm.LED_COLOR = "YELLOW" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "OFF" : dfrm.FLASH = "ON"

                Call dfrm.frmShow()
                Do
                    'wait push button -------------------
                    If dfrm.DIALOG_RESULT = "NG" Then
                        mErrorComment = "Setting NG "
                        Throw New ApplicationException(mErrorComment)
                    End If

                    If dfrm.DIALOG_RESULT = "OK" Then
                        Exit Do
                    End If
                    DoEvents()
                Loop
            Catch ex As Exception
                SysLogWrite(True, TextBox_MSCode.Text & "(" & TextBox_SerialNo.Text & ")" & " : " & mErrorComment)
                Exit Sub
            Finally
                dfrm.Close()
                dfrm = Nothing
            End Try

            myPlc.TestMain(ComboBox_TestItem.Text)

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
        Finally
            cADO = Nothing
        End Try

        Me.ComboBox_TestItem.Text = ""
        Manual = "N"
        Me.Button_ResultClear_Click(sender, e)
    End Sub

    Private Function TEMP_Read(ByVal Column As String) As String

        '検査型式の読込み
        Dim dbName As String = mAppPath & "Result\Result_Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = Column
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        TEMP_Read = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            TEMP_Read = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function
    Private Function PLC_RUN(ByVal IP As String, ByVal scmd As String) As Long
        Dim cLnk As New cls080_F3LinkToLE
        Dim rCmd As String = ""
        Dim mErrorComment As String

        cLnk.IPAdd = IP

        Try
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Exit Function
            End If
            PLC_RUN = 0
        Catch ex As Exception
            Throw New System.Exception("PLC Commnucation NG!")
        Finally
            cLnk = Nothing
        End Try

    End Function

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2_selfItem.SelectedIndexChanged

        Dim dbName As String = mAppPath & "Result\Result_Data\Selftest_Result.mdb"
        Dim tbname As String = "tblRES"
        Dim sSQL As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim kField As String = ""
        Dim SearchR As String = ""
        Dim InputSetReg2 As Single = 0.0
        Dim mOIC As String = ""
        Dim Test_Status As String = "Continue"

        If ComboBox2_selfItem.Text = "" Then
            Exit Sub
        End If
        selftest_Fin = False

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbname

            sSQL = "SELECT " & "*" & " FROM " & tbname & " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)


            If cADO.DB_DATA_COUNT = 0 Then
                sSQL = "INSERT INTO " & tbname & " ("
                sSQL = sSQL & "StartDate" & ", "
                sSQL = sSQL & "TestMode" & ", "
                sSQL = sSQL & ComboBox2_selfItem.Text
                sSQL = sSQL & ") VALUES("
                sSQL = sSQL & "'" & Format(DateTime.Now, "yyyy-MM-dd") & "', "
                sSQL = sSQL & "'" & "Self" & "', "
                sSQL = sSQL & "'" & "Fail" & "'"
                sSQL = sSQL & ")"
            Else

                sSQL = "UPDATE " & tbname & " SET "
                sSQL &= "TestMode" & " = '" & "Self" & "' , "
                sSQL &= ComboBox2_selfItem.Text & " = '" & "Fail" & "'"
                sSQL &= " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"
            End If

            cADO.DB_SQL = sSQL
            Call cADO.DB_EXEC_JET()

            cADO = Nothing

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString


        Finally
            cADO = Nothing
        End Try

        'PPT ==============================================================
        If ComboBox2_selfItem.Text = "Continue" Or ComboBox2_selfItem.Text = "Pass" Then
        Else

            Try
                dfrm = New frm500_MessagePopUp
                dfrm.DISP_TITLENAME = "Inspection "
                If ComboBox2_selfItem.Text = "検査なし" Then
                    dfrm.DISP_FILENAME = mAppPath & "PPT\Connector\x.ppt"
                Else
                    dfrm.DISP_FILENAME = mAppPath & "PPT\Connector\" & ComboBox2_selfItem.Text & ".ppt"
                End If
                dfrm.LED_COLOR = "YELLOW" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "OFF" : dfrm.FLASH = "ON"

                Call dfrm.frmShow()
                Do
                    'wait push button -------------------
                    If dfrm.DIALOG_RESULT = "NG" Then
                        mErrorComment = "Setting NG "
                        Throw New ApplicationException(mErrorComment)
                    End If

                    If dfrm.DIALOG_RESULT = "OK" Then
                        Exit Do
                    End If
                    DoEvents()
                Loop
            Catch ex As Exception
                SysLogWrite(True, TextBox_MSCode.Text & "(" & TextBox_SerialNo.Text & ")" & " : " & mErrorComment)
                Exit Sub
            Finally
                dfrm.Close()
                dfrm = Nothing
            End Try

            myPlc = New clsHiPot

            myPlc.TestMain(ComboBox2_selfItem.Text)
        End If

        Try
            If selftest_NG = False Then
                Exit Try
            End If
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbname


            sSQL = "UPDATE " & tbname & " SET "
            sSQL &= ComboBox2_selfItem.Text & " = '" & "Pass" & "'"
            sSQL &= " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"

            cADO.DB_SQL = sSQL
            Call cADO.DB_EXEC_JET()


            sSQL = "SELECT " & "*" & " FROM " & tbname & " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            For I = 1 To sData.Length - 2
                If sData(I) <> "Pass" Then
                    ComboBox2_selfItem.Text = "Continue"
                    Exit For
                Else
                    ComboBox2_selfItem.Text = "Pass"
                End If
            Next

            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbname

            If ComboBox2_selfItem.Text = "Pass" Then

                sSQL = "UPDATE " & tbname & " SET "
                sSQL &= "TestMode" & " = '" & "Normal" & "'"
                sSQL &= " WHERE StartDate = '" & Format(DateTime.Now, "yyyy-MM-dd") & "'"

                cADO.DB_SQL = sSQL
                Call cADO.DB_EXEC_JET()
                selftest_Fin = True
            End If


        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing
            Me.Button_ResultClear_Click(sender, e)
        End Try

    End Sub
    Private Sub ErrorData_Upload()
        Dim sSQL As String = ""
        Dim mErrorCode As String = ""

        Dim RDATE_E As String = Format(Now, "yyyy-MM-dd")
        Dim RTIME_E As String = Format(Now, "H:mm:ss")

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim fCount As Integer = 0
        Dim rCount As Integer = 0

        cADO = New cls010_MDB_Operation

        If Manual = "N" Then

            Try
                'Get of temporary data ====================================================================
                dbName = mAppPath & "Result\Result_Data\Traceability_Temp.mdb"
                tbName = "tblResult"

                sSQL = "SELECT * FROM " & tbName

                cADO = New cls010_MDB_Operation

                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(fName, sData)

                fCount = cADO.DB_FIELD_COUNT
                rCount = cADO.DB_RECORD_COUNT

                If rCount = 0 Or fCount = 0 Then
                    Err_Close("Traceability_Temp Data Nothing")
                End If
            Catch ex As Exception

            End Try

            If mInspectionResult = "PASS" Or mInspectionResult = "" Then
                Exit Sub
            Else
                Try

                    'dbName = "\\10.36.14.215\Fam3\Err_Data\" & Format(DateTime.Now, "yyyy-MM") & "_NG.mdb"             'Ver 0.01 KDH DEL
                    dbName = "\\10.36.14.211\aiec\Err_Data\" & Format(DateTime.Now, "yyyy-MM") & "_NG.mdb"             'Ver 0.01 KDH ADD
                    tbName = "tblResult"

                    sSQL = "INSERT INTO " & tbName & " ("
                    sSQL = sSQL & "StartDateTime,TestName,ResultComment,TestMeasure,TestSpec,MSCODE,SerialNo,Operator,Approver,SystemName"
                    sSQL = sSQL & ",RunMode,Result,XJNo,OrderNo,ItemNo,CompNo,SystemVersion,EquipmentNo,InspectionModel,EndDateTime,Linkage"
                    sSQL = sSQL & ") VALUES ("
                    sSQL = sSQL & "'" & sData(12) & " " & sData(13) & "' , "
                    sSQL = sSQL & "'" & mInspectionTestItem & "' , "
                    sSQL = sSQL & "'" & TEMP_Read("ResultComment") & "' , "
                    sSQL = sSQL & "'" & mInspectionMeasure & "' , "
                    sSQL = sSQL & "'" & mInspectionSpec & "' , "
                    sSQL = sSQL & "'" & sData(7) & "' , "
                    sSQL = sSQL & "'" & sData(1) & "' , "
                    sSQL = sSQL & "'" & sData(14) & "' , "
                    sSQL = sSQL & "'" & "---" & "' , "
                    sSQL = sSQL & "'" & sData(18) & "' , "
                    sSQL = sSQL & "'" & "AUT" & "' , "
                    sSQL = sSQL & "'" & "NG" & "' , "
                    sSQL = sSQL & "'" & sData(10) & "' , "
                    sSQL = sSQL & "'" & Mid(sData(3), 1, 10) & "' , "
                    sSQL = sSQL & "'" & Mid(sData(3), 11, 6) & "' , "
                    sSQL = sSQL & "'" & sData(4) & "' , "
                    sSQL = sSQL & "'" & sData(19) & "' , "
                    sSQL = sSQL & "'" & "1" & "' , "
                    sSQL = sSQL & "'" & sData(7) & "' , "
                    sSQL = sSQL & "'" & RDATE_E & " " & RTIME_E & "' , "
                    sSQL = sSQL & "'" & sData(3) & "'"
                    sSQL = sSQL & ");"

                    cADO = New cls010_MDB_Operation

                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName

                    cADO.DB_SQL = sSQL

                    Call cADO.DB_EXEC_JET()

                    'NG品の“StartDateTime、MSCODE、SerialNo”をmdbに保存、処理票印字SystemのSearch用
                    'dbName = "\\10.36.14.215\Fam3\Err_Data\TestSearch.mdb"         'Ver 0.01 KDH DEL
                    dbName = "\\10.36.14.211\aiec\Err_Data\TestSearch.mdb"         'Ver 0.01 KDH DEL
                    tbName = "tblDate"

                    sSQL = "INSERT INTO " & tbName & " (StartDateTime , MSCODE , SerialNo) VALUES ('" & Format(DateTime.Now, "yyyy-MM-dd") & " " & sData(13) & "' , '" & sData(7) & "' , '" & sData(1) & "')"
                    cADO.DB_SQL = sSQL

                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName
                    Call cADO.DB_EXEC_JET()

                Catch ex As Exception
                    Err_Close("Err Data Up Load NG")
                Finally
                    cPsm = Nothing
                    cADO = Nothing
                End Try
            End If

        End If

    End Sub
    Private Sub WaitTimer(ByVal tm As Double)
        '==========================================================================================
        ' Function : Wait timer
        ' Data     : tm : Wait time(Sec)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim sTime As Double
        Dim nTime As Double
        Dim eTime As Double

        sTime = Microsoft.VisualBasic.DateAndTime.Timer()

        Do While tm > eTime
            nTime = Microsoft.VisualBasic.DateAndTime.Timer()
            eTime = nTime - sTime
            If eTime < 0 Then eTime = 24.0# * 3600.0# - sTime + nTime
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

End Class