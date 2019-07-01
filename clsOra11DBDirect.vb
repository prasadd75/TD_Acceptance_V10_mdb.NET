Option Strict Off
Option Explicit On
Friend Class clsOra11DBDirect
	
	Public Connected As Boolean
	
	'///// TD���ް�X�V�ɔ������C  2013.03.15 /////////////////////////
	
	'Private Const msFILE_NAME = "clsPostgresSqlDBDirect"
	Private Const msFILE_NAME As String = "clsOra11DBDirect"
	
	'/////////////////////////////////////////////////////////////////
	
	'-------------------------------------------------------------------------------------
	'This is PostgreSQL
	
	'///// TD���ް�X�V�ɔ������C  2013.03.15 /////////////////////////
	
	'Public dbPgsqlFW As Connection
	Public dbOra11 As ADODB.Connection
	
	'/////////////////////////////////////////////////////////////////
	
	Private Const NUM_ATTEMPS As Short = 10 'ATTEMPT TO CONNECT TO THE DATABASE THIS NO. OF TIMES
	Private Const NO_ERROR As Short = 0 'NO ERROR OCCURRED
	Private Const ERROR_OCCURRED As Short = 1 'AN ERROR OCCURRED
	
	Public Function ConnectDb() As Short
		
		On Error GoTo ErrorHandler
		
		Dim strConnect As String
		Dim intAttempts As Short
		
		intAttempts = 1
		
		'UPGRADE_WARNING: Screen �v���p�e�B Screen.MousePointer �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'UPGRADE_ISSUE: vbNormal ���A�b�v�O���[�h����萔������ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"' ���N���b�N���Ă��������B
        'UPGRADE_ISSUE: Screen �v���p�e�B Screen.MousePointer �̓J�X�^�� �}�E�X�|�C���^���T�|�[�g���܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"' ���N���b�N���Ă��������B
        'UPGRADE_WARNING: Screen �v���p�e�B Screen.MousePointer �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
        'System.Windows.Forms.Cursor.Current = vbNormal

        '///// TD���ް�X�V�ɔ������C  2013.03.15 /////////////////////////

        'Set dbPgsqlFW = CreateObject("ADODB.Connection")
        dbOra11 = CreateObject("ADODB.Connection")
		
		'/////////////////////////////////////////////////////////////////
		
		strConnect = "dsn=" & gsDbName & ";"
		strConnect = strConnect & "uid=" & gsUsername & ";"
		strConnect = strConnect & "pwd=" & gsPassword
		
		'///// TD���ް�X�V�ɔ������C  2013.03.15 /////////////////////////
		
		'With dbPgsqlFW
		With dbOra11
			
			'/////////////////////////////////////////////////////////////////
			
			.Open((strConnect))
			.CursorLocation = 3
			
		End With
		
		Connected = True
		
		Exit Function
		
ErrorHandler: 
		
		If intAttempts < NUM_ATTEMPS Then
			intAttempts = intAttempts + 1
			Resume 
		Else
            'UPGRADE_ISSUE: vbNormal ���A�b�v�O���[�h����萔������ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"' ���N���b�N���Ă��������B
            'UPGRADE_ISSUE: Screen �v���p�e�B Screen.MousePointer �̓J�X�^�� �}�E�X�|�C���^���T�|�[�g���܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"' ���N���b�N���Ă��������B
            'UPGRADE_WARNING: Screen �v���p�e�B Screen.MousePointer �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
            'System.Windows.Forms.Cursor.Current = vbNormal
            'WriteMsg "ConnectADO: " & Err.Number & " " & Err.Description
            Connected = False
			ConnectDb = ERROR_OCCURRED
		End If
		
	End Function
	
	Public Function CloseDb() As Short
		
		On Error GoTo ErrorHandler
		
		CloseDb = NO_ERROR
		
		'///// TD���ް�X�V�ɔ������C  2013.03.15 /////////////////////////
		
		'dbPgsqlFW.Close
		dbOra11.Close()
		
		'/////////////////////////////////////////////////////////////////
		
		Exit Function
		
ErrorHandler: 
		
		CloseDb = ERROR_OCCURRED
		Resume Next
		
	End Function
	
	Public Function GetNextPKey(ByVal sTable As String) As Integer
		
		Dim lLstRecord As Integer
		Dim rDataset As Object
		Dim sSqlString As String
		
		sSqlString = "select max(SYSID) from " & sTable
		
		Debug.Print(sSqlString)
		
		'///// TD���ް�X�V�ɔ������C  2013.03.15 /////////////////////////
		
		'Set rDataset = dbPgsqlFW.Execute(sSqlString, 0&)
		rDataset = dbOra11.Execute(sSqlString, 0)
		
		'/////////////////////////////////////////////////////////////////
		
		'UPGRADE_WARNING: �I�u�W�F�N�g rDataset.Fields �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		'UPGRADE_WARNING: Null/IsNull() �̎g�p��������܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' ���N���b�N���Ă��������B
		If IsDbNull(rDataset.Fields(0)) Then
			lLstRecord = 1
		Else
			'UPGRADE_WARNING: �I�u�W�F�N�g rDataset.Fields �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
			lLstRecord = rDataset.Fields(0)
			lLstRecord = lLstRecord + 1
		End If
		
		GetNextPKey = lLstRecord
		
		'UPGRADE_NOTE: �I�u�W�F�N�g rDataset ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		rDataset = Nothing
		
	End Function
	
	'UPGRADE_NOTE: Class_Initialize �� Class_Initialize_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Private Sub Class_Initialize_Renamed()
		
		'WriteMsg "clsDBDirect Initialized"
		Call ConnectDb()
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate �� Class_Terminate_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Private Sub Class_Terminate_Renamed()
		
		'WriteMsg "clsDBDirect Terminated"
		Call CloseDb()
		
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class