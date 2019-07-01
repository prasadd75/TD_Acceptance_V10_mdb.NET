Option Strict Off
Option Explicit On
Friend Class clsDBDirect
	' Purpose:      This object is used to read/write the ADE site data directly from/to the
	'               database using RDO.
	'               NOTE: If the database schema changes this code will not longer work
	'                     The following tables are accessed
	'                       AdeSiteData - This is a custom table that was added to the EDC table space.
	' Inputs:       None
	' Dependencies: GlobalFunctions.bas
	'               INIFunctions.bas
	' Returns:
	'
	' Revision History:
	' AUTHOR            DATE       DESCRIPTION
	'-------------------------------------------------------------------------------------
	' akt               6/18/98     Original - Patterned after clsTabInfo.
	'
	'
	
	Public Connected As Boolean
	Private Const msFILE_NAME As String = "clsDBDirect"
	
	
	'This is ADO.
	Public dbFW As ADODB.Connection
	Public rsFW As ADODB.Recordset
	
	'Private l_RS As RDO.rdoResultset
	
	Private Const NUM_ATTEMPS As Short = 3 'ATTEMPT TO CONNECT TO THE DATABASE THIS NO. OF TIMES
	Private Const NO_ERROR As Short = 0 'NO ERROR OCCURRED
	Private Const ERROR_OCCURRED As Short = 1 'AN ERROR OCCURRED
	
	
	Public Function ConnectADO() As Short
		On Error GoTo ErrorHandler
		Dim strConnect As String
		Dim intAttempts As Short
		
		
		ConnectADO = NO_ERROR
		intAttempts = 1
		
		'    Screen.MousePointer = vbHourglass
		'    Screen.MousePointer = vbNormal
		
		strConnect = "Data Source=" & gsMdbPath
		
		dbFW = New ADODB.Connection
		rsFW = New ADODB.Recordset
		
		With dbFW
			.Provider = "Microsoft.Jet.OLEDB.4.0"
			.ConnectionString = strConnect
			.Open()
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
            '        WriteMsg "ConnectADO: " & Err.Number & " " & Err.Description
            Connected = False
			ConnectADO = ERROR_OCCURRED
		End If
	End Function
	
	
	'Private Function l_CloseAllResultSets() As Integer
	'    On Error GoTo ErrorHandler
	'
	'    l_CloseAllResultSets = NO_ERROR
	'    For Each l_RS In dbFW.rdoResultsets
	'        l_RS.Close
	'        Set l_RS = Nothing
	'    Next
	'
	'ErrorHandler:
	'    l_CloseAllResultSets = ERROR_OCCURRED
	'    Resume Next
	'
	'End Function
	
	Public Function CloseDB() As Short
		'Dim Inst_Connect As ADODB.Connection
		'Dim rtnval As Integer
		
		On Error GoTo ErrorHandler
		
		'rtnval = l_CloseAllResultSets()
		
		'CloseDB = NO_ERROR
		'For Each Inst_Connect In envFW.rdoConnections
		'Inst_Connect.Close
		dbFW.Close()
		'Next
		
		Exit Function
		
ErrorHandler: 
		CloseDB = ERROR_OCCURRED
		Resume Next
		
	End Function
	Public Function GetNextPKey(ByVal sTable As String) As Integer
		Dim lLstRecord As Integer
		Dim rDataset As New ADODB.Recordset
		Dim sSqlString As String
		
		sSqlString = "select max(SysId) from " & sTable & ";"
		
		rDataset.Open(sSqlString, dbFW)
		
		'UPGRADE_WARNING: Null/IsNull() �̎g�p��������܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"' ���N���b�N���Ă��������B
		If IsDbNull(rDataset.Fields(0).Value) Then
			lLstRecord = 1
		Else
			lLstRecord = rDataset.Fields(0).Value
			lLstRecord = lLstRecord + 1
		End If
		
		
		GetNextPKey = lLstRecord

        'UPGRADE_NOTE: �I�u�W�F�N�g rDataset ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
        rDataset = Nothing

    End Function
	
	'UPGRADE_NOTE: Class_Initialize �� Class_Initialize_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Private Sub Class_Initialize_Renamed()
		'    WriteMsg "clsDBDirect Initialized"
		Call ConnectADO()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate �� Class_Terminate_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Private Sub Class_Terminate_Renamed()
		'    WriteMsg "clsDBDirect Terminated"
		Call CloseDB()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class