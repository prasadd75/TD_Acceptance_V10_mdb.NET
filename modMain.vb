Option Strict Off
Option Explicit On

Imports System.IO
Imports GrapeCity.ActiveReports.Viewer.Win
'Imports GrapeCity.Viewer


Module modMain
	
	Public gsMdbPath As String
	Public gsPath As String
	Public gsRecLotId As String
	Public gsRecResourceId As String
	Public gsNetWeight As String
	
	Public gsSTC_ProductCode_Org As String
	Public gsSTC_ProductCode_Cores As String
	Public gsSTC_ProductCode_Chips As String
	Public gsSTC_ProductCode_Fragment As String
	Public gsCommPort As String
	Public gsCommPortSettings As String
	Public gsSTC_ProductCode As String
	Public gsWarehouseId As String
	Public gsFabId As String
	Public gsCustomerId As String
    Public gsRpt_MdbPath As String

    Public gsRpt_TdMaterialFilePath As String

    '-----------------------------------------------------------------
    'Add the TDF-O code by Hirokazu Furukawa on 2009/0109
    Public gsSTC_ProductCode_TDFO As String
	
	'-----------------------------------------------------------------
	'Add the Hemlock code by Hirokazu Furukawa on 2017/10/20
	Public gsSTC_ProductCode_Hemlock As String
	Public gsBatchSeq As String
	Public gsRecordTime As String
	
	Public gsDbName As String
	Public gsUsername As String
	Public gsPassword As String
	
	'UPGRADE_WARNING: Sub Main() が完了したときにアプリケーションは終了します。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"' をクリックしてください。
	Public Sub Main()

        'Dim oIniFile As clsIniFile

        '-----------------------------------------------------------------
        'Copy the MDB File parameter from Ini File
        'oIniFile = New clsIniFile

        'oIniFile.FullPath = My.Application.Info.DirectoryPath & "\iniFile\TD.ini"

        'Debug.Print(oIniFile.FullPath)

        'oIniFile.SectionName = "TD"

        'gsMdbPath = oIniFile.GetString("MDB_File_PATH")
        gsMdbPath = My.Settings.MDB_File_PATH

        'oIniFile.SectionName = "TD_STC_ProductCode"
        'gsSTC_ProductCode_Org = oIniFile.GetString("Org")
        'gsSTC_ProductCode_Cores = oIniFile.GetString("Cores")
        'gsSTC_ProductCode_Chips = oIniFile.GetString("Chips")
        'gsSTC_ProductCode_Fragment = oIniFile.GetString("Fragment")

        gsSTC_ProductCode_Org = My.Settings.Org
        gsSTC_ProductCode_Cores = My.Settings.Cores
        gsSTC_ProductCode_Chips = My.Settings.Chips
        gsSTC_ProductCode_Fragment = My.Settings.Fragments


        '-----------------------------------------------------------------
        'Add the TDF-O code by Hirokazu Furukawa on 2009/0109
        'gsSTC_ProductCode_TDFO = My.Settings.oIniFile.GetString("TDF-O_Org")

        '-----------------------------------------------------------------
        'Add the Hemlock code by Hirokazu Furukawa on 2017/10/20
        'gsSTC_ProductCode_Hemlock = oIniFile.GetString("Hemlock_Org")
        gsSTC_ProductCode_Hemlock = My.Settings.Hemlock_Org

        'oIniFile.SectionName = "SerialPort"
        'gsCommPort = oIniFile.GetString("Port")
        'gsCommPortSettings = oIniFile.GetString("Setting")
        gsCommPort = My.Settings.Serial_PortId
        gsCommPortSettings = My.Settings.Serial_Port_Setting

        'oIniFile.SectionName = "MIS_Data"
        'gsWarehouseId = oIniFile.GetString("WarehouseId")
        'gsFabId = oIniFile.GetString("FabId")
        'gsCustomerId = oIniFile.GetString("CustomerId")
        gsWarehouseId = My.Settings.MIS_WarehouseId
        gsFabId = My.Settings.MIS_FabId
        gsCustomerId = My.Settings.MIS_CustomerId

        'oIniFile.SectionName = "Report_Mdb"
        'gsRpt_MdbPath = oIniFile.GetString("MDB_File_PATH")
        gsRpt_MdbPath = My.Settings.RPT_MDB_File_Path

        'oIniFile.SectionName = "Database"
        'gsDbName = oIniFile.GetString("DbName")
        'gsUsername = oIniFile.GetString("Username")
        'gsPassword = oIniFile.GetString("Password")
        gsDbName = My.Settings.DB_DbName
        gsUsername = My.Settings.DB_Username
        gsPassword = My.Settings.DB_Password

        'UPGRADE_NOTE: オブジェクト oIniFile をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
        'oIniFile = Nothing

        gsPath = My.Application.Info.DirectoryPath

        gsRpt_TdMaterialFilePath = My.Settings.TdMaterialReportFilePath


        '-----------------------------------------------------------------
        'gsMdbPath = App.Path
        frmRegistration.Start()
		
	End Sub
	
	Public Function DeleteAllReportData(ByVal sTableName As String) As Object
		
		Dim sSql As String
		Dim dbDirect As clsDBDirectForRpt
		Dim rsDirect As ADODB.Recordset
		
		dbDirect = New clsDBDirectForRpt
		rsDirect = New ADODB.Recordset
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'sSql = "DELETE FROM " & sTableName & ";"
		sSql = "DELETE FROM " & sTableName
		
		'sSql = "DELETE FROM TD_Material_Warehousing_rpt_tbl;"
		
		'/////////////////////////////////////////////////////////////////
		
		Debug.Print(sSql)
		
		rsDirect.Open(sSql, dbDirect.dbFW)
		
		'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		rsDirect = Nothing
		'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbDirect = Nothing
		
	End Function
	
	Public Function RetrieveTdMaterialReportData(ByVal oTDMaterialRpt As clsTDMaterialRpt) As Object
		
		Dim sSql As String
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Dim dbDirect As clsPostgresSqlDBDirect
		'Dim dbDirect As clsOra11DBDirect
		
		Dim dbDirect As clsDBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		Dim rsDirect As ADODB.Recordset
		
		'-----------------------------------------------------------------
		'Initialize clsOra11DBDirect
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Set dbDirect = New clsPostgresSqlDBDirect
		'Set dbDirect = New clsOra11DBDirect
		
		dbDirect = New clsDBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		rsDirect = New ADODB.Recordset
		
		'-----------------------------------------------------------------
		'Build SQL Command
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		sSql = "Select * FROM td_material_tbl WHERE REC_Lot = '" & oTDMaterialRpt.Lot_IngotId & "';"
		'sSQL = "SELECT * FROM material_tbl WHERE REC_Lot = '" & oTDMaterialRpt.Lot_IngotId & "'"
		
		
		
		'/////////////////////////////////////////////////////////////////
		
		Debug.Print(sSql)
		
		'-----------------------------------------------------------------
		'Send SQL Command to ADO
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Set rsDirect = dbDirect.dbPgsqlFW.Execute(sSQL)
		'Set rsDirect = dbDirect.dbOra11.Execute(sSQL)
		
		rsDirect.Open(sSql, dbDirect.dbFW)
		
		'/////////////////////////////////////////////////////////////////
		
		If rsDirect.EOF = True Then
			MsgBox("入力されたロット番号は、存在しません。")
			
			'UPGRADE_NOTE: オブジェクト oTDMaterialRpt をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			oTDMaterialRpt = Nothing
			
			'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			rsDirect = Nothing
			'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			dbDirect = Nothing
			
			Exit Function
		Else
			oTDMaterialRpt.NetWeight = CStr(Val(rsDirect.Fields("NetWeight").Value) * 1000)
		End If
		
	End Function
	
	Public Function StoreTdMaterialReportData(ByVal oTDMaterialRpt As clsTDMaterialRpt) As Object
		
		Dim sSql As String
		Dim dbDirect As clsDBDirectForRpt
		Dim rsDirect As ADODB.Recordset
		Dim i As Short
		Dim lPkey As Integer
		
		'-----------------------------------------------------------------
		'Initialize clsDBDirectForRpt
		
		dbDirect = New clsDBDirectForRpt
		
		'-----------------------------------------------------------------
		'Get next record #
		
		lPkey = dbDirect.GetNextPKey("TD_Material_Warehousing_rpt_tbl")
		
		rsDirect = New ADODB.Recordset
		
		'-----------------------------------------------------------------
		'Buld SQL Command
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'sSQL = "INSERT INTO TD_Material_Warehousing_rpt_tbl" & _
		''        " VALUES (" & _
		''                    lPkey & "," & _
		''                    "'" & oTDMaterialRpt.WarehouseId & "'," & _
		''                    "'" & oTDMaterialRpt.FabId & "'," & _
		''                    "'" & oTDMaterialRpt.CustomerId & "'," & _
		''                    "'" & oTDMaterialRpt.STC_ProductCode & "'," & _
		''                    "'" & oTDMaterialRpt.Lot_IngotId & "'," & _
		''                    "'" & oTDMaterialRpt.Case_IngotId & "'," & _
		''                    "'" & oTDMaterialRpt.NetWeight & "'," & _
		''                    "'" & oTDMaterialRpt.Number & "'," & _
		''                    "'" & oTDMaterialRpt.PrintDate & "'," & _
		''                    "'" & oTDMaterialRpt.DocumentId & "'" & _
		''            ");"
		
		sSql = "INSERT INTO TD_Material_Warehousing_rpt_tbl"
		sSql = sSql & " VALUES ("
		sSql = sSql & lPkey & ","
		sSql = sSql & "'" & oTDMaterialRpt.WarehouseId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.FabId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.CustomerId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.STC_ProductCode & "',"
		sSql = sSql & "'" & oTDMaterialRpt.Lot_IngotId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.Case_IngotId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.NetWeight & "',"
		sSql = sSql & "'" & oTDMaterialRpt.Number & "',"
		sSql = sSql & "'" & oTDMaterialRpt.PrintDate & "',"
		sSql = sSql & "'" & oTDMaterialRpt.DocumentId & "'"
		sSql = sSql & ")"
		
		'/////////////////////////////////////////////////////////////////
		
		Debug.Print(sSql)
		
		rsDirect.Open(sSql, dbDirect.dbFW)
		
		'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		rsDirect = Nothing
		'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbDirect = Nothing
		
	End Function
	
	Public Function StoreTdMaterialReportDataToHdb(ByVal oTDMaterialRpt As clsTDMaterialRpt) As Object
		
		Dim sSql As String
		Dim dbDirect As clsDBDirectForRpt
		Dim rsDirect As ADODB.Recordset
		Dim i As Short
		Dim lPkey As Integer
		
		'-----------------------------------------------------------------
		'Initialize clsDBDirectForRpt
		
		dbDirect = New clsDBDirectForRpt
		
		'-----------------------------------------------------------------
		'Get next record #
		
		lPkey = dbDirect.GetNextPKey("TD_Material_Warehousing_rpt_hdb_tbl")
		
		rsDirect = New ADODB.Recordset
		
		'-----------------------------------------------------------------
		'Buld SQL Command
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'sSql = "INSERT INTO TD_Material_Warehousing_rpt_hdb_tbl" & _
		''        " VALUES (" & _
		''                    lPkey & "," & _
		''                    "'" & oTDMaterialRpt.WarehouseId & "'," & _
		''                    "'" & oTDMaterialRpt.FabId & "'," & _
		''                    "'" & oTDMaterialRpt.CustomerId & "'," & _
		''                    "'" & oTDMaterialRpt.STC_ProductCode & "'," & _
		''                    "'" & oTDMaterialRpt.Lot_IngotId & "'," & _
		''                    "'" & oTDMaterialRpt.Case_IngotId & "'," & _
		''                    "'" & oTDMaterialRpt.NetWeight & "'," & _
		''                    "'" & oTDMaterialRpt.Number & "'," & _
		''                    "'" & oTDMaterialRpt.PrintDate & "'," & _
		''                    "'" & oTDMaterialRpt.DocumentId & "'" & _
		''            ");"
		
		sSql = "INSERT INTO TD_Material_Warehousing_rpt_hdb_tbl"
		sSql = sSql & " VALUES ("
		sSql = sSql & lPkey & ","
		sSql = sSql & "'" & oTDMaterialRpt.WarehouseId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.FabId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.CustomerId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.STC_ProductCode & "',"
		sSql = sSql & "'" & oTDMaterialRpt.Lot_IngotId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.Case_IngotId & "',"
		sSql = sSql & "'" & oTDMaterialRpt.NetWeight & "',"
		sSql = sSql & "'" & oTDMaterialRpt.Number & "',"
		sSql = sSql & "'" & oTDMaterialRpt.PrintDate & "',"
		sSql = sSql & "'" & oTDMaterialRpt.DocumentId & "'"
		sSql = sSql & ")"
		
		'/////////////////////////////////////////////////////////////////
		
		Debug.Print(sSql)
		
		rsDirect.Open(sSql, dbDirect.dbFW)
		
		'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		rsDirect = Nothing
		'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbDirect = Nothing
		
	End Function
	
	Public Function GetTdProdyctType(ByVal sRecResourceType As String) As String
		
		Dim sSql As String
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Dim dbDirect As clsPostgresSqlDBDirect
		'Dim dbDirect As clsOra11DBDirect
		
		Dim dbDirect As clsDBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		Dim rsDirect As ADODB.Recordset
		Dim sTdProductType As String
		
		'-----------------------------------------------------------------
		'Initialize clsDBDirect
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Set dbDirect = New clsPostgresSqlDBDirect
		'Set dbDirect = New clsOra11DBDirect
		
		dbDirect = New clsDBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		rsDirect = New ADODB.Recordset
		
		'-----------------------------------------------------------------
		'Buld SQL Command
		
		sSql = "SELECT tdproducttype FROM td_producttype_master_tbl WHERE rec_resource = '" & sRecResourceType & "';"
		
		Debug.Print(sSql)
		
		'-----------------------------------------------------------------
		'Send SQL Command
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Set rsDirect = dbDirect.dbPgsqlFW.Execute(sSQL)
		'Set rsDirect = dbDirect.dbOra11.Execute(sSQL)
		
		rsDirect.Open(sSql, dbDirect.dbFW)
		
		'/////////////////////////////////////////////////////////////////
		
		sTdProductType = rsDirect.Fields("tdproducttype").Value
		
		'-----------------------------------------------------------------
		'Release clsDBDirect
		
		'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		rsDirect = Nothing
		'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbDirect = Nothing
		
		GetTdProdyctType = sTdProductType
		
	End Function
	
	Public Function GenerateBatchSeq(ByRef sRecLotId As String) As String
		
		Dim sBatchSeqId As String
		Dim sSql As String
		
		'Dim dbDirect As clsOra11DBDirect
		
		Dim dbDirect As clsDBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		Dim rsDirect As ADODB.Recordset
		Dim lPkey As Integer
		Dim sLastBatchSeqId As String
		Dim sLastBatchSeqId_Seq As String
		Dim sRecordtime As String
		Dim sProduct As String
		
		'-----------------------------------------------------------------
		'Initialize clsOra11DBDirect
		
		'Set dbDirect = New clsOra11DBDirect
		
		dbDirect = New clsDBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		rsDirect = New ADODB.Recordset
		
		sSql = "SELECT * FROM TD_BATCH_SEQ_CTRL_TBL WHERE BatchId = '" & sRecLotId & "';"
		'sSql = "SELECT * FROM BATCH_SEQ_CTRL_TBL WHERE reclotid = '" & sRecLotId & "'"
		
		'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		
		Debug.Print(sSql)
		
		'Set rsDirect = dbDirect.dbOra11.Execute(sSql)
		
		rsDirect.Open(sSql, dbDirect.dbFW)
		
		'/////////////////////////////////////////////////////////////////
		
		If rsDirect.EOF = True Then
			
			'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			rsDirect = Nothing
			'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			dbDirect = Nothing
			
			'-------------------------------------------------------------
			'Initialize clsOra11DBDirect
			
			'Set dbDirect = New clsOra11DBDirect
			
			dbDirect = New clsDBDirect
			
			'/////////////////////////////////////////////////////////////
			
			rsDirect = New ADODB.Recordset
			
			'-------------------------------------------------------------
			'Get next record #
			
			'lPkey = dbDirect.GetNextPKey("BATCH_SEQ_CTRL_TBL")
			
			lPkey = dbDirect.GetNextPKey("TD_BATCH_SEQ_CTRL_TBL")
			
			'-------------------------------------------------------------
			'Buld SQL Command
			
			gsRecordTime = VB6.Format(Now, "YYYYMMDDhhmmss")
			sLastBatchSeqId_Seq = "01"
			
			'Build Batch Sequence Id
			sLastBatchSeqId = sRecLotId & "-" & sLastBatchSeqId_Seq
			
			sSql = "INSERT INTO TD_BATCH_SEQ_CTRL_TBL" & " VALUES (" & lPkey & "," & "'" & sRecLotId & "'," & "'" & sLastBatchSeqId & "'," & "'" & sLastBatchSeqId_Seq & "'," & "'" & gsRecordTime & "'" & ");"
			
			'Set rsDirect = dbDirect.dbPgsqlFW.Execute(sSql)
			
			'sSql = "INSERT INTO TD_BATCH_SEQ_CTRL_TBL"
			'sSql = sSql & " VALUES ("
			'sSql = sSql & lPkey & ","
			'sSql = sSql & "'" & sLastBatchSeqId & "',"
			'sSql = sSql & "'" & sLastBatchSeqId_Seq & "',"
			'sSql = sSql & "'" & gsRecordTime & "'"
			'sSql = sSql & ")"
			
			'Set rsDirect = dbDirect.dbOra11.Execute(sSql)
			
			rsDirect.Open(sSql, dbDirect.dbFW)
			
			'/////////////////////////////////////////////////////////////
			
			Debug.Print(sSql)
			
			'rsDirect.Open sSql, dbDirect.dbFW
			
			'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			rsDirect = Nothing
			'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			dbDirect = Nothing
			
		Else
			
			sLastBatchSeqId = rsDirect.Fields("LAST_Batch_SEQ").Value
			sLastBatchSeqId_Seq = rsDirect.Fields("LAST_SEQ").Value
			
			'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			rsDirect = Nothing
			'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			dbDirect = Nothing
			
			'---------------------------------------------------------------------------------
			'
			If Val(sLastBatchSeqId_Seq) < 9 Then
				sLastBatchSeqId_Seq = "0" & CStr(Val(sLastBatchSeqId_Seq) + 1)
			ElseIf Val(sLastBatchSeqId_Seq) < 99 And Val(sLastBatchSeqId_Seq) > 8 Then 
				sLastBatchSeqId_Seq = CStr(Val(sLastBatchSeqId_Seq) + 1)
				'        Else
				'            sLastBatchSeqId_Seq = CStr(Val(sLastBatchSeqId_Seq) + 1)
			End If
			
			'---------------------------------------------------------------------------------
			'Build Batch Sequence Id
			
			sLastBatchSeqId = sRecLotId & "-" & sLastBatchSeqId_Seq
			
			gsRecordTime = VB6.Format(Now, "YYYYMMDDhhmmss")
			
			'-------------------------------------------------------------
			'Initialize clsOra11DBDirect
			
			'Set dbDirect = New clsOra11DBDirect
			
			dbDirect = New clsDBDirect
			
			'/////////////////////////////////////////////////////////////
			
			rsDirect = New ADODB.Recordset
			
			sSql = "UPDATE TD_BATCH_SEQ_CTRL_TBL SET " & "LAST_Batch_SEQ = '" & sLastBatchSeqId & "'," & "LAST_SEQ = '" & sLastBatchSeqId_Seq & "'," & "RecordTime = '" & gsRecordTime & "'" & " WHERE BatchId = '" & sRecLotId & "';"
			
			'Set rsDirect = dbDirect.dbPgsqlFW.Execute(sSql)
			
			'sSql = "UPDATE BATCH_SEQ_CTRL_TBL SET "
			'sSql = sSql & "LAST_Batch_SEQ = '" & sLastBatchSeqId & "',"
			'sSql = sSql & "LAST_SEQ = '" & sLastBatchSeqId_Seq & "',"
			'sSql = sSql & "RecordTime = '" & gsRecordTime & "'"
			'sSql = sSql & " WHERE ProductType = '" & sProductType & "' and RecLotId = '" & sRecLotId & "'"
			
			'Set rsDirect = dbDirect.dbOra11.Execute(sSql)
			
			rsDirect.Open(sSql, dbDirect.dbFW)
			
			'///////////////////////////////////////////////////////////////////////////////////////////
			
			Debug.Print(sSql)
			
			'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			rsDirect = Nothing
			'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
			dbDirect = Nothing
			
		End If
		
		GenerateBatchSeq = sLastBatchSeqId
		
	End Function
End Module