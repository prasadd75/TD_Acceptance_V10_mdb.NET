Option Strict Off
Option Explicit On
Friend Class frmRegistration
	Inherits System.Windows.Forms.Form

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click

        With Me
            .txtRecLotId.Text = ""
            .txtRecResourceId.Text = ""
            .txtNetWeight.Text = ""

            Me.cmdOK.Enabled = False
        End With

    End Sub
    'Friend WithEvents reportPreview As GrapeCity.ActiveReports.Viewer.Win.Viewer

    Private Sub cmdEnd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEnd.Click
		End
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		Dim sSql As String
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		'    Dim dbDirect As clsDBDirect
		'Dim dbDirect As clsPostgresSqlDBDirect
		'Dim dbDirect As clsOra11DBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		'    Dim rsDirect As ADODB.Recordset
		
		'入力データの取得
		
		With Me
			
			gsRecLotId = StrConv(Trim(.txtRecLotId.Text), VbStrConv.UpperCase)
			gsRecResourceId = StrConv(Trim(.txtRecResourceId.Text), VbStrConv.UpperCase)
			gsNetWeight = StrConv(Trim(.txtNetWeight.Text), VbStrConv.UpperCase)
			
		End With
		
		'-------------------------------------------------------------------------------------
		'Validate entered data
		
		If gsRecLotId = "" Then
			Exit Sub
		ElseIf gsRecResourceId = "" Then 
			Exit Sub
		ElseIf gsNetWeight = "" Then 
			Exit Sub
		Else
			
			'-----------------------------------------------------------------------------
			'Initialize clsOra11DBDirect
			
			'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////////////////
			'            Set dbDirect = New clsDBDirect
			'Set dbDirect = New clsPostgresSqlDBDirect
			'Set dbDirect = New clsOra11DBDirect
			
			'/////////////////////////////////////////////////////////////////////////////
			
			'            Set rsDirect = New ADODB.Recordset
			
			'-----------------------------------------------------------------------------
			'Build SQL Command
			
			'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////////////////
			
			'sSQL = "SELECT * FROM material_tbl WHERE REC_Lot = '" & gsRecLotId & "';"
			'sSQL = "SELECT * FROM material_tbl WHERE REC_Lot = '" & gsRecLotId & "'"
			
			'            sSql = "SELECT * FROM td_material_tbl WHERE REC_Lot = '" & gsRecLotId & "';"
			
			'/////////////////////////////////////////////////////////////////////////////
			
			'            Debug.Print sSql
			
			'-----------------------------------------------------------------------------
			'Send SQL Command to ADO
			
			'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////////////////
			'            rsDirect.Open sSql, dbDirect.dbFW
			'Set rsDirect = dbDirect.dbPgsqlFW.Execute(sSQL)
			'Set rsDirect = dbDirect.dbOra11.Execute(sSQL)
			
			'/////////////////////////////////////////////////////////////////////////////
			
			'            If rsDirect.EOF = False Then
			'                MsgBox ("入力したTDロットは既に登録が終わっていますので、処理が行えません。")
			'
			'               '-------------------------------------------------------------------------
			'               'Init screen
			'
			'                With Me
			'                    .txtNetWeight.Text = ""
			'                    .txtRecLotId.Text = ""
			'                    .txtRecResourceId.Text = ""
			'                    .cmdOK.Enabled = False
			'                End With
			'
			'                Exit Sub
			'
			'            End If
			
			'----------------------------------------------------------------------------
			' Generate the BatchSequence Id
			'Apply for Hemlock
			gsBatchSeq = GenerateBatchSeq(gsRecLotId)
			
			'-----------------------------------------------------------------------------
			'Store entered Rec Lot Data to database
			
			Call RegisterRecLotData(gsRecLotId, gsRecResourceId, gsNetWeight, gsBatchSeq)
			
			'-----------------------------------------------------------------------------
			'Print TD Material Report
			
			Call PrintTdMaterialReport()
			
			'-----------------------------------------------------------------------------
			'Init screen
			
			With Me
				.txtNetWeight.Text = ""
				.txtRecLotId.Text = ""
				.txtRecResourceId.Text = ""
				.cmdOK.Enabled = False
			End With
			
		End If
		
	End Sub
	
	Public Sub Start()

        Me.ShowDialog()
        Me.txtRecLotId.Focus()
		Me.cmdOK.Enabled = False
		
	End Sub
	
	Private Sub txtRecLotId_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRecLotId.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then
			txtRecResourceId.Focus()
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtRecResourceId_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRecResourceId.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then
			txtNetWeight.Focus()
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtNetWeight_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNetWeight.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then
			
			If txtRecLotId.Text = "" Then
				GoTo EventExitSub
			ElseIf txtRecResourceId.Text = "" Then 
				GoTo EventExitSub
			ElseIf txtRecLotId.Text = "" Then 
				GoTo EventExitSub
			Else
				Me.cmdOK.Enabled = True
			End If
			
		End If
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Public Function RegisterRecLotData(ByVal sRecLotId As String, ByVal sRecResourceId As String, ByVal sNetWeight As String, ByVal sBatchSeq As String) As Object
		
		Dim sSql As String

        '///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////

        'Dim dbDirect As clsPostgresSqlDBDirect
        'Dim dbDirect As clsOra11DBDirect

        Dim dbDirect As clsDBDirect

        '/////////////////////////////////////////////////////////////////

        Dim rsDirect As ADODB.Recordset
        Dim sDate As String
		
		sDate = VB6.Format(Now, "YYYYMMDDhhmmss")
		
		'-----------------------------------------------------------------
		'Initialize ADO Session
		'-----------------------------------------------------------------
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Set dbDirect = New clsPostgresSqlDBDirect
		'Set dbDirect = New clsOra11DBDirect
		
		dbDirect = New clsDBDirect
		
		'/////////////////////////////////////////////////////////////////
		
		rsDirect = New ADODB.Recordset
		
		'-----------------------------------------------------------------
		' Build SQL Command
		'-----------------------------------------------------------------
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'sSql = "INSERT INTO material_tbl " & _
		''        "VALUES (" & _
		''        "'" & sRecLotId & "'," & _
		''        "'" & sRecResourceId & "'," & _
		''        sNetWeight & "," & _
		''        "'" & sDate & "'," & _
		''        "'" & " " & "'," & _
		''        "'" & "N" & "'" & _
		''        ");"
		
		sSql = "INSERT INTO td_material_tbl ("
		sSql = sSql & "rec_lot, "
		sSql = sSql & "rec_resource, "
		sSql = sSql & "netweight, "
		sSql = sSql & "registrationtime, "
		sSql = sSql & "consumeddate, "
		sSql = sSql & "consumedflg, "
		sSql = sSql & "Batch_SEQ) "
		sSql = sSql & " VALUES ("
		sSql = sSql & "'" & sRecLotId & "',"
		sSql = sSql & "'" & sRecResourceId & "',"
		sSql = sSql & sNetWeight & ","
		sSql = sSql & "'" & sDate & "',"
		sSql = sSql & "'" & " " & "',"
		sSql = sSql & "'" & "N" & "',"
		sSql = sSql & "'" & sBatchSeq & "');"
		
		'/////////////////////////////////////////////////////////////////
		
		Debug.Print(sSql)
		
		'///// TDｻｰﾊﾞｰ更新に伴う改修  2013.03.15 /////////////////////////
		
		'Set rsDirect = dbDirect.dbPgsqlFW.Execute(sSQL)
		'Set rsDirect = dbDirect.dbOra11.Execute(sSQL)
		
		rsDirect.Open(sSql, dbDirect.dbFW)
		
		'/////////////////////////////////////////////////////////////////
		
		'-----------------------------------------------------------------
		'Release class object
		'-----------------------------------------------------------------
		'UPGRADE_NOTE: オブジェクト rsDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		rsDirect = Nothing
		'UPGRADE_NOTE: オブジェクト dbDirect をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		dbDirect = Nothing



    End Function
	
	Public Function PrintTdMaterialReport() As Object
        'Dim rptTD_Material As Object

        Dim sRecLotId As String
		Dim oTDMaterialRpt As clsTDMaterialRpt
		Dim oTDMaterialRpts As clsTDMaterialRpts
		
		'-----------------------------------------------------------------
		'Add the TDF-O by Hirokazu Furukawa on 2009/0109
		Dim sTdProductType As String
		Dim sTdResourceId As String
		'-----------------------------------------------------------------
		'Retrieve Rec Lot Id
		
		sRecLotId = Me.txtRecLotId.Text
		
		'-----------------------------------------------------------------
		'Add the TDF-O by Hirokazu Furukawa on 2009/0109
		sTdResourceId = Me.txtRecResourceId.Text
		sTdProductType = GetTdProdyctType(sTdResourceId)
		'-----------------------------------------------------------------
		'Instance class oTDMaterialRpt
		
		oTDMaterialRpt = New clsTDMaterialRpt
		
		'-----------------------------------------------------------------
		'
		With oTDMaterialRpt
			
			'-------------------------------------------------------------
			'Add the TDF-O by Hirokazu Furukawa on 2009/0109
			If sTdProductType = "TD" Then
				.STC_ProductCode = gsSTC_ProductCode_Org
			ElseIf sTdProductType = "TDF-O" Then 
				.STC_ProductCode = gsSTC_ProductCode_TDFO
			Else
				'-------------------------------------------------------------
				'Add the Hemlock by Hirokazu Furukawa on 2017/10/20
				.STC_ProductCode = gsSTC_ProductCode_Hemlock
			End If
			
			.CustomerId = gsCustomerId
			.FabId = gsFabId
			.WarehouseId = gsWarehouseId
			.Number = "1"
			.Case_IngotId = sRecLotId
			.Lot_IngotId = sRecLotId
			.DocumentId = VB6.Format(Now, "YYYYMMDDhhmmss")
			.PrintDate = VB6.Format(Now, "YYYY/MM/DD hh:mm")
			
		End With
		
		'-----------------------------------------------------------------
		'Delete all the report data in reporting database
		Call DeleteAllReportData("TD_Material_Warehousing_rpt_tbl")
		
		'-----------------------------------------------------------------
		'Retrieve Report data from database
		Call RetrieveTdMaterialReportData(oTDMaterialRpt)
		
		Call StoreTdMaterialReportData(oTDMaterialRpt)
		
		Call StoreTdMaterialReportDataToHdb(oTDMaterialRpt)

        '-----------------------------------------------------------------
        'Display Report
        'rptTD_Material.show

        Dim RptViewForm As New ReportViewForm

        RptViewForm.ShowDialog(Me)
        'RptViewForm.Show(Me)


        RptViewForm.Dispose()


        'reportPreview.LoadDocument(pageReport.Document)

        '-----------------------------------------------------------------
        'Release class oTDMaterialRpt

        'UPGRADE_NOTE: オブジェクト oTDMaterialRpt をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
        'oTDMaterialRpt = Nothing


    End Function
End Class