Option Strict Off
Option Explicit On
Friend Class clsTDMaterialRpts
	Implements System.Collections.IEnumerable
	'ｺﾚｸｼｮﾝを保持するﾛｰｶﾙ変数
	Private mCol As Collection
	Public Sub AddObject(ByVal oTDMaterialRpt As clsTDMaterialRpt)
		mCol.Add(oTDMaterialRpt)
	End Sub
	
	
	Public Function Add(ByRef WarehouseId As String, ByRef FabId As String, ByRef CustomerId As String, ByRef STC_ProductCode As String, ByRef Lot_IngotId As String, ByRef Case_IngotId As String, ByRef NetWeight As String, ByRef Number As String, ByRef PrintDate As String, ByRef DocumentId As String, Optional ByRef sKey As String = "") As clsTDMaterialRpt
		'新規ｵﾌﾞｼﾞｪｸﾄを作成します。
		Dim objNewMember As clsTDMaterialRpt
		objNewMember = New clsTDMaterialRpt
		
		
		'ﾒｿｯﾄﾞに渡すﾌﾟﾛﾊﾟﾃｨを設定します。
		objNewMember.WarehouseId = WarehouseId
		objNewMember.FabId = FabId
		objNewMember.CustomerId = CustomerId
		objNewMember.STC_ProductCode = STC_ProductCode
		objNewMember.Lot_IngotId = Lot_IngotId
		objNewMember.Case_IngotId = Case_IngotId
		objNewMember.NetWeight = NetWeight
		objNewMember.Number = Number
		objNewMember.PrintDate = PrintDate
		objNewMember.DocumentId = DocumentId
		If Len(sKey) = 0 Then
			mCol.Add(objNewMember)
		Else
			mCol.Add(objNewMember, sKey)
		End If
		
		
		'作成されたｵﾌﾞｼﾞｪｸﾄを返します。
		Add = objNewMember
		'UPGRADE_NOTE: オブジェクト objNewMember をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		objNewMember = Nothing
		
		
	End Function
	
	Default Public ReadOnly Property Item(ByVal vntIndexKey As Object) As clsTDMaterialRpt
		Get
			'ｺﾚｸｼｮﾝの要素を参照するときに使用します。
			'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
			'保持するために Variant で宣言されています。
			'構文: Set foo = x.Item(xyz) または Set foo = x.Item(5)
			Item = mCol.Item(vntIndexKey)
		End Get
	End Property
	
	
	
	Public ReadOnly Property Count() As Integer
		Get
			'ｺﾚｸｼｮﾝの要素数を取得するときに使用します。
			'構文: Debug.Print x.Count
			Count = mCol.Count()
		End Get
	End Property
	
	
	'UPGRADE_NOTE: NewEnum プロパティがコメント アウトされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"' をクリックしてください。
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'このﾌﾟﾛﾊﾟﾃｨは、For...Each 構文を使用して
			'ｺﾚｸｼｮﾝを列挙できるようにします。
			'NewEnum = mCol._NewEnum
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: コレクション列挙子を返すには、コメントを解除して以下の行を変更してください。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"' をクリックしてください。
		'GetEnumerator = mCol.GetEnumerator
	End Function
	
	
	Public Sub Remove(ByRef vntIndexKey As Object)
		'ｺﾚｸｼｮﾝから要素を削除するときに使用します。
		'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
		'保持するために Variant で宣言されています。
		'構文: x.Remove(xyz)
		
		
		mCol.Remove(vntIndexKey)
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize は Class_Initialize_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
	Private Sub Class_Initialize_Renamed()
		'このｸﾗｽが作成されたときに、ｺﾚｸｼｮﾝを作成します。
		mCol = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate は Class_Terminate_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
	Private Sub Class_Terminate_Renamed()
		'このｸﾗｽが終了するときに、ｺﾚｸｼｮﾝを破棄します。
		'UPGRADE_NOTE: オブジェクト mCol をガベージ コレクトするまでこのオブジェクトを破棄することはできません。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' をクリックしてください。
		mCol = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class