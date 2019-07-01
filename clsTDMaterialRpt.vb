Option Strict Off
Option Explicit On
Friend Class clsTDMaterialRpt
	'ﾌﾟﾛﾊﾟﾃｨ値を保持するためのﾛｰｶﾙ変数。
	Private mvarWarehouseId As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarFabId As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarCustomerId As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarSTC_ProductCode As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarLot_IngotId As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarCase_IngotId As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarNetWeight As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarNumber As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarPrintDate As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	Private mvarDocumentId As String 'ﾛｰｶﾙ ｺﾋﾟｰ
	
	
	Public Property DocumentId() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.DocumentId
			DocumentId = mvarDocumentId
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.DocumentId = 5
			mvarDocumentId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property PrintDate() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.PrintDate
			PrintDate = mvarPrintDate
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.PrintDate = 5
			mvarPrintDate = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Number() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.Number
			Number = mvarNumber
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.Number = 5
			mvarNumber = Value
		End Set
	End Property
	
	
	
	
	
	Public Property NetWeight() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.NetWeight
			NetWeight = mvarNetWeight
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.NetWeight = 5
			mvarNetWeight = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Case_IngotId() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.Case_IngotId
			Case_IngotId = mvarCase_IngotId
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.Case_IngotId = 5
			mvarCase_IngotId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Lot_IngotId() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.Lot_IngotId
			Lot_IngotId = mvarLot_IngotId
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.Lot_IngotId = 5
			mvarLot_IngotId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property STC_ProductCode() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.STC_ProductCode
			STC_ProductCode = mvarSTC_ProductCode
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.STC_ProductCode = 5
			mvarSTC_ProductCode = Value
		End Set
	End Property
	
	
	
	
	
	Public Property CustomerId() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.CustomerId
			CustomerId = mvarCustomerId
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.CustomerId = 5
			mvarCustomerId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property FabId() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.FabId
			FabId = mvarFabId
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.FabId = 5
			mvarFabId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property WarehouseId() As String
		Get
			'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
			'Syntax: Debug.Print X.WarehouseId
			WarehouseId = mvarWarehouseId
		End Get
		Set(ByVal Value As String)
			'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
			'Syntax: X.WarehouseId = 5
			mvarWarehouseId = Value
		End Set
	End Property
End Class