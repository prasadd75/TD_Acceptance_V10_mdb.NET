Option Strict Off
Option Explicit On
Friend Class clsTDMaterialRpt
	'�����è�l��ێ����邽�߂�۰�ٕϐ��B
	Private mvarWarehouseId As String '۰�� ��߰
	Private mvarFabId As String '۰�� ��߰
	Private mvarCustomerId As String '۰�� ��߰
	Private mvarSTC_ProductCode As String '۰�� ��߰
	Private mvarLot_IngotId As String '۰�� ��߰
	Private mvarCase_IngotId As String '۰�� ��߰
	Private mvarNetWeight As String '۰�� ��߰
	Private mvarNumber As String '۰�� ��߰
	Private mvarPrintDate As String '۰�� ��߰
	Private mvarDocumentId As String '۰�� ��߰
	
	
	Public Property DocumentId() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.DocumentId
			DocumentId = mvarDocumentId
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.DocumentId = 5
			mvarDocumentId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property PrintDate() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.PrintDate
			PrintDate = mvarPrintDate
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.PrintDate = 5
			mvarPrintDate = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Number() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.Number
			Number = mvarNumber
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.Number = 5
			mvarNumber = Value
		End Set
	End Property
	
	
	
	
	
	Public Property NetWeight() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.NetWeight
			NetWeight = mvarNetWeight
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.NetWeight = 5
			mvarNetWeight = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Case_IngotId() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.Case_IngotId
			Case_IngotId = mvarCase_IngotId
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.Case_IngotId = 5
			mvarCase_IngotId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property Lot_IngotId() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.Lot_IngotId
			Lot_IngotId = mvarLot_IngotId
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.Lot_IngotId = 5
			mvarLot_IngotId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property STC_ProductCode() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.STC_ProductCode
			STC_ProductCode = mvarSTC_ProductCode
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.STC_ProductCode = 5
			mvarSTC_ProductCode = Value
		End Set
	End Property
	
	
	
	
	
	Public Property CustomerId() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.CustomerId
			CustomerId = mvarCustomerId
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.CustomerId = 5
			mvarCustomerId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property FabId() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.FabId
			FabId = mvarFabId
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.FabId = 5
			mvarFabId = Value
		End Set
	End Property
	
	
	
	
	
	Public Property WarehouseId() As String
		Get
			'�����è�̒l���擾����Ƃ��ɁA������̉E�ӂŎg�p���܂��B
			'Syntax: Debug.Print X.WarehouseId
			WarehouseId = mvarWarehouseId
		End Get
		Set(ByVal Value As String)
			'�����è�ɒl��������Ƃ��ɁA������̍��ӂŎg�p���܂��B
			'Syntax: X.WarehouseId = 5
			mvarWarehouseId = Value
		End Set
	End Property
End Class