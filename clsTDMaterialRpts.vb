Option Strict Off
Option Explicit On
Friend Class clsTDMaterialRpts
	Implements System.Collections.IEnumerable
	'�ڸ��݂�ێ�����۰�ٕϐ�
	Private mCol As Collection
	Public Sub AddObject(ByVal oTDMaterialRpt As clsTDMaterialRpt)
		mCol.Add(oTDMaterialRpt)
	End Sub
	
	
	Public Function Add(ByRef WarehouseId As String, ByRef FabId As String, ByRef CustomerId As String, ByRef STC_ProductCode As String, ByRef Lot_IngotId As String, ByRef Case_IngotId As String, ByRef NetWeight As String, ByRef Number As String, ByRef PrintDate As String, ByRef DocumentId As String, Optional ByRef sKey As String = "") As clsTDMaterialRpt
		'�V�K��޼ު�Ă��쐬���܂��B
		Dim objNewMember As clsTDMaterialRpt
		objNewMember = New clsTDMaterialRpt
		
		
		'ҿ��ނɓn�������è��ݒ肵�܂��B
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
		
		
		'�쐬���ꂽ��޼ު�Ă�Ԃ��܂��B
		Add = objNewMember
		'UPGRADE_NOTE: �I�u�W�F�N�g objNewMember ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		objNewMember = Nothing
		
		
	End Function
	
	Default Public ReadOnly Property Item(ByVal vntIndexKey As Object) As clsTDMaterialRpt
		Get
			'�ڸ��݂̗v�f���Q�Ƃ���Ƃ��Ɏg�p���܂��B
			'vntIndexKey �� ���ޯ���܂��ͷ��̂ǂ��炩��
			'�ێ����邽�߂� Variant �Ő錾����Ă��܂��B
			'�\��: Set foo = x.Item(xyz) �܂��� Set foo = x.Item(5)
			Item = mCol.Item(vntIndexKey)
		End Get
	End Property
	
	
	
	Public ReadOnly Property Count() As Integer
		Get
			'�ڸ��݂̗v�f�����擾����Ƃ��Ɏg�p���܂��B
			'�\��: Debug.Print x.Count
			Count = mCol.Count()
		End Get
	End Property
	
	
	'UPGRADE_NOTE: NewEnum �v���p�e�B���R�����g �A�E�g����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"' ���N���b�N���Ă��������B
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'���������è�́AFor...Each �\�����g�p����
			'�ڸ��݂�񋓂ł���悤�ɂ��܂��B
			'NewEnum = mCol._NewEnum
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: �R���N�V�����񋓎q��Ԃ��ɂ́A�R�����g���������Ĉȉ��̍s��ύX���Ă��������B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"' ���N���b�N���Ă��������B
		'GetEnumerator = mCol.GetEnumerator
	End Function
	
	
	Public Sub Remove(ByRef vntIndexKey As Object)
		'�ڸ��݂���v�f���폜����Ƃ��Ɏg�p���܂��B
		'vntIndexKey �� ���ޯ���܂��ͷ��̂ǂ��炩��
		'�ێ����邽�߂� Variant �Ő錾����Ă��܂��B
		'�\��: x.Remove(xyz)
		
		
		mCol.Remove(vntIndexKey)
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize �� Class_Initialize_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Private Sub Class_Initialize_Renamed()
		'���̸׽���쐬���ꂽ�Ƃ��ɁA�ڸ��݂��쐬���܂��B
		mCol = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate �� Class_Terminate_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Private Sub Class_Terminate_Renamed()
		'���̸׽���I������Ƃ��ɁA�ڸ��݂�j�����܂��B
		'UPGRADE_NOTE: �I�u�W�F�N�g mCol ���K�x�[�W �R���N�g����܂ł��̃I�u�W�F�N�g��j�����邱�Ƃ͂ł��܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"' ���N���b�N���Ă��������B
		mCol = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class