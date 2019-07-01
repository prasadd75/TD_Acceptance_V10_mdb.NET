<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmRegistration
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdEnd As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtNetWeight As System.Windows.Forms.TextBox
	Public WithEvents txtRecResourceId As System.Windows.Forms.TextBox
	Public WithEvents txtRecLotId As System.Windows.Forms.TextBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdEnd = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtNetWeight = New System.Windows.Forms.TextBox()
        Me.txtRecResourceId = New System.Windows.Forms.TextBox()
        Me.txtRecLotId = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdEnd
        '
        Me.cmdEnd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnd.Location = New System.Drawing.Point(208, 168)
        Me.cmdEnd.Name = "cmdEnd"
        Me.cmdEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEnd.Size = New System.Drawing.Size(97, 25)
        Me.cmdEnd.TabIndex = 10
        Me.cmdEnd.Text = "終了"
        Me.cmdEnd.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(104, 168)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(97, 25)
        Me.cmdOK.TabIndex = 4
        Me.cmdOK.Text = "受入"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(312, 168)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(105, 25)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "キャンセル"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtNetWeight)
        Me.Frame1.Controls.Add(Me.txtRecResourceId)
        Me.Frame1.Controls.Add(Me.txtRecLotId)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(0, 32)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(417, 129)
        Me.Frame1.TabIndex = 6
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "TDデータ登録"
        '
        'txtNetWeight
        '
        Me.txtNetWeight.AcceptsReturn = True
        Me.txtNetWeight.BackColor = System.Drawing.SystemColors.Window
        Me.txtNetWeight.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNetWeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNetWeight.Location = New System.Drawing.Point(152, 88)
        Me.txtNetWeight.MaxLength = 0
        Me.txtNetWeight.Name = "txtNetWeight"
        Me.txtNetWeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNetWeight.Size = New System.Drawing.Size(257, 26)
        Me.txtNetWeight.TabIndex = 3
        Me.txtNetWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtRecResourceId
        '
        Me.txtRecResourceId.AcceptsReturn = True
        Me.txtRecResourceId.BackColor = System.Drawing.SystemColors.Window
        Me.txtRecResourceId.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRecResourceId.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRecResourceId.Location = New System.Drawing.Point(152, 56)
        Me.txtRecResourceId.MaxLength = 0
        Me.txtRecResourceId.Name = "txtRecResourceId"
        Me.txtRecResourceId.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRecResourceId.Size = New System.Drawing.Size(257, 25)
        Me.txtRecResourceId.TabIndex = 2
        Me.txtRecResourceId.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtRecLotId
        '
        Me.txtRecLotId.AcceptsReturn = True
        Me.txtRecLotId.BackColor = System.Drawing.SystemColors.Window
        Me.txtRecLotId.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRecLotId.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRecLotId.Location = New System.Drawing.Point(152, 24)
        Me.txtRecLotId.MaxLength = 0
        Me.txtRecLotId.Name = "txtRecLotId"
        Me.txtRecLotId.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRecLotId.Size = New System.Drawing.Size(257, 26)
        Me.txtRecLotId.TabIndex = 1
        Me.txtRecLotId.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label4.Location = New System.Drawing.Point(8, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(145, 25)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "重量(Kg)"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label3.Location = New System.Drawing.Point(8, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(145, 25)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "リソース番号(Resource)"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label2.Location = New System.Drawing.Point(8, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(145, 25)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "ロット番号"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(425, 25)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "TD原料受け入れ処理"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmRegistration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(423, 200)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdEnd)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRegistration"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmRegistration"
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class