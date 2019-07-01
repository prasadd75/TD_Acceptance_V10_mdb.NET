<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportViewForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.reportPreview = New GrapeCity.ActiveReports.Viewer.Win.Viewer()
        Me.SuspendLayout()
        '
        'reportPreview
        '
        Me.reportPreview.CurrentPage = 0
        Me.reportPreview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.reportPreview.Location = New System.Drawing.Point(0, 0)
        Me.reportPreview.Name = "reportPreview"
        Me.reportPreview.PreviewPages = 0
        '
        '
        '
        '
        '
        '
        Me.reportPreview.Sidebar.ParametersPanel.ContextMenu = Nothing
        Me.reportPreview.Sidebar.ParametersPanel.Text = "パラメータ"
        Me.reportPreview.Sidebar.ParametersPanel.Width = 200
        '
        '
        '
        Me.reportPreview.Sidebar.SearchPanel.ContextMenu = Nothing
        Me.reportPreview.Sidebar.SearchPanel.Text = "検索"
        Me.reportPreview.Sidebar.SearchPanel.Width = 200
        '
        '
        '
        Me.reportPreview.Sidebar.ThumbnailsPanel.ContextMenu = Nothing
        Me.reportPreview.Sidebar.ThumbnailsPanel.Text = "サムネイル"
        Me.reportPreview.Sidebar.ThumbnailsPanel.Width = 200
        Me.reportPreview.Sidebar.ThumbnailsPanel.Zoom = 0.1R
        '
        '
        '
        Me.reportPreview.Sidebar.TocPanel.ContextMenu = Nothing
        Me.reportPreview.Sidebar.TocPanel.Expanded = True
        Me.reportPreview.Sidebar.TocPanel.Text = "見出しマップラベル"
        Me.reportPreview.Sidebar.TocPanel.Width = 200
        Me.reportPreview.Sidebar.Width = 200
        Me.reportPreview.Size = New System.Drawing.Size(824, 440)
        Me.reportPreview.TabIndex = 0
        '
        'ReportViewForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(824, 440)
        Me.Controls.Add(Me.reportPreview)
        Me.Name = "ReportViewForm"
        Me.Text = "ReportViewForm"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents reportPreview As GrapeCity.ActiveReports.Viewer.Win.Viewer
End Class
