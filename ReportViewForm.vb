Imports System.IO
Imports GrapeCity.ActiveReports

Public Class ReportViewForm

    Private Sub ReportViewForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim rptPath As New FileInfo(gsRpt_TdMaterialFilePath)
        Dim pageReport As New GrapeCity.ActiveReports.PageReport(rptPath)
        Dim pageDocument As New GrapeCity.ActiveReports.Document.PageDocument(pageReport)

        System.Threading.Thread.Sleep(1000)





        reportPreview.LoadDocument(pageDocument)

        'reportPreview.LoadDocument(PageReport.Document)

    End Sub

End Class