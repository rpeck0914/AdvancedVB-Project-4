Imports DataBaseTables.Tables

Public Class WebForm2
    Inherits System.Web.UI.Page

    'Dim TheConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=J:\CIS212\RobertPeckProject4\RobertPeckProject4\Northwind.mdb"
    Private aDataRun As New DataBaseTables.Tables.DataBaseTableSelection()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            GridView1.DataSource = (aDataRun.RetrieveCustomers("f")).Tables(0)
            GridView1.DataBind()

        Catch ex As Exception

        End Try
    End Sub

End Class