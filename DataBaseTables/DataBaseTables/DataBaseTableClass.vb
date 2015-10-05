' Program: Business Rule & Data Classes
'  Author: Robert Peck
'    Data: 03/09/2015

Option Strict On
Option Explicit On
Imports DataBaseConnectorLibrary.LibraryClass

Namespace Tables
    Public Class DataBaseTableSelection
        'Private _TheConnectionString As String = String.Empty
        Private _ErrorMessage As String
        Private aFunction As DataBaseConnectorLibrary.LibraryClass.DataBaseConnection
        Private aWebFunction As DataBaseConnectorLibrary.LibraryClass.DataBaseConnection

        Sub New(ByVal ConnectionString As String)
            aFunction = New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection(ConnectionString)
            '_TheConnectionString = ConnectionString
        End Sub

        Sub New()
            aWebFunction = New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection()
        End Sub

#Region " Get Data "

        Public Function FillGrid(ByVal FilterString As String) As Data.DataSet
            'Dim aFunction As New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection(_TheConnectionString)
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim _TheDataSet As New Data.DataSet

            Try
                With TheOleDbCommand                        ' sets the command query
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT ProductID, ProductName, UnitPrice, " &
                                   "FROM Products " &
                                   "WHERE ProductName LIKE ? + '%' " &
                                   "ORDER BY ProductName"
                    .Parameters.Add("@ProductName", System.Data.OleDb.OleDbType.VarChar).Value = FilterString
                End With

                _TheDataSet = aFunction.ReturnDataSet(TheOleDbCommand)
            Catch ex As Exception
                _ErrorMessage = aFunction.ErrorMessage
            End Try

            Return _TheDataSet
        End Function

        Public Function RetrieveProductNamePrice(ByVal FilterString As String) As OleDb.OleDbDataReader
            'Dim aFunction As New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection(_TheConnectionString)
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim _TheDataReader As OleDb.OleDbDataReader = Nothing

            Try
                With TheOleDbCommand                    ' sets the command query
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT ProductName, UnitPrice " &
                                   "FROM Products " &
                                   "WHERE ProductName LIKE ? + '%' " &
                                   "ORDER BY ProductName"
                    .Parameters.Add("@ProductName", System.Data.OleDb.OleDbType.VarChar).Value = FilterString
                End With

                _TheDataReader = aFunction.ReturnDataReader(TheOleDbCommand)

            Catch ex As Exception
                _ErrorMessage = aFunction.ErrorMessage
            End Try

            Return _TheDataReader
        End Function

        Public Function FillListBox() As Data.DataSet
            'Dim aFunction As New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection(_TheConnectionString)
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim _TheDataSet As New Data.DataSet

            Try
                With TheOleDbCommand                                    ' sets the command query
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT ProductID, ProductName " &
                                   "FROM Products " &
                                   "ORDER BY ProductName"
                End With

                _TheDataSet = aFunction.ReturnDataSet(TheOleDbCommand)

            Catch ex As Exception
                _ErrorMessage = aFunction.ErrorMessage
            End Try

            Return _TheDataSet
        End Function

        Public Function RetrieveCustomers(ByVal FilterString As String) As Data.DataSet
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim _TheDataSet As New Data.DataSet

            Try
                With TheOleDbCommand                        ' sets the command query
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT CustomerID, CompanyName, ContactName, City, Country, Phone, Fax " &
                                   "FROM Customers " &
                                   "WHERE CustomerID LIKE ? + '%' " &
                                   "ORDER BY CustomerID"
                    .Parameters.Add("@CustomerID", System.Data.OleDb.OleDbType.VarChar).Value = FilterString
                End With

                _TheDataSet = aWebFunction.ReturnDataSet(TheOleDbCommand)
            Catch ex As Exception
                _ErrorMessage = aWebFunction.ErrorMessage
            End Try

            Return _TheDataSet
        End Function

        Public Function RetrieveProducts(ByVal FilterString As String) As Data.DataSet
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim _TheDataSet As New Data.DataSet

            Try
                With TheOleDbCommand
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT ProductID, ProductName, UnitPrice, UnitsInStock, UnitsOnOrder " &
                                   "FROM Products " &
                                   "WHERE ProductID BETWEEN 3 AND 20 " &
                                   "ORDER BY ProductName"
                    .Parameters.Add("@ProductID", System.Data.OleDb.OleDbType.VarChar).Value = FilterString
                End With

                _TheDataSet = aWebFunction.ReturnDataSet(TheOleDbCommand)
            Catch ex As Exception
                _ErrorMessage = aWebFunction.ErrorMessage
            End Try

            Return _TheDataSet
        End Function

        Public Function RetrieveSuppliers(ByVal FilterString As String) As Data.DataSet
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim _TheDataSet As New Data.DataSet

            Try
                With TheOleDbCommand
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT SupplierID, CompanyName, ContactName, City, Country, Phone, Fax " &
                                   "FROM Suppliers " &
                                   "WHERE SupplierID BETWEEN 12 AND 28 " &
                                   "ORDER BY CompanyName"
                    .Parameters.Add("@SupplierID", System.Data.OleDb.OleDbType.VarChar).Value = FilterString
                End With

                _TheDataSet = aWebFunction.ReturnDataSet(TheOleDbCommand)
            Catch ex As Exception
                _ErrorMessage = aWebFunction.ErrorMessage
            End Try

            Return _TheDataSet
        End Function
#End Region

#Region " Alter Data "

        Public Function CreateData(ByVal Name As String, ByVal Price As Decimal) As Integer
            'Dim aFunction As New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection(_TheConnectionString)
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim ReturnValue As Integer = Nothing

            With TheOleDbCommand                                    ' sets the command query
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO Products (ProductName, UnitPrice)" &
                                "VALUES(Name, Price)"
                .Parameters.Add("@Name", System.Data.OleDb.OleDbType.VarChar).Value = Name
                .Parameters.Add("@Price", System.Data.OleDb.OleDbType.Decimal).Value = Price
            End With

            ReturnValue = aFunction.AlterData(TheOleDbCommand)
            Return ReturnValue
        End Function

        Public Function UpdateData(ByVal Name As String, ByVal Price As Decimal, ByVal ID As Integer) As Integer
            'Dim aFunction As New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection(_TheConnectionString)
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim ReturnValue As Integer = Nothing

            With TheOleDbCommand
                .CommandType = CommandType.Text
                .CommandText = "UPDATE Products SET [UnitPrice] = @UnitPrice, [ProductName] = @ProductName WHERE [ProductID] = @ProductID"
                .Parameters.Add("@UnitPrice", System.Data.OleDb.OleDbType.Decimal).Value = Price
                .Parameters.Add("@ProductName", System.Data.OleDb.OleDbType.VarChar).Value = Name
                .Parameters.Add("@ProductID", System.Data.OleDb.OleDbType.Integer).Value = ID
            End With

            ReturnValue = aFunction.AlterData(TheOleDbCommand)
            Return ReturnValue
        End Function

        Public Function DeleteData(ByVal Name As String, ByVal ID As Integer) As Integer
            'Dim aFunction As New DataBaseConnectorLibrary.LibraryClass.DataBaseConnection(_TheConnectionString)
            Dim TheOleDbCommand As New System.Data.OleDb.OleDbCommand
            Dim ReturnValue As Integer = Nothing

            With TheOleDbCommand                                    ' sets the command query
                .CommandType = CommandType.Text
                .CommandText = "DELETE FROM Products WHERE ProductID = @ID AND ProductName = @Name"
                .Parameters.Add("@ID", System.Data.OleDb.OleDbType.Integer).Value = ID
                .Parameters.Add("@Name", System.Data.OleDb.OleDbType.VarChar).Value = Name
            End With

            ReturnValue = aFunction.AlterData(TheOleDbCommand)
            Return ReturnValue
        End Function

#End Region

#Region " Properties "
        Public ReadOnly Property ErrorMessage As String
            Get
                Return _ErrorMessage
            End Get
        End Property
#End Region
    End Class
End Namespace