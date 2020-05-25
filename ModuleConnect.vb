Module ModuleConnect
    Public dbConnection As dbConnector.dbSelector.dbConn
    Public userSchema As String
    Public UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1 As String
    Public UserSchema2, DatabaseName2, DatabaseHost2 As String
    Public DataAdap As OleDb.OleDbDataAdapter
    Public DataRead As System.Data.OleDb.OleDbDataReader
    Public Connect_Command As OleDb.OleDbCommand
    Public trans As OleDb.OleDbTransaction
    Public myDataSet As New DataSet


    Sub ConnectDataBase(ByVal StrSelect As String, ByVal userSchema As String, ByVal DatabaseName As String, ByVal DatabaseHost As String)
        dbConnection = New dbConnector.dbSelector.dbConn
        dbConnection.set_dbConnector(userSchema, DatabaseBrand, DatabaseName, DatabaseHost)
        dbConnection.set_Command(Connect_Command, StrSelect)
    End Sub

End Module
