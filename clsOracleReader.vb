Imports System.Data.OracleClient
Imports System.IO

''' <summary>
''' This class is a backend interface for all the operations on the database
''' </summary>
''' <remarks></remarks>

Public Class clsOracleReader

    ''' <summary>
    ''' Connection to the Oracle database 
    ''' </summary>
    ''' <remarks>This member of the class is shared, because we use the same connection for all the project</remarks>
    Private Shared mConn As OracleConnection

    ''' <summary>
    ''' Command object used to make queries
    ''' </summary>
    ''' <remarks></remarks>
    Private oCommand As New OracleCommand

    ''' <summary>
    ''' Adapter object used to make queries
    ''' </summary>
    ''' <remarks></remarks>
    Private oDataAdapter As OracleDataAdapter

    ''' <summary>
    ''' Dataset object used to make queries
    ''' </summary>
    ''' <remarks></remarks>
    Private oDataset As DataSet

    ''' <summary>
    ''' Event raised at the end of a query, used to bind the datagridview async
    ''' </summary>
    ''' <remarks></remarks>
    Public Event QueryCompleted()

    ''' <summary>
    ''' Dataset Property
    ''' </summary>
    ''' <value></value>
    ''' <returns>Return the dataset after the QueryCompleted() event was raised</returns>
    ''' <remarks>ReadOnly Property</remarks>
    ReadOnly Property Dataset() As DataSet
        Get
            Return oDataset
        End Get
    End Property

    ''' <summary>
    ''' Execute the connection to the database with the provided connection string
    ''' </summary>
    ''' <param name="strConnectionString">Connection string</param>
    ''' <returns>Return the status of the connection</returns>
    ''' <remarks></remarks>
    Public Function ConnectTo(ByVal strConnectionString As String) As Boolean

        mConn = New OracleConnection(strConnectionString)
        Try
            mConn.Open()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        If mConn.State = ConnectionState.Open Then
            ConnectTo = True
        Else : ConnectTo = False
        End If

    End Function

    ''' <summary>
    ''' Get the status of the connection to the database
    ''' </summary>
    ''' <returns>Return a value from the Enum ConnectionState</returns>
    ''' <remarks></remarks>
    Public Function GetState() As ConnectionState
        GetState = mConn.State
    End Function

    ''' <summary>
    ''' Abort the running query
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InterruptQuery()
        oCommand.Cancel()
    End Sub

    ''' <summary>
    ''' Execute a query
    ''' </summary>
    ''' <param name="sqlQuery">SQL Query string</param>
    ''' <remarks></remarks>
    Public Sub ExecuteQuery(ByVal sqlQuery As Object)

        Try
            oCommand.CommandType = CommandType.Text
            oCommand.CommandText = sqlQuery
            oCommand.Connection = mConn
            oCommand.ExecuteNonQuery()

            oDataAdapter = New OracleDataAdapter(oCommand)
            oDataset = New DataSet
            oDataAdapter.Fill(oDataset)

        Catch e As Threading.ThreadAbortException
            MsgBox("Current operation was aborted.", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)

        Finally
            RaiseEvent QueryCompleted()
        End Try

    End Sub

    ''' <summary>
    ''' Clear the internal object after executed a query
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()
        If Not oDataAdapter Is Nothing Then oDataAdapter.Dispose()
        oDataAdapter = Nothing
        If Not oDataset Is Nothing Then oDataset.Dispose()
        oDataset = Nothing
    End Sub

    ''' <summary>
    ''' Close the connection to the database
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Disconnect()
        Try
            If Not mConn Is Nothing Then
                If mConn.State = ConnectionState.Open Then mConn.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
