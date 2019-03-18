Imports System.Data.SqlClient
'Imports Microsoft.Office.Interop.Word

Public Class SQLClass1

End Class

Public Class SQLControl
    Private DBCon As New SqlConnection With {.ConnectionString = "Data Source=NJEDIOB01;Initial Catalog=dbDAVMS01;Password=L5tm5insql;Persist Security Info=True;User ID=obaid100"}

    'Data Source=10.50.160.92;Initial Catalog=dbDAVMS01;Persist Security Info=True;User ID=obaid1
    '10.50.160.51
    'imfi-lenovo-111
    'Data Source=NJEDIOB01;Initial Catalog=dbDAVMS01;Persist Security Info=True;User ID=obaid100
    Private DBCmd As SqlCommand

    ' DB DATA
    Public DBDA As SqlDataAdapter
    Public DBDT As DataTable
    Public DBDS As DataSet
    Public AuthUser As String

    Public Function HasConnection() As Boolean
        Try
            DBCon.Open()
            DBCon.Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return False
    End Function

    Public Sub RunQuery(Query As String)
        Try
            DBCon.Open()

            ' CREATE COMMAND
            DBCmd = New SqlCommand(Query, DBCon)

            'FILL DATA SET
            DBDA = New SqlDataAdapter(DBCmd)
            DBDS = New DataSet
            DBDA.Fill(DBDS)

            DBCon.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            ' MAKE SURE CONN IS CLOSE
            DBCon.Close()

        End Try
    End Sub
    ' QUERY PARAMETERS
    Public Params As New List(Of SqlParameter)

    ' QUERY STATISTICS
    Public RecordCount As Integer
    Public Exception As String

    Public Sub New()

    End Sub

    ' ALLOW CONNECTION OVERRIDE
    Public Sub New(ConnectionString As String)
        DBCon = New SqlConnection(ConnectionString)
    End Sub

    ' EXECUTE QUERY SUB
    Public Sub ExecQuery(Query As String, Optional ReturnIdentity As Boolean = False)

        ' RESET QUERY STATICTIS
        RecordCount = 0
        Exception = ""

        Try
            DBCon.Open()

            'CREATE DB COMMAND
            DBCmd = New SqlCommand(Query, DBCon)

            ' LOAD PARAMS INTO DB CMD
            Params.ForEach(Sub(p) DBCmd.Parameters.Add(p))

            ' CLEAR PARAM LIST
            Params.Clear()

            ' EXECUTE CMD & FILL DATA SET
            DBDT = New DataTable
            DBDA = New SqlDataAdapter(DBCmd)
            RecordCount = DBDA.Fill(DBDT)


            If ReturnIdentity = True Then
                '  SCOPE_IDENTITY() ----- Session and Scope
                ' @@IDENTITY ------------ Session only
                'It will work in this db open and close db session
                Dim ReturnQury As String = "SELECT @@IDENTITY As LastID"
                ' IDENT_CURRENT(Table name) ----- Last IDENTITY IN Table, any SCOPE, any SESSION 

                DBCmd = New SqlCommand(ReturnQury, DBCon)
                DBDT = New DataTable
                DBDA = New SqlDataAdapter(DBCmd)
                RecordCount = DBDA.Fill(DBDT)

            End If

        Catch ex As Exception
            'CAPTURE ERROR
            Exception = "ExecQuery Error: " & vbNewLine & ex.Message

        Finally

            ' CLOSE CONNECTION
            If DBCon.State = ConnectionState.Open Then DBCon.Close()

        End Try


    End Sub
    ' ADD PARAMS
    Public Sub AddParam(Name As String, Value As Object)
        Dim NewParam As New SqlParameter(Name, Value)
        Params.Add(NewParam)

    End Sub

    'ERROR CHECKING
    Public Function HasException(Optional Report As Boolean = False) As Boolean
        If String.IsNullOrEmpty(Exception) Then Return False
        If Report = True Then MsgBox(Exception, MsgBoxStyle.Critical, "Exception: ")
        Return True
    End Function

End Class