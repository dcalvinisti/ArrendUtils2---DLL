Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class Bitacora
    Private Shared _connStringPattern As String
    Private Shared _connString As String
    Public Shared Property connString As String
        Get
            Return _connString
        End Get
        Set(cadena As String)
            If Regex.IsMatch(cadena, _connStringPattern) Then
                Dim Match As Match = Regex.Match(cadena, _connStringPattern)

                _connString = "Data Source=" + Match.Groups(1).Value + _
                    "; Initial Catalog=" + Match.Groups(2).Value + _
                    "; Uid=" + Match.Groups(3).Value + _
                    "; pwd=" + Match.Groups(4).Value
            Else
                Throw New System.Exception("Cadena mal formateada, usar: server\instance:database;user,password")
            End If
        End Set
    End Property

    Private Shared _connx As SqlConnection
    Public Shared Property connx As SqlConnection
        Get
            If IsNothing(_connx) Then _connx = New SqlConnection

            _connx.ConnectionString = connString

            Return _connx
        End Get
        Set(newConn As SqlConnection)
            If IsNothing(newConn.ConnectionString) Then newConn.ConnectionString = connString

            _connx = newConn
        End Set
    End Property

    Public Sub New(Optional ByVal newConnString = "")
        '   Inicializando propiedades.
        _connStringPattern = "^([\w_]+\\[\w_]+):([\w_]+);([\w_]+),([\w_\:]+)$"

        '   Si no se proporcionó cadena de conexión, se asigna una por defecto
        ' a una base de datos de pruebas, de lo contrario se asigna la que se proporcionó.
        If newConnString = "" Then connString = "arrenddb\sql2008:PruebaLAB;sap_funcional,arrendadora0207" _
        Else connString = newConnString

    End Sub

    Public Sub excepcion(e As System.Exception, _
                         currentUser As String, _
                         modulo As String, _
                         form As String, _
                         Optional ByVal expMsg As String = "", _
                         Optional ByVal ntsMsg As String = "")

        Dim dfltMsg As String = e.GetType.ToString & vbCrLf & e.Message & vbCrLf & e.StackTrace
        Dim comando As New SqlCommand
        Try
            With comando
                .Connection = connx
                .CommandText = <query>
                                   INSERT INTO [Reportes].[dbo].[Excepciones_Log]
                                SELECT GETDATE()            -- Hora
                                    , '<%= currentUser %>'  -- Usuario
                                    , '<%= modulo %>'              -- Modulo
                                    , '<%= form %>'      -- Pantalla
                                    , DB_NAME()             -- db
                                    , '<%= e.GetHashCode %>'    -- codExcepcion
                                    , '<%= If(expMsg = "", dfltMsg, expMsg) %>'  -- msgExcepcion
                                    , '<%= If(ntsMsg = "", dfltMsg, ntsMsg) %>'  -- notas
                               </query>.Value

                connx.Open()

                .ExecuteNonQuery()
            End With
        Catch ex As Exception
            MsgBox(ex.GetHashCode & vbCrLf & ex.Message & vbCrLf & ex.StackTrace)
        Finally
            connx.Close()
            comando.Dispose()
            comando = Nothing
            cleanConnx()
        End Try
    End Sub

    Private Sub cleanConnx()
        connx.Dispose()
        connx = Nothing
    End Sub
End Class
