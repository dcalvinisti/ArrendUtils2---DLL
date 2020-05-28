Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Data.Odbc
Imports System.Xml.XPath
Imports System.Xml

Public Class DatosGenereral

    Public Property cServer As String = ""
    Public Property cBaseD As String = ""
    Public Property cUsu As String = ""
    Public Property cPass As String = ""
    Public Property cPathReport As String = ""


    Public Shared Property nMensaje As String = ""
    Public Shared Property SiGrabo As Integer = 0
    'Public Shared Property RequesId As Integer = 0
    'Public Shared Property Success As Boolean = False
    'Public Shared Property Exception As Boolean = False


    Dim ConfigCE As New appConfigEditor("ConfigSQL.xml")

    Dim vServer As String = ConfigCE.getAppSettingValue("ServerSQL").ToString
    Dim vData As String = ConfigCE.getAppSettingValue("DataBase").ToString
    Dim vUsu As String = ConfigCE.getAppSettingValue("DataUser").ToString
    Dim vPass As String = ConfigCE.getAppSettingValue("DataPassword").ToString

    Dim str_conexion As String = "Server=" & vServer & "; Initial Catalog=" & vData & "; User Id=" & vUsu & "; Password=" & vPass & ";"
    Dim conexion As New SqlConnection
    Dim cmd As SqlCommand


    'Constructor
    Public Sub New()
        ' extraer()

        Me.cServer = vServer
        Me.cBaseD = vData
        Me.cUsu = vUsu
        Me.cPass = vPass
    End Sub

    '*Variables ---------------------------------
    Public Property srt_conexion() As String
        Get
            Return Me.str_conexion
        End Get
        Set(ByVal str As String)
            Me.str_conexion = str
        End Set
    End Property

    Public Sub New(ByVal str As String)
        Me.str_conexion = str
    End Sub

    'Para verificar la entrada si hay o no conexion
    Public Function Conecto() As Boolean
        Dim SiHay As Boolean = False
        Try
            conexion.ConnectionString = str_conexion
            conexion.Open()

            If conexion.State() = ConnectionState.Open Then
                conexion.Close()
            End If
            SiHay = True
        Catch ex As Exception
            MsgBox("ERROR: NO HAY CONEXION")
        End Try

        Return SiHay

    End Function

    Public Sub consulta_non_query(ByVal consulta As String)

        'Este metodo recibe como parametro la consulta completa y sirve para hacer INSERT, UPDATE Y DELETE
        conexion.ConnectionString = str_conexion
        cmd = New SqlCommand(consulta, conexion)
        conexion.Open()
        Try
            cmd.ExecuteNonQuery()

            MsgBox("La operacion se realizo con exito!", MsgBoxStyle.Information, "Operacion exitosa!")
        Catch ex As Exception
            ' MsgBox("Error al operar con la base de datos!", MsgBoxStyle.Critical, "Error!")
        End Try
        conexion.Close()
    End Sub


    Public Function consulta_reader(ByVal consulta As String) As DataTable

        'Este metodo recibe como parametro la consulta completa y sirve para hacer SELECT
        Dim dt As New DataTable
        conexion.ConnectionString = str_conexion
        cmd = New SqlCommand(consulta, conexion)
        conexion.Open()

        Try
            dt.Load(cmd.ExecuteReader())
        Catch ex As Exception

            MsgBox("La Tabla o Base Datos es Incorrecta en SQL.. [" & ex.Message & "]", MsgBoxStyle.Critical, "Error!")

        End Try
        conexion.Close()
        Return dt

    End Function

End Class