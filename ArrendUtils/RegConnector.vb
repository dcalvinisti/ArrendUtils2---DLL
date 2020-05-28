Imports Microsoft.Win32


'=====================================
'NOMBRE: RegConnector.
'Clase Pública creada el 24 de marzo del 2015.
'Con el propósito de generar la cadena de conexión a partir del registro del sistema
'Desarrollado por Saul Hernandez <shernandez@leasing.com.gt> 
'como parte del equipo de desarrollo en Arrend.
'=====================================
'MODIFICACIONES
'A continuación se listan las modificaciones más importantes realizadas a la clase:
'FECHA      |MODIFICADO POR     |RÁZÓN Y DESCRIPCIÓN DE LA MODIFICACIÓN
'24/03/15   |Saúl Hernandez     |Creación de la clase.
'=====================================
'COSAS POR HACER:
'...

Public Class RegConnector


    Private rutaRegedit As String
    Private _p1 As String

    Sub New(ByVal p1 As String)
        ' TODO: Complete member initialization 
        Me.rutaKey = p1
        Me.conectaregistro()
    End Sub

    Public Property rutaKey() As String
        Get
            Return rutaRegedit
        End Get
        Set(ByVal value As String)
            rutaRegedit = value
        End Set
    End Property


    Public Property servidor As String = ""
    Public Property bdsap As String = ""
    Public Property usuariodb As String = ""
    Public Property passworddb As String = ""
    Public Property usuariosap As String = ""
    Public Property passwordsap As String = ""
    Public Property tuconexion As String = ""

    Public Sub conectaregistro()

        'Dim KeyPath As String = rutaRegedit
        Dim ConfigCE As New appConfigEditor(rutaRegedit)

        Me.servidor = ConfigCE.getAppSettingValue("Server").ToString

        If ConfigCE.getAppSettingValue("Host").ToString = "UAT" Then
            Me.bdsap = ConfigCE.getAppSettingValue("DBCompanyUAT").ToString  '//AdmonUAT
        Else
            Me.bdsap = ConfigCE.getAppSettingValue("DBCompany").ToString     '//Admon 
        End If
        Me.usuariodb = ConfigCE.getAppSettingValue("DBUser").ToString
        Me.passworddb = ConfigCE.getAppSettingValue("DBPassword").ToString
        Me.usuariosap = ConfigCE.getAppSettingValue("UserSAP").ToString
        Me.passwordsap = ConfigCE.getAppSettingValue("PasswordSAP").ToString

        'Regresa el string de todo el RegWindows.
        Me.tuconexion = Me.servidor & ":" & Me.bdsap & ";" & Me.usuariodb & "," & Me.passworddb & ";" & Me.usuariosap & "," & Me.passwordsap
    End Sub

End Class
