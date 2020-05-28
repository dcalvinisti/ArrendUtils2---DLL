Imports System.Xml.Linq
Imports System.Xml
Imports System.IO
Imports System.Reflection

Module RutinasConfig

    Public Class appConfigEditor
        Private _configName As String = String.Empty
        Private _configDoc As XDocument = Nothing
        Private _getAppSettings = Nothing
        Private _path As String = String.Empty

        Private Property getAppSettings()
            Get
                Return _getAppSettings
            End Get

            Set(ByVal value)
                _getAppSettings = value
            End Set
        End Property

        Public Sub New(ByVal ConfigurationName As String)
            _configName = ConfigurationName
            loadConfigurationFile(_configName)
        End Sub
        Private Sub loadConfigurationFile(ByVal configName As String)
            _configDoc = New XDocument
            Dim appDir As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName.CodeBase)
            _path = Path.Combine(appDir, configName)
            _configDoc = XDocument.Load(_path)
        End Sub

        Private Function getDescendant(ByVal key As String) As XElement
            Dim getAppSettings = _configDoc.Descendants("appSettings").Elements
            Dim elem As XElement = Nothing

            For Each setting In getAppSettings
                Dim keyAtt As XAttribute = setting.Attribute("key")
                If keyAtt.Value = key Then
                    elem = setting

                    Exit For
                End If
            Next

            Return elem
        End Function

        Public Sub Add(ByVal key As String, ByVal Value As String)
            loadConfigurationFile(_configName)

            Dim element As New XElement("add", New XAttribute("key", key), New XAttribute("value", Value))
            Dim addElement = _configDoc.Descendants("appSettings").Elements

            'addElement.AddAfterSelf(element)

            _configDoc.Save(_path)

        End Sub

        Public Function getAppSettingValue(ByVal key As String) As String
            Dim getAppSettings As XElement = getDescendant(key)
            Dim configValue = getAppSettings.Attribute("value").Value

            Return configValue.ToString
        End Function

        Public Sub setAppSettingValue(ByVal key As String, ByVal newValue As String)
            Dim getAppSettings As XElement = getDescendant(key)

            getAppSettings.SetAttributeValue("value", newValue)
            _configDoc.Save(_path)
        End Sub

    End Class
End Module
