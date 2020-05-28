Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.CallType
Imports System.Data.SqlClient
Imports SAPbobsCOM

Public Class SAPConnector
    ' The following are properties of the class that can be accesed via get methods
    Private Shared _company As New SAPbobsCOM.Company
    Public Shared Property company() As Company
        Get
            Return _company
        End Get
        Set(ByVal value As Company)
            _company = value
        End Set
    End Property

    Private _regresaMensaje As String

    Public Property regresaMensaje() As String
        Get
            Return _regresaMensaje
        End Get
        Set(ByVal value As String)
            _regresaMensaje = value
        End Set
    End Property


    Private connStringValue As String = ""
    Private currentTransaction As transactionOnHold
    Private connStringPattern As String = ""
    Private fieldValuePattern As String = ""
    Private keysPattern As String = ""
    Private isTransactional As Boolean = True

    'The following is a dictionary that holds the primary keys for the UDT in SAP
    Private primKeys As New Dictionary(Of String, String)
    Private tablesDict As New Dictionary(Of SAPbobsCOM.BoObjectTypes, String)

    Dim Datos As New DatosGenereral()
    Dim dTable As New DataTable
    Dim dTable2 As New DataTable

    Private Structure transactionOnHold
        Public bObject As Object
        Public transactionType As String
    End Structure

    Public Property ConnString() As String
        Get
            Return connStringValue
        End Get
        Set(ByVal value As String)
            Dim match As Match = Regex.Match(value, connStringPattern)
            If match.Success Then
                'If connection string is good, then we can assign the matched groups to their respective properties in the company object
                connStringValue = value

                company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                company.Server = match.Groups(1).Value
                company.CompanyDB = match.Groups(2).Value
                company.DbUserName = match.Groups(3).Value
                company.DbPassword = match.Groups(4).Value

                company.UserName = match.Groups(5).Value
                company.Password = match.Groups(6).Value
            Else
                Throw New System.Exception("Wrong connection string: " & value)
            End If
        End Set
    End Property

    Public Property transactional As Boolean
        Get
            Return isTransactional
        End Get
        Set(value As Boolean)
            isTransactional = value
        End Set
    End Property

    Public ReadOnly Property connected As Boolean
        Get
            Return (Not IsNothing(company)) AndAlso company.Connected
        End Get
    End Property

    Public ReadOnly Property OnHold
        Get
            Return currentTransaction.bObject
        End Get
    End Property

    Public Sub New(Optional ByVal newConnString As String = "")
        Try
            If IsNothing(company) Then company = New SAPbobsCOM.Company
            isTransactional = True

            fieldValuePattern = "([\w_]+) ?= ?'([\w\/@.,_\:\-\$ ]+)'" 'This captures any field='value' pair
            connStringPattern = "^([\w_]+\\[\w_]+):([\w_]+);([\w_]+),([\w_]+);([\w_]+),([\w\:_]+)$"
            keysPattern = "[\w_]+='[\w\.,_ ]+'( and [\w_]+='[\w\.,_ ]+')*" ' a key='value' pair followed by "and" and one or more pairs


            buildKeysDict() 'Building primary keys dictionary
            buildTablesDict() 'Building tables dictionary


            If Not newConnString = "" Then ConnString = newConnString
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
    End Sub

    Public Function connect() As Boolean
        If connected Then Return True
        If IsNothing(company) Then company = New SAPbobsCOM.Company

        Dim retorno As Integer = 0, mensaje As String = ""
        Dim match As Match = Regex.Match(ConnString, connStringPattern)

        ConnString = connStringValue ' Reassigning the connection string
        company.Connect()

        If Not company.Connected Then
            company.GetLastError(retorno, mensaje)
            Throw New System.Exception(mensaje)
            Return False
        End If

        Return True

    End Function

    Public Sub disconnect()
        Dim msg As String = "", rtn As Integer = 0
        If Not IsNothing(company) Then
            company.Disconnect()
            company = Nothing

        End If

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

    End Sub

#Region "SAPBO CRUD" ' This region implements the four basic operations on the SAP Database
    Public Function getBObject(ByVal bObjectType As SAPbobsCOM.BoObjectTypes) As Object
        If Not connected Then connect()

        Dim bObject As Object = company.GetBusinessObject(bObjectType)

        If isTransactional Then disconnect()

        Return bObject
    End Function

    Public Function getBObject(ByVal tName As String) As Object
        If Not connected Then connect()

        Dim usrTable As SAPbobsCOM.UserTable = company.UserTables.Item(tName)

        For Each field In usrTable.UserFields.Fields
            If field.value.ToString <> "" Then field.value = ""
        Next

        If isTransactional Then disconnect()

        Return usrTable
    End Function

    Public Sub removeOnKey(ByRef bObject As Object, ByVal key As String)
        If Not connected Then connect()

        If Not bObject.getBykey(key) Then
            If isTransactional Then disconnect()
            Throw New System.Exception("No se encontró registros")
        End If

        Try
            bObject.remove()
        Catch ex As Exception
            bObject.Cancel()
        End Try

        If isTransactional Then disconnect()
    End Sub

    Public Sub removeObject(ByVal sapType As SAPbobsCOM.BoObjectTypes, ByVal keys As String)
        Dim qString As String = "SELECT * FROM " & tablesDict(sapType) & " WHERE " & keys
        Dim bObject As Object = getBObject(sapType)

        Dim rSet As SAPbobsCOM.Recordset = getRecordSet(qString)
        If rSet.RecordCount = 0 Then Throw New System.Exception("No se encontraron registros")
        rSet.MoveFirst()
        If Not bObject.GetByKey(rSet.Fields.Item(0).Value) Then Throw New System.Exception("No se encontró la tupla")

        bObject.remove()
        Dim rtn As Integer = 0, msg As String = ""
        company.GetLastError(rtn, msg)
        If rtn <> 0 Then Throw New System.Exception(msg)
    End Sub

    Public Sub removeObject(ByVal tName As String, ByVal keys As String)
        Dim table As SAPbobsCOM.UserTable = getBObject(tName)
        Dim rSet As SAPbobsCOM.Recordset = CType(company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        rSet.DoQuery("Select * FROM [@" + tName + "] WHERE " + keys)
        Dim rtn As Integer = 0, msg As String = ""

        While (Not rSet.EoF)
            Dim key As String = rSet.Fields.Item(0).Value.ToString().Trim()
            If (table.GetByKey(key)) Then
                table.Remove()
                company.GetLastError(rtn, msg)
                If rtn <> 0 Then Throw New System.Exception(msg)
            End If
            rSet.MoveNext()
        End While
    End Sub

#End Region

#Region "Utilities"
    Public Function getRecordSet(ByVal query As String) As SAPbobsCOM.Recordset
        If Not connected Then connect()

        Dim rSet As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rSet.DoQuery(query)

        If isTransactional Then disconnect()

        Return rSet
    End Function

    Public Sub commit()
        If Not connected Then connect()

        Dim rtn As Long = 0, msg As String = ""

        Select Case currentTransaction.transactionType
            Case "I"
                currentTransaction.bObject.add()
            Case "U"
                currentTransaction.bObject.update()
            Case "D"
                currentTransaction.bObject.remove()
        End Select

        company.GetLastError(rtn, msg)

        If rtn <> 0 Then
            Throw New System.Exception(msg)
        Else
            DatosGenereral.SiGrabo = 1
        End If

        If isTransactional Then disconnect()
    End Sub

    Private Sub setOnHold(ByRef bObject As Object, ByVal transactionType As String)
        currentTransaction.bObject = bObject
        currentTransaction.transactionType = transactionType
    End Sub


    Private Function containsKeys(ByVal fields As String, ByVal tName As String) As Boolean
        Dim allInFiels As Boolean = True

        For Each key In primKeys(tName).Split(CChar(",")) 'Checks that every key is in fields
            allInFiels = Regex.IsMatch(fields.Trim(), "(^| )" & key & "=") AndAlso allInFiels
        Next key

        Return allInFiels

    End Function

    Private Function checkKeys(ByVal keys As String, ByVal tName As String, ByVal permissive As Boolean) As Boolean

        Return Regex.IsMatch(keys.Trim(), keysPattern) _
                AndAlso containsKeys(keys, tName) _
                AndAlso (permissive OrElse getRecordSet("SELECT * FROM [@" + tName + "] WHERE " + keys).EoF)

    End Function

    Private Sub buildKeysDict()
        primKeys.Add("MContrato", "Code") ' "U_NumEmpresa,U_CodProducto,U_CodSubProducto,U_NumContrato")
        primKeys.Add("MCuota_Contrato", "Code") ' "U_NumEmpresa,U_CodProducto,U_CodSubProducto,U_NumContrato,U_NumCuota")
    End Sub

    Private Sub buildTablesDict()
        tablesDict.Add(SAPbobsCOM.BoObjectTypes.oInvoices, " OINV ")    'Facturas
        tablesDict.Add(SAPbobsCOM.BoObjectTypes.oOrders, " ORDR ")      'Ordenes de venta
        tablesDict.Add(SAPbobsCOM.BoObjectTypes.oBusinessPartners, " OCRD ")    'Socios de negocios
        tablesDict.Add(SAPbobsCOM.BoObjectTypes.oChartOfAccounts, " OJDT ")     'Partidas Contables
        tablesDict.Add(SAPbobsCOM.BoObjectTypes.oIncomingPayments, " ORCT ")    'Recibos
    End Sub
#End Region

#Region "Campos"

    'Private Function BuildDictionary(ByVal _tabla As String) As Dictionary(Of String, Element)
    '    Dim elements As New Dictionary(Of String, Element)

    '    dTable = Datos.consulta_reader("SELECT * FROM dbo.tablasdll WHERE Tabla = '" & _tabla & "'")

    '    If dTable.Rows.Count = 0 Then
    '        MsgBox("No Existen datos !")
    '    Else

    '        'Setea Campos
    '        For Each dRow As DataRow In dTable.Rows

    '            AddToDictionary(elements, dRow.Item("Id").ToString, dRow.Item("Tabla").ToString, dRow.Item("Campo").ToString, dRow.Item("pKey").ToString)

    '        Next
    '    End If

    '    Return elements
    'End Function

    'Private Sub AddToDictionary(ByVal elements As Dictionary(Of String, Element), ByVal eid As Integer, ByVal etabla As String, ByVal ecampo As String, ByVal epkey As Boolean)
    '    Dim theElement As New Element

    '    theElement.eId = eid
    '    theElement.eTabla = etabla
    '    theElement.eCampo = ecampo
    '    theElement.epKey = epkey

    '    elements.Add(key:=theElement.eId, value:=theElement)
    'End Sub

    Public Class Element
        Public Property eId As Integer
        Public Property eTabla As String
        Public Property eCampo As String
        Public Property epKey As Boolean
    End Class

    'Funcion donde se obtine el ultimo numero de Code 
    'Y evitar que se entre en confuciones y que el Code y Name sean unicos y
    'con su correlativo independiente por tabla
    Public Function CorreCode() As String
        Dim Siguiente_Code As String = ""

        dTable = Datos.consulta_reader("SELECT LEFT(convert(varchar, GETDATE() , 112),10) + REPLACE(REPLACE(RIGHT(convert(varchar, getdate(), 121),12),':',''),'.','')  As MaxCode")

        If dTable.Rows.Count > 0 Then

            'Setea Campos
            For Each dRow As DataRow In dTable.Rows
                Siguiente_Code = dRow.Item("MaxCode").ToString
            Next
        End If
        Return Siguiente_Code
    End Function

#End Region

#Region "AutoFacets MyArrendLeasing"

    'PUT de Customer
    'Programador: Saul
    'Fecha: 30/08/2019
    'Actualiza Datos de Clientes
    'Public Sub EditCustomer((ByRef bObject As Object, ByVal content() As String, ByVal key As String)
    Public Sub EditCustomer(ByVal content() As String, ByVal key As String)

        Try
            If Not connected Then connect()
            Dim matches As MatchCollection

            Dim nErr As Long
            Dim errMsg As String = ""
            Dim chk As Integer = 0


            'Create the BusinessPartners object
            Dim vBP As SAPbobsCOM.BusinessPartners
            'Dim elements As Dictionary(Of String, Element) = BuildDictionary("OCRD")
            'Dim GrabaLines(content.Count) As String
            'Dim Linea As Integer = 0
            DatosGenereral.SiGrabo = 0
            Dim paso As Boolean = False

            vBP = getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            If Not vBP.GetByKey(key) Then
                Throw New System.Exception("Código de Cliente NO existe " & key)
            Else

                Try

                    matches = Regex.Matches(content(0), fieldValuePattern)
                    For Each Match As Match In matches
                        'Match.Groups.Item(1).Value = Campo 
                        'Match.Groups.Item(2).Value = 'Valor

                        If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                            vBP.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            CallByName(vBP, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                        End If
                    Next Match

                    paso = True
                Catch ex As Exception
                    DatosGenereral.SiGrabo = 0
                End Try

                'Actualiza datos
                If paso Then
                    If DatosGenereral.SiGrabo = 0 Then

                        chk = vBP.Update()
                        If (chk = 0) Then
                            DatosGenereral.SiGrabo = 1
                        Else
                            DatosGenereral.SiGrabo = 1
                            SAPConnector.company.GetLastError(nErr, errMsg)
                            If (0 <> nErr) Then
                                MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    'Asigna vehiculo a piloto
    'Programador: Gonzalo Morales
    'Fecha: 12/05/2020
    Function AsignVehicle(Customer As String, License As String, Plate As String)
        Dim Vehicle As String = ""
        If Not connected Then connect()
        Dim BP As SAPbobsCOM.BusinessPartners
        BP = getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        Dim CT As SAPbobsCOM.ContactEmployees
        Dim RS As SAPbobsCOM.Recordset
        RS = getRecordSet("select cardcode from ocrd where cardcode = '" & Customer & "'")
        If RS.EoF Then
            Return "Customer does not exist"
        End If
        RS = getRecordSet("select u_identificacion from ocpr where u_identificacion = '" & License & "'")
        If RS.EoF Then
            Return "Driver does not exist"
        End If
        RS = getRecordSet("select itemcode from itm13 where attritxt1 = '" & Plate & "'")
        If RS.EoF Then
            Return "Vehicle does not exist"
        End If
        Vehicle = RS.Fields.Item(0).Value
        RS = getRecordSet("select u_itemcode from ocpr where u_itemcode = '" & Vehicle & "' and u_identificacion != '" & License & "'")
        If Not RS.EoF Then
            Return "Vehicle has already been asigned to another driver"
        End If
        If BP.GetByKey(Customer) Then
            CT = BP.ContactEmployees
            If CT.Count > 0 Then
                For i As Integer = 0 To CT.Count - 1
                    CT.SetCurrentLine(i)
                    If BP.ContactEmployees.UserFields.Fields.Item("U_Identificacion").Value = License Then
                        Exit For
                    End If
                Next
                BP.ContactEmployees.UserFields.Fields.Item("U_Itemcode").Value = Vehicle
                Dim Retval As Integer = BP.Update
                If Retval <> 0 Then
                    Dim Msg As String = company.GetLastErrorDescription
                    Return Msg
                End If
            End If
        End If
        Return "Vehicle has been assigned successfully"
    End Function

    'PUT de Driver 
    'Programador: Saul
    'Fecha: 30/08/2019
    'Actualiza Datos a Empleados (Contactos de Clientes)
    Function EditDriver(ByVal content() As String, ByVal _Cliente As String, ByVal _Licencia As String, ByVal _Contrato As String, ByVal _Placa As String) As String
        If Not connected Then connect()
        Dim matches As MatchCollection

        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0
        Dim NickName As String = ""
        Dim eMail As String = ""

        'Create the BusinessPartners object
        Dim vBP As SAPbobsCOM.BusinessPartners
        Dim sboContacts As SAPbobsCOM.ContactEmployees
        'Dim elements As Dictionary(Of String, Element) = BuildDictionary("OCPR")
        'Dim GrabaLines(content.Count) As String
        'Dim Linea As Integer = 0
        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim pasodos As Boolean = False
        Dim Vehiculo As String = ""
        Dim SQLstring As String = ""
        DatosGenereral.SiGrabo = 0

        vBP = getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        sboContacts = vBP.ContactEmployees

        'Verifica si Existe el Cliente su llave Primaria
        If vBP.GetByKey(_Cliente) Then

            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then
                'Busco el Nombre
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches

                    If Match.Groups.Item(1).Value.ToUpper = "E_MAIL" Then
                        eMail = Match.Groups.Item(2).Value

                        DatosGenereral.SiGrabo = 0
                        DatosGenereral.nMensaje = ""
                        Exit For
                    Else
                        DatosGenereral.SiGrabo = 3
                        DatosGenereral.nMensaje = "The E_Mail field is not included.."
                    End If

                Next Match
            End If

            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then

                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select name From [OCPR] WHERE LTRIM(RTRIM(E_MailL)) = '" & eMail.Trim & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        NickName = rSet2.Fields.Item(0).Value

                        'DatosGenereral.SiGrabo = 5
                        'DatosGenereral.nMensaje = "The E_Mail driver already exist in Customer " & rSet2.Fields.Item(1).Value

                        rSet2.MoveNext()
                    End While
                End If
            End If


            ''Ubica el campo nombre dentro del Arreglo (content)
            'If DatosGenereral.SiGrabo = 0 Then


            '    ''Busco el Nombre
            '    'matches = Regex.Matches(content(0), fieldValuePattern)
            '    'For Each Match As Match In matches

            '    '    If Match.Groups.Item(1).Value.ToUpper = "NAME" Then
            '    '        NickName = Match.Groups.Item(2).Value

            '    '        Exit For
            '    '    Else
            '    '        DatosGenereral.SiGrabo = 2
            '    '        DatosGenereral.nMensaje = "The Name field is not included.."
            '    '    End If

            '    'Next Match

            '    NickName = vBP.ContactPerson
            'End If


            'Verifica si Existe la Placa o y Contrato
            'If DatosGenereral.SiGrabo = 0 Then

            '    SQLstring = "SELECT a.ItemCode FROM ITM13 a"
            '    SQLstring = SQLstring & " INNER Join OITM b ON b.ItemCode = a.ItemCode"
            '    SQLstring = SQLstring & " INNER Join [@MCONTRATO] c ON c.U_NumContrato = b.U_Contrato"
            '    SQLstring = SQLstring & " WHERE a.AttriTxt1 = '" & _Placa & "'"
            '    SQLstring = SQLstring & " AND c.U_NumContrato = '" & _Contrato & "'"
            '    SQLstring = SQLstring & " And c.U_CodEstado = 'F'"

            '    Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            '    If rSet.RecordCount > 0 Then
            '        rSet.MoveFirst()
            '        While Not rSet.EoF

            '            Vehiculo = rSet.Fields.Item(0).Value

            '            rSet.MoveNext()
            '        End While
            '    Else
            '        DatosGenereral.SiGrabo = 3
            '        DatosGenereral.nMensaje = "The license plate is incorrect or Contract doesn't exist.."
            '    End If
            'End If

            ''Busca si ya existe la llave primaria Name
            'If DatosGenereral.SiGrabo = 0 Then
            '    'Busco el Nombre
            '    matches = Regex.Matches(content(0), fieldValuePattern)
            '    For Each Match As Match In matches

            '        If Match.Groups.Item(1).Value.ToUpper = "E_MAIL" Then
            '            eMail = Match.Groups.Item(2).Value

            '            DatosGenereral.SiGrabo = 0
            '            DatosGenereral.nMensaje = ""
            '            Exit For
            '        Else
            '            DatosGenereral.SiGrabo = 3
            '            DatosGenereral.nMensaje = "The E_Mail field is not included.."
            '        End If

            '    Next Match
            'End If

            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then

                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select E_MailL, name From [OCPR] WHERE CardCode = '" & _Cliente & "'  AND LTRIM(RTRIM(E_MailL)) = '" & eMail.Trim & "' AND LTRIM(RTRIM(U_Identificacion)) != '" & _Licencia & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        DatosGenereral.SiGrabo = 5
                        DatosGenereral.nMensaje = "The E_Mail driver already exist in Customer " & rSet2.Fields.Item(1).Value

                        rSet2.MoveNext()
                    End While
                End If
            End If


            If DatosGenereral.SiGrabo = 0 Then
                sboContacts = vBP.ContactEmployees
                If sboContacts.Count > 0 Then
                    For i As Integer = 0 To sboContacts.Count - 1
                        sboContacts.SetCurrentLine(i)
                        'If sboContacts.Name = NickName Then
                        If vBP.ContactEmployees.UserFields.Fields.Item("U_Identificacion").Value = _Licencia Then
                            pasodos = True
                            i = sboContacts.Count
                        End If
                    Next
                End If
            End If

            'Verifica si Existe El Drive    
            If pasodos = False Then

                If DatosGenereral.SiGrabo = 0 Then
                    DatosGenereral.SiGrabo = 4
                    DatosGenereral.nMensaje = "The driver doesn't exist.."
                End If
            Else
                Dim DRV As SAPbobsCOM.Recordset
                DRV = getRecordSet("select u_itemcode from ocpr where u_identificacion = '" & _Licencia & "'")
                If Not DRV.EoF Then
                    Vehiculo = DRV.Fields.Item(0).Value
                End If
                Try
                    matches = Regex.Matches(content(0), fieldValuePattern)
                    For Each Match As Match In matches
                        'Match.Groups.Item(1).Value = Campo 
                        'Match.Groups.Item(2).Value = 'Valor

                        If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                            vBP.ContactEmployees.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            If Match.Groups.Item(1).Value = "Gender" Then
                                If Match.Groups.Item(2).Value = "M" Then
                                    vBP.ContactEmployees.Gender = BoGenderTypes.gt_Male
                                ElseIf Match.Groups.Item(2).Value = "F" Then
                                    vBP.ContactEmployees.Gender = BoGenderTypes.gt_Female
                                End If
                            ElseIf Match.Groups.Item(1).Value = "Active" Then
                                If Match.Groups.Item(2).Value = "N" Then
                                    vBP.ContactEmployees.Active = BoYesNoEnum.tNO
                                    Dim RVEH As SAPbobsCOM.Recordset
                                    RVEH = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Try
                                        RVEH.DoQuery("update oitm set u_lastdriver = '" & NickName & "' where itemcode = '" & Vehiculo & "'")
                                    Catch ex As Exception
                                        DatosGenereral.SiGrabo = 6
                                        DatosGenereral.nMensaje = ex.Message
                                    End Try
                                    Vehiculo = ""
                                Else
                                    vBP.ContactEmployees.Active = BoYesNoEnum.tYES
                                End If
                            Else
                                CallByName(vBP.ContactEmployees, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                            End If
                        End If
                    Next Match

                    paso = True

                Catch ex As Exception

                    DatosGenereral.SiGrabo = 5
                    DatosGenereral.nMensaje = "Error SAP: " + ex.Message

                    paso = False

                End Try
            End If


            'Actualiza datos
            If paso Then
                If DatosGenereral.SiGrabo = 0 Then

                    vBP.ContactEmployees.UserFields.Fields.Item("U_Itemcode").Value = Vehiculo

                    chk = vBP.Update()
                    If (chk = 0) Then
                        DatosGenereral.SiGrabo = 1
                        DatosGenereral.nMensaje = "The driver was successfully updated "
                    Else
                        SAPConnector.company.GetLastError(nErr, errMsg)

                        If (0 <> nErr) Then
                            ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                            DatosGenereral.SiGrabo = 8
                            DatosGenereral.nMensaje = "Error SAP:" + Str(nErr) + "," + errMsg

                        End If
                    End If
                End If
            End If
        Else
            DatosGenereral.SiGrabo = 2
            DatosGenereral.nMensaje = "The customer doesn't exist " & _Cliente
        End If

        Return DatosGenereral.nMensaje
    End Function

    'POST de Driver 
    'Programador: Saul
    'Fecha: 30/08/2019
    'Graba Datos a Empleados (Contactos de Clientes)
    Function AddDriver(ByVal content() As String, ByVal _Cliente As String, ByVal _Licencia As String, ByVal _Contrato As String, ByVal _Placa As String) As String
        If Not connected Then connect()
        Dim matches As MatchCollection

        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0
        Dim NickName As String = ""
        Dim eMail As String = ""

        'Create the BusinessPartners object
        Dim vBP As SAPbobsCOM.BusinessPartners
        Dim sboContacts As SAPbobsCOM.ContactEmployees
        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim Vehiculo As String = ""
        Dim SQLstring As String = ""
        DatosGenereral.SiGrabo = 0

        vBP = getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        sboContacts = vBP.ContactEmployees

        'Verifica si Existe el Cliente su llave Primaria
        If vBP.GetByKey(_Cliente) Then

            'Ubica el campo nombre dentro del Arreglo (content)
            If DatosGenereral.SiGrabo = 0 Then

                'Busco el Nombre
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches

                    If Match.Groups.Item(1).Value.ToUpper = "NAME" Then
                        NickName = Match.Groups.Item(2).Value

                        DatosGenereral.SiGrabo = 0
                        DatosGenereral.nMensaje = ""
                        Exit For
                    Else
                        DatosGenereral.SiGrabo = 3
                        DatosGenereral.nMensaje = "The Name field is not included.."
                    End If

                Next Match

            End If

            'Ubica el campo E_mail dentro del Arreglo (content)
            If DatosGenereral.SiGrabo = 0 Then

                'Busco el Nombre
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches

                    If Match.Groups.Item(1).Value.ToUpper = "E_MAIL" Then
                        eMail = Match.Groups.Item(2).Value

                        DatosGenereral.SiGrabo = 0
                        DatosGenereral.nMensaje = ""
                        Exit For
                    Else
                        DatosGenereral.SiGrabo = 3
                        DatosGenereral.nMensaje = "The E_Mail field is not included.."
                    End If

                Next Match

            End If

            'Verifica si Existe la Placa o y Contrato
            'If DatosGenereral.SiGrabo = 0 Then

            '    SQLstring = "SELECT a.ItemCode FROM ITM13 a"
            '    SQLstring = SQLstring & " INNER Join OITM b ON b.ItemCode = a.ItemCode"
            '    SQLstring = SQLstring & " INNER Join [@MCONTRATO] c ON c.U_NumContrato = b.U_Contrato"
            '    SQLstring = SQLstring & " WHERE a.AttriTxt1 = '" & _Placa & "'"
            '    SQLstring = SQLstring & " AND c.U_NumContrato = '" & _Contrato & "'"
            '    SQLstring = SQLstring & " And c.U_CodEstado = 'F'"

            '    Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            '    If rSet.RecordCount > 0 Then
            '        rSet.MoveFirst()
            '        While Not rSet.EoF

            '            Vehiculo = rSet.Fields.Item(0).Value

            '            rSet.MoveNext()
            '        End While
            '    Else
            '        DatosGenereral.SiGrabo = 4
            '        DatosGenereral.nMensaje = "The license plate is incorrect or Contract |'t exist.."
            '    End If

            'End If


            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then


                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select Name from [OCPR] WHERE CardCode = '" & _Cliente & "'  AND LTRIM(RTRIM(Name)) = '" & NickName.Trim & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        DatosGenereral.SiGrabo = 5
                        DatosGenereral.nMensaje = "The Name driver already exist.."

                        rSet2.MoveNext()
                    End While
                End If
            End If

            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then

                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select E_MailL From [OCPR] WHERE CardCode = '" & _Cliente & "'  AND LTRIM(RTRIM(E_MailL)) = '" & eMail.Trim & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        DatosGenereral.SiGrabo = 5
                        DatosGenereral.nMensaje = "The E_Mail driver for customer " & _Cliente & " already exist.."

                        rSet2.MoveNext()
                    End While
                End If
            End If

            'Validacion del Piloto por ItemCode
            If DatosGenereral.SiGrabo = 0 Then
                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select U_Identificacion from [OCPR] WHERE CardCode = '" & _Cliente & "'  AND Position = 'Driver' AND U_Identificacion ='" & _Licencia & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        DatosGenereral.SiGrabo = 6
                        DatosGenereral.nMensaje = "The Id driver already exist.."

                        rSet2.MoveNext()
                    End While
                End If
            End If

            If DatosGenereral.SiGrabo = 0 Then
                Try

                    'If (vBP.ContactEmployees.Count = 0) Then
                    vBP.ContactEmployees.Add()
                    'Else
                    '    If (vBP.ContactEmployees.Count > 0) Then
                    '        vBP.ContactEmployees.Add()
                    '    End If
                    'End If
                    vBP.ContactEmployees.Name = NickName
                    vBP.ContactEmployees.UserFields.Fields.Item("U_Identificacion").Value = _Licencia
                    vBP.ContactEmployees.UserFields.Fields.Item("U_Itemcode").Value = Vehiculo
                    vBP.ContactEmployees.UserFields.Fields.Item("U_PaisDireccion").Value = "1"

                    vBP.ContactEmployees.Title = "Piloto"
                    vBP.ContactEmployees.Position = "Driver"
                    vBP.ContactEmployees.Active = BoYesNoEnum.tYES

                    matches = Regex.Matches(content(0), fieldValuePattern)
                    For Each Match As Match In matches
                        'Match.Groups.Item(1).Value = Campo 
                        'Match.Groups.Item(2).Value = 'Valor
                        If Match.Groups.Item(1).Value = "Active" Then
                            Continue For
                        End If

                        If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                            vBP.ContactEmployees.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            If Match.Groups.Item(1).Value = "Gender" Then
                                If Match.Groups.Item(2).Value = "M" Then
                                    vBP.ContactEmployees.Gender = BoGenderTypes.gt_Male
                                ElseIf Match.Groups.Item(2).Value = "F" Then
                                    vBP.ContactEmployees.Gender = BoGenderTypes.gt_Female
                                End If
                            Else
                                CallByName(vBP.ContactEmployees, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                            End If
                        End If
                    Next Match

                    paso = True

                Catch ex As Exception

                    DatosGenereral.SiGrabo = 7
                    DatosGenereral.nMensaje = "Error SAP: " + ex.Message

                    paso = False

                End Try

                'Actualiza datos
                If paso Then
                    If DatosGenereral.SiGrabo = 0 Then

                        chk = vBP.Update()
                        If (chk = 0) Then
                            DatosGenereral.SiGrabo = 1
                            DatosGenereral.nMensaje = "The driver was successfully added"
                        Else
                            SAPConnector.company.GetLastError(nErr, errMsg)


                            If (0 <> nErr) Then
                                ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                                DatosGenereral.SiGrabo = 8
                                DatosGenereral.nMensaje = "Error SAP:" + Str(nErr) + "," + errMsg

                            End If
                        End If
                    End If
                End If

            End If
        Else
            DatosGenereral.SiGrabo = 2
            DatosGenereral.nMensaje = "The customer doesn't exist " & _Cliente
        End If

        Return DatosGenereral.nMensaje
    End Function

    'PUT de Sucursal  
    'Programador: Saul
    'Fecha: 30/08/2019
    'Actualiza Datos a Empleados (Contactos de Clientes)
    Function EditOutlet(ByVal content() As String, ByVal _Cliente As String, ByVal _Tipo As String, ByVal _Codigo As Integer) As String
        If Not connected Then connect()
        Dim matches As MatchCollection

        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0
        Dim NickName As String = ""
        Dim eMail As String = ""

        'Create the BusinessPartners object
        Dim vBP As SAPbobsCOM.BusinessPartners
        Dim sboContacts As SAPbobsCOM.ContactEmployees
        'Dim elements As Dictionary(Of String, Element) = BuildDictionary("OCPR")
        'Dim GrabaLines(content.Count) As String
        'Dim Linea As Integer = 0
        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim pasodos As Boolean = False
        Dim Vehiculo As String = ""
        Dim SQLstring As String = ""
        Dim nTitle As String = ""
        Dim nPosition As String = ""
        DatosGenereral.SiGrabo = 0

        If _Tipo = "Outlet" Then
            nPosition = "Outlet"
            nTitle = "Sucursal"
        Else
            nPosition = "Executive"
            nTitle = "Asesor"
        End If

        vBP = getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        sboContacts = vBP.ContactEmployees

        'Verifica si Existe el Cliente su llave Primaria
        If vBP.GetByKey(_Cliente) Then

            'Ubica el campo nombre dentro del Arreglo (content)
            If DatosGenereral.SiGrabo = 0 Then

                'Busco el Nombre
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches

                    If Match.Groups.Item(1).Value.ToUpper = "NAME" Then
                        NickName = Match.Groups.Item(2).Value

                        Exit For
                    Else
                        DatosGenereral.SiGrabo = 2
                        DatosGenereral.nMensaje = "The Name field is not included.."
                    End If

                Next Match
            End If

            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then

                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select E_MailL From [OCPR] WHERE CardCode = '" & _Cliente & "'  AND LTRIM(RTRIM(E_MailL)) = '" & eMail.Trim & "' AND LTRIM(RTRIM(Name)) != '" & NickName & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        DatosGenereral.SiGrabo = 5
                        DatosGenereral.nMensaje = "The E_Mail " & nTitle & " already exist.."

                        rSet2.MoveNext()
                    End While
                End If
            End If

            If DatosGenereral.SiGrabo = 0 Then
                sboContacts = vBP.ContactEmployees
                If sboContacts.Count > 0 Then
                    For i As Integer = 1 To sboContacts.Count - 1
                        sboContacts.SetCurrentLine(i)
                        If sboContacts.InternalCode = _Codigo Then
                            pasodos = True
                            i = sboContacts.Count
                        End If
                    Next
                End If
            End If

            'Verifica si Existe El Drive    
            If pasodos = False Then
                DatosGenereral.SiGrabo = 4
                DatosGenereral.nMensaje = "The " & nTitle & " doesn't exist.."
            Else

                Try
                    matches = Regex.Matches(content(0), fieldValuePattern)
                    For Each Match As Match In matches
                        'Match.Groups.Item(1).Value = Campo 
                        'Match.Groups.Item(2).Value = 'Valor

                        If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                            vBP.ContactEmployees.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            CallByName(vBP.ContactEmployees, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                        End If
                    Next Match

                    paso = True

                Catch ex As Exception

                    DatosGenereral.SiGrabo = 5
                    DatosGenereral.nMensaje = "Error SAP: " + ex.Message

                    paso = False

                End Try
            End If

            'Actualiza datos
            If paso Then
                If DatosGenereral.SiGrabo = 0 Then

                    vBP.ContactEmployees.UserFields.Fields.Item("U_Itemcode").Value = Vehiculo

                    chk = vBP.Update()
                    If (chk = 0) Then
                        DatosGenereral.SiGrabo = 1
                        DatosGenereral.nMensaje = "The " & nTitle & " was successfully updated "
                    Else
                        SAPConnector.company.GetLastError(nErr, errMsg)

                        If (0 <> nErr) Then
                            ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                            DatosGenereral.SiGrabo = 8
                            DatosGenereral.nMensaje = "Error SAP:" + Str(nErr) + "," + errMsg

                        End If
                    End If
                End If
            End If
        Else
            DatosGenereral.SiGrabo = 2
            DatosGenereral.nMensaje = "The supplier doesn't exist " & _Cliente
        End If

        Return DatosGenereral.nMensaje
    End Function

    'POST de Sucursal 
    'Programador: Saul
    'Fecha: 30/08/2019
    'Graba Datos a Empleados (Contactos de Clientes)
    Function AddOutlet(ByVal content() As String, ByVal _Cliente As String, ByVal _Tipo As String) As String
        If Not connected Then connect()
        Dim matches As MatchCollection

        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0
        Dim NickName As String = ""
        Dim eMail As String = ""

        'Create the BusinessPartners object
        Dim vBP As SAPbobsCOM.BusinessPartners
        Dim sboContacts As SAPbobsCOM.ContactEmployees
        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim Vehiculo As String = ""
        Dim SQLstring As String = ""
        Dim nTitle As String = ""
        Dim nPosition As String = ""
        DatosGenereral.SiGrabo = 0

        If _Tipo = "Outlet" Then
            nPosition = "Outlet"
            nTitle = "Sucursal"
        Else
            nPosition = "Executive"
            nTitle = "Asesor"
        End If

        vBP = getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        sboContacts = vBP.ContactEmployees

        'Verifica si Existe el Cliente su llave Primaria
        If vBP.GetByKey(_Cliente) Then

            'Ubica el campo nombre dentro del Arreglo (content)
            If DatosGenereral.SiGrabo = 0 Then

                'Busco el Nombre
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches

                    If Match.Groups.Item(1).Value.ToUpper = "NAME" Then
                        NickName = Match.Groups.Item(2).Value

                        DatosGenereral.SiGrabo = 0
                        DatosGenereral.nMensaje = ""
                        Exit For
                    Else
                        DatosGenereral.SiGrabo = 3
                        DatosGenereral.nMensaje = "The Name field is not included.."
                    End If

                Next Match

            End If

            'Ubica el campo E_mail dentro del Arreglo (content)
            If DatosGenereral.SiGrabo = 0 Then

                'Busco el Nombre
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches

                    If Match.Groups.Item(1).Value.ToUpper = "E_MAIL" Then
                        eMail = Match.Groups.Item(2).Value

                        DatosGenereral.SiGrabo = 0
                        DatosGenereral.nMensaje = ""
                        Exit For
                    Else
                        DatosGenereral.SiGrabo = 3
                        DatosGenereral.nMensaje = "The E_Mail field is not included.."
                    End If

                Next Match

            End If


            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then

                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select Name from [OCPR] WHERE CardCode = '" & _Cliente & "'  AND LTRIM(RTRIM(Name)) = '" & NickName.Trim & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        DatosGenereral.SiGrabo = 5
                        If _Tipo = "Outlet" Then
                            DatosGenereral.nMensaje = "The Name " & nTitle & " already exist.."
                        Else
                        End If

                        rSet2.MoveNext()
                    End While
                End If
            End If

            'Busca si ya existe la llave primaria Name
            If DatosGenereral.SiGrabo = 0 Then


                Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet("Select E_MailL From [OCPR] WHERE CardCode = '" & _Cliente & "'  AND LTRIM(RTRIM(E_MailL)) = '" & eMail.Trim & "'")

                If rSet2.RecordCount > 0 Then
                    rSet2.MoveFirst()

                    While Not rSet2.EoF

                        DatosGenereral.SiGrabo = 5
                        DatosGenereral.nMensaje = "The E_Mail " & nTitle & " already exist.."

                        rSet2.MoveNext()
                    End While
                End If
            End If

            If DatosGenereral.SiGrabo = 0 Then
                Try

                    vBP.ContactEmployees.Add()
                    vBP.ContactEmployees.Name = NickName
                    vBP.ContactEmployees.Title = nTitle
                    vBP.ContactEmployees.Position = nPosition

                    matches = Regex.Matches(content(0), fieldValuePattern)
                    For Each Match As Match In matches
                        'Match.Groups.Item(1).Value = Campo 
                        'Match.Groups.Item(2).Value = 'Valor

                        If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                            vBP.ContactEmployees.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            CallByName(vBP.ContactEmployees, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                        End If
                    Next Match

                    paso = True

                Catch ex As Exception

                    DatosGenereral.SiGrabo = 7
                    DatosGenereral.nMensaje = "Error SAP: " + ex.Message

                    paso = False

                End Try

                'Actualiza datos
                If paso Then
                    If DatosGenereral.SiGrabo = 0 Then

                        chk = vBP.Update()
                        If (chk = 0) Then
                            DatosGenereral.SiGrabo = 1
                            DatosGenereral.nMensaje = "The " & nTitle & " was successfully added"
                        Else
                            SAPConnector.company.GetLastError(nErr, errMsg)

                            If (0 <> nErr) Then
                                ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                                DatosGenereral.SiGrabo = 8
                                DatosGenereral.nMensaje = "Error SAP:" + Str(nErr) + "," + errMsg

                            End If
                        End If
                    End If
                End If

            End If
        Else
            DatosGenereral.SiGrabo = 2
            DatosGenereral.nMensaje = "The supplier doesn't exist " & _Cliente
        End If

        Return DatosGenereral.nMensaje
    End Function


    'POST de Fuel Card
    'Programador: Saul
    'Fecha: 04/09/2019
    'Graba Datos Vales de Combustible
    Function AddFuelCard(ByVal content() As String, ByVal _Contrato As String, ByVal _Placa As String) As String
        If Not connected Then connect()
        Dim matches As MatchCollection

        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        'Create the BusinessPartners object
        Dim uTable As SAPbobsCOM.UserTable = getBObject("FUELCARD")
        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim MaxCorrela As String = CorreCode()
        Dim pasouno As Boolean = False
        DatosGenereral.nMensaje = ""

        'Verifica se Agrego el Name (es la llave del Driver)

        Dim Vehiculo As String = ""
        Dim FechaUltima As String = ""
        Dim SQLstring As String = ""

        SQLstring = "SELECT a.ItemCode,c.U_FechaPago FROM ITM13 a "
        SQLstring = SQLstring & " INNER Join OITM b ON b.ItemCode = a.ItemCode "
        SQLstring = SQLstring & " INNER Join("
        SQLstring = SQLstring & "				SELECT TOP 1 cd.U_FechaPago, cc.U_NumContrato, cc.U_CodEstado FROM [@MCUOTA_CONTRATO] cd"
        SQLstring = SQLstring & "                INNER Join [@MCONTRATO] cc ON cc.U_NumContrato = cd.U_NumContrato"
        SQLstring = SQLstring & "                WHERE cc.U_NumContrato ='" & _Contrato & "'"
        SQLstring = SQLstring & "                ORDER BY U_NumCuota desc "
        SQLstring = SQLstring & "				) c  ON c.U_NumContrato = b.U_Contrato"
        SQLstring = SQLstring & " WHERE a.AttriTxt1 = '" & _Placa & "'"
        SQLstring = SQLstring & " AND c.U_NumContrato = '" & _Contrato & "'"
        SQLstring = SQLstring & " AND c.U_CodEstado = 'F'"

        Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

        If rSet.RecordCount > 0 Then
            rSet.MoveFirst()
            While Not rSet.EoF

                Vehiculo = rSet.Fields.Item(0).Value
                FechaUltima = rSet.Fields.Item(1).Value
                pasouno = True

                rSet.MoveNext()
            End While
        End If

        If pasouno Then

            Try
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches
                    'Match.Groups.Item(1).Value = Campo 
                    'Match.Groups.Item(2).Value = 'Valor

                    If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                        uTable.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                    Else
                        If Match.Groups.Item(1).Value = "Code" Or Match.Groups.Item(1).Value = "Name" Then
                            CallByName(uTable, Match.Groups.Item(1).Value, [Let], MaxCorrela)
                        Else
                            CallByName(uTable, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                        End If
                    End If
                Next Match

                paso = True
            Catch ex As Exception

                DatosGenereral.SiGrabo = 2
                DatosGenereral.nMensaje = "Error SAP: " + ex.Message

                paso = False

            End Try

            'Actualiza datos
            If paso Then
                If DatosGenereral.SiGrabo = 0 Then

                    'Actualiza la Placa para que sierva de relacion con el Vehiculo
                    uTable.UserFields.Fields.Item("U_LicensePlate").Value = _Placa

                    'If _CardNumber.Trim = "" Then
                    '    _CardNumber = MaxCorrela
                    'End If
                    uTable.UserFields.Fields.Item("U_FuelCardNumber").Value = MaxCorrela
                    uTable.UserFields.Fields.Item("U_ExpiryDate").Value = FechaUltima
                    uTable.UserFields.Fields.Item("U_Status").Value = "Active"

                    chk = uTable.Add()
                    If (chk = 0) Then
                        DatosGenereral.SiGrabo = 1
                        DatosGenereral.nMensaje = "Fuel Card was successfully recorded "
                    Else
                        SAPConnector.company.GetLastError(nErr, errMsg)

                        If (0 <> nErr) Then
                            ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                            DatosGenereral.SiGrabo = 4
                            DatosGenereral.nMensaje = "Error SAP:" + Str(nErr) + "," + errMsg

                        End If
                    End If

                End If
            End If
        Else
            DatosGenereral.SiGrabo = 3
            DatosGenereral.nMensaje = "The license plate is incorrect or Contract doesn't exist.."
        End If

        Return DatosGenereral.nMensaje
    End Function
#End Region

#Region "Call Center"

    'POST de Call Center Addrequest
    'Programador: Saul HA
    'Fecha: 01/06/2020
    'Llamada de Servicio
    Function CallCenterAddrequest(ByVal content() As String, ByVal contentLine() As String, _Contrato As String, _Placa As String, _DealerReference As String, _CreatedBy As String, ByRef _Exception As Boolean, ByRef _NoDocEntry As Integer, _RequestComeFrom As String, _Action As String, ByRef _SuccessSaved As Boolean, ByRef _WarningMsg As String) As String
        If Not connected Then connect()

        Dim matches As MatchCollection

        Dim nErr As Long = 0
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        'Create the BusinessPartners object
        Dim vOrder As SAPbobsCOM.Documents

        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim lngKey As String = ""

        Dim nTotalLinea As Double = 0.00
        Dim nPrecio As Double = 0.00
        Dim nCanti As Integer = 0.00
        Dim SQLstring As String = ""
        Dim Proveedor As String = ""
        Dim CanPo As Integer = 0
        Dim CanOutPo As Integer = 0
        Dim TotPo As Double = 0.00
        Dim Meses As Integer = 0
        Dim MileContrato As Integer = 0
        Dim MileConsumidas As Integer = 0
        Dim FechaPO As String = ""
        Dim CuenConta As String = ""
        Dim TipoCuenta As String = ""
        Dim EstadoSolicita As String = ""
        Dim Moneda As String = ""
        Dim NomProvee As String = ""
        Dim CodigoClie As String = ""
        Dim Direccion As String = ""
        Dim ActivoFijo As String = ""
        Dim SerieUnida As String = ""

        If _Action = "Add" Then
            _NoDocEntry = 0
        End If

        'Crea el Objeto
        vOrder = getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)


        'Busco el Nombre
        If DatosGenereral.SiGrabo = 0 Then
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "U_ESTADOPO" Then
                    EstadoSolicita = Match.Groups.Item(2).Value

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                    Exit For
                Else
                    DatosGenereral.SiGrabo = 27
                    DatosGenereral.nMensaje = "[027 Call Center] The U_EstadoPO field is not included"
                End If

            Next Match

        End If

        If DatosGenereral.SiGrabo = 0 Then

            If EstadoSolicita = "0009" Then
                DatosGenereral.SiGrabo = 28
                DatosGenereral.nMensaje = "[028 Call Center] The Status is Canceled, it can't be modified "
            End If

        End If


        If DatosGenereral.SiGrabo = 0 Then

            If Not vOrder.GetByKey(_NoDocEntry) And _Action = "Update" Then
                DatosGenereral.SiGrabo = 26
                DatosGenereral.nMensaje = "[026 Call Center] The Purchase Order Does Not Exist "
            End If

        End If

        If DatosGenereral.SiGrabo = 0 Then
            'Script para ver datos de Contrato
            SQLstring = <sql>
                             SELECT TOP (1) a.ItemCode,                                    
                                a.AttriTxt2,                     
                                b.U_TipoServicio,
                                d.CardName,
                                ISNULL(d.Phone1,'')+' '+ISNULL(d.Phone2,'') As TelephoneNo,
                                d.CardCode,
                                e.Name,
                                e.U_Identificacion,
                                e.E_MailL,
                                b.U_KilometrajeCon,
                                b.U_TipoServicio
                            FROM ITM13 a 
                            INNER Join OITM b ON b.ItemCode = a.ItemCode 
                            INNER JOIN [@MCONTRATO] c ON c.U_NumContrato = b.U_Contrato 
                            INNER JOIN [OCRD] d ON d.U_CodigoMilenia = c.U_CodCliente 
                            INNER JOIN [OCPR] e ON e.CardCode = d.CardCode 
                            WHERE a.AttriTxt1 = '<%= _Placa %>'
                            AND c.U_NumContrato = '<%= _Contrato %>'
                            AND c.U_CodEstado = 'F'
                   </sql>.Value

            Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet.RecordCount > 0 Then
                rSet.MoveFirst()
                While Not rSet.EoF

                    'Campos Adicionales
                    NomProvee = rSet.Fields.Item(3).Value     '//Nombre Proveedor
                    CodigoClie = rSet.Fields.Item(5).Value    '//Codigo del Cliente
                    Direccion = rSet.Fields.Item(4).Value      '//Direccion
                    ActivoFijo = rSet.Fields.Item(0).Value    '//Codigo Articulo
                    SerieUnida = rSet.Fields.Item(1).Value    '//Serie Unidad

                    'Cuenta contable 
                    TipoCuenta = rSet.Fields.Item(10).Value

                    rSet.MoveNext()
                    Exit While
                End While

            Else
                DatosGenereral.SiGrabo = 27
                DatosGenereral.nMensaje = "[027 Call Center] The Contract or Plancese don't Exist "
            End If
        End If


        'Obtiene el No. de Cuenta Contable
        If DatosGenereral.SiGrabo = 0 Then

            'Busco el la Moneda
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "DOCCURRENCY" Then
                    Moneda = Match.Groups.Item(2).Value

                    If Moneda = "GTQ" Then
                        Moneda = "QTZ"
                    End If

                    Exit For
                End If

            Next Match

            If Moneda = "USD" Then
                'Script para ver datos de Contrato
                SQLstring = <sql>
                            Select b.AcctCode, 'USD' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostosD,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
            Else
                'Script para ver datos de Contrato
                SQLstring = <sql>
                            Select b.AcctCode, 'QTZ' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostos,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
            End If

            Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet2.RecordCount > 0 Then
                rSet2.MoveFirst()
                While Not rSet2.EoF

                    CuenConta = rSet2.Fields.Item(0).Value

                    rSet2.MoveNext()
                End While
            End If
        End If


        'Inicia la Recoleccion del Array con todos los campos para Aderirlos en SAP
        If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then
            Try

                'Agrega la Cabecera
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches
                    'Match.Groups.Item(1).Value = Campo 
                    'Match.Groups.Item(2).Value = 'Valor

                    If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                        vOrder.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                    Else
                        CallByName(vOrder, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                    End If

                Next Match

                If _Action = "Add" Then

                    vOrder.Lines.ItemCode = "9600"
                    vOrder.Lines.Quantity = "1"
                    vOrder.Lines.AccountCode = CuenConta
                    vOrder.Lines.Price = 0.00

                Else

                    vOrder.Lines.SetCurrentLine("0")
                    vOrder.Lines.Delete()

                    vOrder.Lines.ItemCode = "9600"
                    vOrder.Lines.Quantity = "1"
                    vOrder.Lines.AccountCode = CuenConta
                    vOrder.Lines.Price = 0.00



                End If

                paso = True

            Catch ex As Exception

                DatosGenereral.SiGrabo = 28
                DatosGenereral.nMensaje = "[028 Call Center] Error SAP: " + ex.Message

                paso = False

            End Try

            'Actualiza datos
            If paso Then
                If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then

                    vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"
                    vOrder.UserFields.Fields.Item("U_EstadoPO").Value = "0002"  '//New

                    'Atividades Estatus
                    Select Case EstadoSolicita
                        Case = "0007"       '//Requested
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0007"   '//Initial [@SMR_ESTADO_ACT]
                        Case = "0001"       '//Approved
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0002"   '//Approved
                        Case = "0006"       '//Rejected
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0003"   '//Rejected
                        Case = "0009"       '//Canceled
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0004"   '//Deleted
                    End Select


                    If _Action = "Add" Then

                        'Campos Adicionales
                        vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"
                        vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                        vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                        vOrder.UserFields.Fields.Item("U_Estado").Value = "FORMALIZADO"
                        vOrder.UserFields.Fields.Item("U_CreatedBy").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_Placa").Value = _Placa
                        vOrder.UserFields.Fields.Item("U_NContrato").Value = _Contrato
                        vOrder.UserFields.Fields.Item("U_NumCredito").Value = _Contrato
                        vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = 0
                        vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita  '//New
                        vOrder.UserFields.Fields.Item("U_RequestComeFrom").Value = _RequestComeFrom
                        vOrder.UserFields.Fields.Item("U_OrdenServicio").Value = _DealerReference

                        vOrder.UserFields.Fields.Item("U_NomCredito").Value = NomProvee     '//Nombre Proveedor
                        vOrder.UserFields.Fields.Item("U_Cliente").Value = CodigoClie       '//Codigo del Cliente
                        vOrder.UserFields.Fields.Item("U_Direccion").Value = Direccion      '//Direccion
                        vOrder.UserFields.Fields.Item("U_CodUnidad").Value = ActivoFijo     '//Codigo Articulo
                        vOrder.UserFields.Fields.Item("U_SerieUnidad").Value = SerieUnida   '//Serie Unidad
                        vOrder.DocCurrency = Moneda

                        vOrder.Series = 130
                        vOrder.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders

                        chk = vOrder.Add()
                    Else
                        'Campos Adicionales
                        vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita
                        vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = Convert.ToInt32(vOrder.UserFields.Fields.Item("U_Actualizaciones").Value) + 1
                        vOrder.DocCurrency = Moneda

                        chk = vOrder.Update()
                    End If

                    If (chk = 0) Then

                        If _Action = "Add" Then
                            _NoDocEntry = SAPConnector.company.GetNewObjectKey
                        End If

                        DatosGenereral.nMensaje = ""
                        DatosGenereral.SiGrabo = 1
                        _SuccessSaved = True

                        'Llamada comentarios--------------------------------------------------------------------------------

                        Dim myarrayCommen(0) As String
                        Dim fieldsComment As String = ""
                        Dim sigrabo As Boolean = False
                        Dim Linea As Integer = 0

                        SQLstring = "Select TOP (1) U_Linea From [@BITAOPOR] WHERE U_DocEntry = '" & _NoDocEntry & "' ORDER BY U_Linea DESC "

                        Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet2.RecordCount > 0 Then
                            Linea = rSet2.Fields.Item(0).Value + 1
                        End If

                        fieldsComment = " Code= '" & "x" & "'"
                        fieldsComment = fieldsComment & " Name='" & "S" & "'"
                        fieldsComment = fieldsComment & " U_Fecha= '" & Now & "'"
                        fieldsComment = fieldsComment & " U_HechoPor= '" & "Sha" & "'"
                        fieldsComment = fieldsComment & " U_DocEntry='" & _NoDocEntry & "'"           '//RequesId
                        fieldsComment = fieldsComment & " U_Linea= '" & Linea & "'"                     '//Control Interno de Linea
                        fieldsComment = fieldsComment & " U_Comentario= '" & vOrder.Comments & "'"    '//Comentario

                        myarrayCommen(0) = fieldsComment

                        sigrabo = CommentsRq(myarrayCommen, _NoDocEntry, Linea, "A")

                        If sigrabo = False Then
                            DatosGenereral.nMensaje = "No Grabo Detalle"
                        End If
                        '------------------------------------------------------------------------------------------------------

                    Else

                        DatosGenereral.SiGrabo = 29
                        DatosGenereral.nMensaje = "[029 Call Center] Error SAP: " + SAPConnector.company.GetLastErrorDescription
                        _SuccessSaved = False
                    End If
                    _Exception = False
                Else
                    _Exception = True
                    _SuccessSaved = False
                End If
            Else
                _SuccessSaved = False
            End If
        Else
            _SuccessSaved = False
        End If
        '-- Termina la grabación exitosa de los campos a SAP

        'If _SuccessSaved Then
        '    _Exception = True
        '    _WarningMsg = "Automatic Auto Approve Not allowed"
        'End If

        'Graba el Mensaje para Control Interno
        If vOrder.GetByKey(_NoDocEntry) Then

            If DatosGenereral.nMensaje.Trim = "" Then
                vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = DatosGenereral.nMensaje
            Else
                vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = _WarningMsg
            End If
            vOrder.Update()
        End If

        DatosGenereral.nMensaje = DatosGenereral.nMensaje & "RequesId: " & _NoDocEntry & " Exception: " & _Exception & " Success: " & _SuccessSaved & " Warning: " & _WarningMsg
        Return DatosGenereral.nMensaje

    End Function
#End Region

#Region "SMR Req"
    'POST de SMR Request Addrequest
    'Programador: Saul
    'Fecha: 02/10/2019
    'Solicitud de Daños
    Function SMRRequestAddrequest(ByVal content() As String, ByVal contentLine() As String, _Contrato As String, _Placa As String, _DealerReference As String, _CreatedBy As String, _MaxQuantity As Integer, _MaxAmount As Double, _MaxMonths As Integer, _MaxOutstanding As Integer, _PerMeleage As Integer, ByRef _Exception As Boolean, ByRef _NoDocEntry As Integer, _AApprove As String, _RequestComeFrom As String, _DaysInBetween As Integer, ByRef _SuccessSaved As Boolean, ByRef _WarningMsg As String) As String
        'Function SMRRequestAddrequest(ByVal content() As String, ByVal contentLine() As String, _Contrato As String, _Placa As String, _DealerReference As String, _CreatedBy As String, _MaxQuantity As Integer, _MaxAmount As Double, _MaxMonths As Integer, _Licencia As String, _MaxOutstanding As Integer, _PerMeleage As Integer, ByRef _Exception As Boolean, ByRef _NoDocEntry As Integer, _AApprove As String, _RequestComeFrom As String) As String
        '"Call Center"
        '"SMR"
        '"Damage"
        '"Replacement"

        DatosGenereral.SiGrabo = 0
        DatosGenereral.nMensaje = ""
        Dim _Action As String = "Add"
        Dim matches2 As MatchCollection
        Dim TipoOrden As String = ""

        'Busco el Nombre
        matches2 = Regex.Matches(content(0), fieldValuePattern)
        For Each Match As Match In matches2

            If Match.Groups.Item(1).Value.ToUpper = "U_TYPEOFREQUEST" Then
                TipoOrden = Match.Groups.Item(2).Value

                DatosGenereral.SiGrabo = 0
                DatosGenereral.nMensaje = ""
                Exit For
            End If

        Next Match

        If _RequestComeFrom = "CALL CENTER" Then

            DatosGenereral.nMensaje = CallCenterAddrequest(content, contentLine, _Contrato, _Placa, _DealerReference, _CreatedBy, _Exception, _NoDocEntry, _RequestComeFrom, _Action, _SuccessSaved, _WarningMsg)
            Return DatosGenereral.nMensaje

        Else


            Select Case TipoOrden

                Case "P02"  'Damage

                    DatosGenereral.nMensaje = RPLAddrequest(content, contentLine, _Contrato, _Placa, _MaxQuantity, _MaxAmount, _MaxOutstanding, _DealerReference, _CreatedBy, _Exception, _NoDocEntry, _AApprove, _RequestComeFrom, _Action, _SuccessSaved, _WarningMsg)
                    Return DatosGenereral.nMensaje

                Case "P03"   'Replacement"

                    DatosGenereral.nMensaje = RRAddrequest(content, contentLine, _Contrato, _Placa, _MaxQuantity, _MaxAmount, _MaxOutstanding, _DealerReference, _CreatedBy, _Exception, _NoDocEntry, _AApprove, _RequestComeFrom, _Action, _SuccessSaved, _WarningMsg)

                    Return DatosGenereral.nMensaje

                Case "P01"   'SMR

                    If Not connected Then connect()

                    Dim matches As MatchCollection

                    Dim nErr As Long = 0
                    Dim errMsg As String = ""
                    Dim chk As Integer = 0

                    'Create the BusinessPartners object
                    Dim vOrder As SAPbobsCOM.Documents

                    DatosGenereral.SiGrabo = 0
                    Dim paso As Boolean = False
                    Dim lngKey As String = ""

                    Dim nTotalLinea As Double = 0.00
                    Dim nPrecio As Double = 0.00
                    Dim nCanti As Integer = 0.00
                    Dim SQLstring As String = ""
                    Dim Proveedor As String = ""
                    Dim CanPo As Integer = 0
                    Dim CanOutPo As Integer = 0
                    Dim TotPo As Double = 0.00
                    Dim Meses As Integer = 0
                    Dim MileContrato As Integer = 0
                    Dim MileConsumidas As Integer = 0
                    Dim FechaPO As String = ""
                    Dim CuenConta As String = ""
                    Dim TipoCuenta As String = ""
                    Dim EstadoSolicita As String = ""
                    Dim Moneda As String = ""
                    Dim NomProvee As String = ""
                    Dim CodigoClie As String = ""
                    Dim Direccion As String = ""
                    Dim ActivoFijo As String = ""
                    Dim SerieUnida As String = ""
                    Dim CodItem As String = ""
                    _NoDocEntry = 0

                    'Crea el Objeto
                    vOrder = getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)


                    'Busco el Nombre
                    If DatosGenereral.SiGrabo = 0 Then
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "U_ESTADOPO" Then
                                EstadoSolicita = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 27
                                DatosGenereral.nMensaje = "[P01-037 SMR] The U_EstadoPO field is not included"
                            End If

                        Next Match

                    End If


                    If DatosGenereral.SiGrabo = 0 Then

                        If Not vOrder.GetByKey(_NoDocEntry) And _Action = "Update" Then
                            DatosGenereral.SiGrabo = 36
                            DatosGenereral.nMensaje = "[P01-036 SMR] The Purchase Order Does Not Exist "
                        End If

                    End If

                    'Verifica si la cantidad de Ordenes de Compra llegaron a su limpite por Proveedor
                    'Condicion No. 1-------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el Nombre
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "CARDCODE" Then
                                Proveedor = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 25
                                DatosGenereral.nMensaje = "[P01-025 SMR] The CardCode field is not included.."
                            End If

                        Next Match

                        If DatosGenereral.SiGrabo = 0 Then
                            'Script para ver datos de Contrato
                            SQLstring = <sql>
                                  Select a.CardCode,  
                                    SUM((CASE when a.DocCur = 'QTZ' then DocTotal  else DocTotalFC END) *  a.DocRate) As TotPO,
                                    COUNT(a.CardCode) As CantPO
                                  From [OPOR] a
                                     inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                                     inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                                     where b.ObjectCode = 22 AND b.Series IN(130,12)
                                     and est.name != 'Cancelled'
                                     and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                                     and a.CANCELED = 'N'
                                     and a.DocStatus = 'O'
                                     AND CardCode = '<%= Proveedor %>'
                                     GROUP BY a.CardCode
                             </sql>.Value

                            Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                            If rSet.RecordCount > 0 Then
                                rSet.MoveFirst()
                                While Not rSet.EoF

                                    CanPo = rSet.Fields.Item(2).Value
                                    TotPo = rSet.Fields.Item(1).Value

                                    rSet.MoveNext()
                                End While
                            Else
                                DatosGenereral.SiGrabo = 26
                                DatosGenereral.nMensaje = "[P01-026 SMR] The Supplier don't Exist "
                            End If
                        End If
                    End If

                    If DatosGenereral.SiGrabo = 0 Then
                        If (CanPo + 1) >= _MaxQuantity And _MaxQuantity > 0 Then
                            DatosGenereral.SiGrabo = 37
                            _WarningMsg = "[P01-037 SMR] Supplier exceeds maximum requests quantity of SMR " & " contact Leasing company"
                        End If
                    End If
                    'Fin de Condicion No. 1 ---------------------------------------------------------------


                    'Verifica si el Monto de Ordenes de Compra llegaron a su limpite por Proveedor
                    'Condicion No. 2-------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then
                        Dim TotActPo As Double = 0.00

                        'Busco el Nombre
                        matches = Regex.Matches(contentLine(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "PRICE" Then
                                TotActPo = Math.Round(Match.Groups.Item(2).Value * 1, 2) + TotActPo

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 27
                                DatosGenereral.nMensaje = "[P01-027 SMR] The Price field is not included.."
                            End If

                        Next Match

                        If DatosGenereral.SiGrabo = 0 Then

                            If TotActPo = 0 Then
                                DatosGenereral.SiGrabo = 28
                                DatosGenereral.nMensaje = "[P01-028 SMR] The requests need one Price "
                            Else

                                If (TotPo + TotActPo) >= _MaxAmount And _MaxAmount > 0 Then
                                    DatosGenereral.SiGrabo = 3
                                    _WarningMsg = "[P01-003 SMR] Supplier exceeds maximum total amount"
                                End If

                            End If
                        End If
                    End If
                    'Fin de Condicion No. 2 ---------------------------------------------------------------


                    'Verifica si el Max Asiganaciones Sobresalientes de Ordenes de Compra llegaron 
                    'a su limpite por Proveedor y por SMR
                    'Condicion No. 3-------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then
                        'Script para ver datos de Contrato
                        SQLstring = <sql>
                     Select COUNT(a.DocEntry) As CantOutPO
                     From OPOR a
                         inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                         inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                         where b.ObjectCode = 22 AND b.Series IN(130,12)
                         and est.name != 'Cancelled'
                         and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                         and a.CANCELED = 'N'
                         AND CardCode = '<%= Proveedor %>'
                         and a.DocStatus = 'O'
                         AND a.U_TypeofRequest = 'P01'  --SMR
                        GROUP BY a.CardCode
                 </sql>.Value

                        Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet2.RecordCount > 0 Then
                            rSet2.MoveFirst()
                            While Not rSet2.EoF

                                CanOutPo = rSet2.Fields.Item(0).Value

                                rSet2.MoveNext()
                            End While
                        Else
                            DatosGenereral.SiGrabo = 29
                            DatosGenereral.nMensaje = "[P01-029 SMR] The Supplier don't Exist "
                        End If

                        If DatosGenereral.SiGrabo = 0 Then
                            If CanOutPo + 1 >= _MaxOutstanding And _MaxOutstanding > 0 Then
                                DatosGenereral.SiGrabo = 4
                                _WarningMsg = "[P01-004 SMR] Supplier exceeds maximum requests Outstanding Assignment " & "contact Leasing company"
                            End If
                        End If
                    End If
                    'Fin de Condicion No. 3 ---------------------------------------------------------------


                    'Verifica si el Porcentaje de Millas ya se cumplio y/o esta llegando
                    'a su limpite por No. de Contrato (Placa)
                    'Condicion No. 4-------------------------------------------------------------------

                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el Nombre
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "U_KILOMETRAJE" Then

                                MileConsumidas = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 38
                                DatosGenereral.nMensaje = "[P01-038 SMR] The U_Kilometraje field is not included.."
                            End If

                        Next Match
                    End If

                    If DatosGenereral.SiGrabo = 0 Then

                        'Script para ver datos de Contrato
                        SQLstring = <sql>
                                       SELECT TOP (1) a.ItemCode,                      
                                            a.AttriTxt2,                     
                                            b.U_TipoServicio,
                                            d.CardName,
                                            ISNULL(d.Phone1,'')+' '+ISNULL(d.Phone2,'') As TelephoneNo,
                                            d.CardCode,
                                            e.Name,
                                            e.U_Identificacion,
                                            e.E_MailL,
                                            b.U_KilometrajeCon,
                                            b.U_TipoServicio
                                       FROM ITM13 a 
                                       INNER Join OITM b ON b.ItemCode = a.ItemCode 
                                       INNER JOIN [@MCONTRATO] c ON c.U_NumContrato = b.U_Contrato 
                                       INNER JOIN [OCRD] d ON d.U_CodigoMilenia = c.U_CodCliente 
                                       INNER JOIN [OCPR] e ON e.CardCode = d.CardCode 
                                       WHERE a.AttriTxt1 = '<%= _Placa %>'
                                       AND c.U_NumContrato = '<%= _Contrato %>'
                                       AND c.U_CodEstado = 'F'
                                </sql>.Value

                        Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet.RecordCount > 0 Then
                            rSet.MoveFirst()
                            While Not rSet.EoF

                                ''Campos Adicionales
                                NomProvee = rSet.Fields.Item(3).Value     '//Nombre Proveedor
                                CodigoClie = rSet.Fields.Item(5).Value    '//Codigo del Cliente
                                Direccion = rSet.Fields.Item(4).Value      '//Direccion
                                ActivoFijo = rSet.Fields.Item(0).Value    '//Codigo Articulo
                                SerieUnida = rSet.Fields.Item(1).Value    '//Serie Unidad

                                MileContrato = Convert.ToInt32(rSet.Fields.Item(9).Value)  '// Kilometraje Contratado

                                'Cuenta contable 
                                TipoCuenta = rSet.Fields.Item(10).Value

                                rSet.MoveNext()
                                Exit While
                            End While

                        Else
                            DatosGenereral.SiGrabo = 31
                            DatosGenereral.nMensaje = "[P01-031 SMR] The Contract or Plancese don't Exist "
                        End If
                    End If

                    'Verifica el porcentaje de Kilometraje    
                    If DatosGenereral.SiGrabo = 0 Then
                        If Convert.ToInt32(100 - (MileConsumidas / MileContrato) * 100) <= _PerMeleage And _PerMeleage > 0 Then
                            DatosGenereral.SiGrabo = 5
                            _WarningMsg = "[P01-005 SMR] Contract exceeds maximum mileage"
                        End If
                    End If

                    'Obtiene el No. de Cuenta Contable
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el la Moneda
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "DOCCURRENCY" Then
                                Moneda = Match.Groups.Item(2).Value

                                If Moneda = "GTQ" Then
                                    Moneda = "QTZ"
                                End If

                                Exit For
                            End If

                        Next Match

                        If Moneda = "USD" Then
                            'Script para ver datos de Contrato
                            SQLstring = <sql>
                            Select b.AcctCode, 'USD' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostosD,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
                        Else
                            'Script para ver datos de Contrato
                            SQLstring = <sql>
                            Select b.AcctCode, 'QTZ' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostos,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
                        End If

                        Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet2.RecordCount > 0 Then
                            rSet2.MoveFirst()
                            While Not rSet2.EoF

                                CuenConta = rSet2.Fields.Item(0).Value

                                rSet2.MoveNext()
                            End While
                        End If
                    End If
                    'Fin de Condicion No. 4 ----------------------------------------------------------------------------


                    'Verifica si la cantidad de Meses ya se cumplio y/o esta llegando
                    'a su limpite por No. de Contrato (Ultima Cuota)
                    'Condicion No. 5-------------------------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el Nombre
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "DOCDATE" Then
                                FechaPO = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 32
                                DatosGenereral.nMensaje = "[P01-032 SMR] The Docdate field is not included.."
                            End If

                        Next Match

                    End If

                    If DatosGenereral.SiGrabo = 0 Then
                        If _MaxMonths > 0 Then
                            SQLstring = "SELECT TOP (1) * FROM ( "
                            SQLstring = SQLstring & " Select TOP (" & _MaxMonths + 1 & ") U_FechaPago  As MesPago "
                            SQLstring = SQLstring & " From  admon.dbo.[@MCONTRATO] c "
                            SQLstring = SQLstring & "                     inner join admon.dbo.[@MCUOTA_CONTRATO] a ON a.U_NumContrato = c.U_NumContrato "
                            SQLstring = SQLstring & "                    where c.U_NumContrato = '" & _Contrato & "' "
                            SQLstring = SQLstring & "                     And c.U_CodEstado = 'F' "
                            SQLstring = SQLstring & "                     order by U_FechaPago desc) X "
                            SQLstring = SQLstring & "             ORDER BY 1 "

                            Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                            If rSet.RecordCount > 0 Then
                                rSet.MoveFirst()
                                While Not rSet.EoF

                                    If FechaPO >= rSet.Fields.Item(0).Value Then
                                        DatosGenereral.SiGrabo = 6
                                        _WarningMsg = "[P01-006 SMR] The request date is close to contract due date"
                                    End If

                                    Exit While
                                End While
                            Else
                                DatosGenereral.SiGrabo = 33
                                DatosGenereral.nMensaje = "[P01-033 SMR] The Contract need one Month "
                            End If
                        End If
                    End If
                    'Fin de Condicion No. 5 ----------------------------------------------------------------------------


                    'Verifica si Los días entre servicios DayInBetween
                    'a su limpite por Tipo de Servicio "M1"
                    'Condicion No. 6-------------------------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el Nombre
                        matches = Regex.Matches(contentLine(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "ITEMCODE" Then
                                CodItem = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 39
                                DatosGenereral.nMensaje = "[P01-039 SMR] The ItemCode field is not included.."
                            End If

                        Next Match

                    End If


                    If DatosGenereral.SiGrabo = 0 Then

                        SQLstring = <sql>
                                         SELECT TOP (1) 
					                       a.DocEntry,
					                       a.DocNum,
					                       a.DocDate
                                         FROM OPOR a
                                             INNER JOIN POR1 ad ON ad.DocEntry = a.DocEntry
                                             inner join NNM1 b ON b.ObjectCode = a.objType AND b.Series = a.Series
                                             inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                                             INNER JOIN (SELECT	a.U_TipoReparacion,
												                    a.U_FamliaRep,
												                    a.ItemCode, 
												                    a.ItemName
										                    FROM OITM a
										                    INNER JOIN [@TIPOREPARACION] b ON a.U_TipoReparacion  = b.Code
										                    WHERE b.Code = 'M1'  --Tipo Servicio Preventivo
									                    ) ot ON ot.ItemCode = ad.ItemCode
                                         WHERE b.ObjectCode = 22 AND b.Series IN(130,12)
                                             and est.name != 'Cancelled'
                                             and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                                             and a.CANCELED = 'N'
                                             and a.DocStatus = 'O'
                                             AND CardCode = '<%= Proveedor %>'
                                             AND ot.ItemCode = '<%= CodItem %>'
                                        ORDER BY a.DocDate DESC
                                </sql>.Value

                        Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet.RecordCount > 0 Then

                            Dim days As Long = DateDiff(DateInterval.Day, Convert.ToDateTime(rSet.Fields.Item(2).Value), Now)
                            Dim remaingdays As Long = _DaysInBetween - days

                            If days < _DaysInBetween And _DaysInBetween > 0 Then

                                DatosGenereral.SiGrabo = 7
                                _WarningMsg = "[P01-007 SMR] Your next service will be available in " & remaingdays & " days."

                            End If

                        End If
                    End If
                    'Fin de Condicion No. 6 ----------------------------------------------------------------------------


                    If _Action = "Add" Then
                        If _AApprove.ToUpper = "YES" Then

                            If DatosGenereral.SiGrabo = 0 Then  '//Grabada sin Errores
                                EstadoSolicita = "0001"     '//Approved
                                _Exception = False
                            Else
                                If DatosGenereral.SiGrabo > 0 And DatosGenereral.SiGrabo < 25 Then
                                    EstadoSolicita = "0007"  '//Requested
                                End If
                                _Exception = True
                            End If

                        Else  '//Aprove = "NO"

                            EstadoSolicita = "0007"  '//Requested
                            _Exception = False

                        End If
                    End If


                    'Inicia la Recoleccion del Array con todos los campos para Aderirlos en SAP
                    If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then
                        Try

                            'Agrega la Cabecera
                            matches = Regex.Matches(content(0), fieldValuePattern)
                            For Each Match As Match In matches
                                'Match.Groups.Item(1).Value = Campo 
                                'Match.Groups.Item(2).Value = 'Valor

                                If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                                    vOrder.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                                Else
                                    CallByName(vOrder, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                                End If

                            Next Match

                            If _Action = "Add" Then
                                'Agrega el Deatalle de Lines
                                matches = Regex.Matches(contentLine(0), fieldValuePattern)
                                For Each Match As Match In matches
                                    'Match.Groups.Item(1).Value = Campo 
                                    'Match.Groups.Item(2).Value = 'Valor

                                    If Left(Match.Groups.Item(1).Value, 2) = "U_" Then

                                        If Match.Groups.Item(1).Value.Trim <> "U_LineNum" Then   '//Control de Lines
                                            vOrder.Lines.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                                        Else
                                            If Match.Groups.Item(2).Value.Trim <> "" Then

                                                'No se permite ingresar mas de un Articulo
                                                If Convert.ToInt32(Match.Groups.Item(2).Value.Trim) > 0 Then
                                                    DatosGenereral.SiGrabo = 41
                                                    DatosGenereral.nMensaje = "[P01-041 SMR] The Document only accepts one Item"
                                                    Exit For
                                                End If

                                                'vOrder.Lines.Add()
                                            End If
                                        End If

                                    Else
                                        'Para Poner el Numero de Cuenta
                                        If Match.Groups.Item(1).Value.Trim = "AccountCode" Then
                                            CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], CuenConta)
                                        Else
                                            CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                                        End If
                                    End If
                                Next Match
                            End If

                            paso = True

                        Catch ex As Exception

                            DatosGenereral.SiGrabo = 34
                            DatosGenereral.nMensaje = "[P01-034 SMR] Error SAP: " + ex.Message

                            paso = False

                        End Try

                        'Actualiza datos
                        If paso Then
                            If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then

                                'Campos Adicionales
                                vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"
                                vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                                vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                                vOrder.UserFields.Fields.Item("U_Estado").Value = "FORMALIZADO"
                                vOrder.UserFields.Fields.Item("U_CreatedBy").Value = _CreatedBy
                                vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                                vOrder.UserFields.Fields.Item("U_Placa").Value = _Placa
                                vOrder.UserFields.Fields.Item("U_NContrato").Value = _Contrato
                                vOrder.UserFields.Fields.Item("U_NumCredito").Value = _Contrato
                                vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = 0        '//Iniando
                                vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita
                                vOrder.UserFields.Fields.Item("U_RequestComeFrom").Value = _RequestComeFrom
                                vOrder.UserFields.Fields.Item("U_OrdenServicio").Value = _DealerReference

                                vOrder.UserFields.Fields.Item("U_NomCredito").Value = NomProvee     '//Nombre Proveedor
                                vOrder.UserFields.Fields.Item("U_Cliente").Value = CodigoClie       '//Codigo del Cliente
                                vOrder.UserFields.Fields.Item("U_Direccion").Value = Direccion      '//Direccion
                                vOrder.UserFields.Fields.Item("U_CodUnidad").Value = ActivoFijo     '//Codigo Articulo
                                vOrder.UserFields.Fields.Item("U_SerieUnidad").Value = SerieUnida   '//Serie Unidad
                                vOrder.DocCurrency = Moneda


                                'Atividades Estatus
                                Select Case EstadoSolicita
                                    Case = "0007"       '//Requested
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0001"   '//Initial [@SMR_ESTADO_ACT]
                                    Case = "0001"       '//Approved
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0002"   '//Approved
                                    Case = "0006"       '//Rejected
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0003"   '//Approved
                                    Case = "0009"       '//Canceled
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0004"   '//Approved
                                End Select


                                If _Action = "Add" Then

                                    vOrder.Series = 130
                                    vOrder.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders

                                    chk = vOrder.Add()
                                End If

                                If (chk = 0) Then

                                    If _Action = "Add" Then
                                        _NoDocEntry = SAPConnector.company.GetNewObjectKey
                                        'DatosGenereral.nMensaje = "[P01-001 SMR] The Purcharse Order was successfully added, the No it's " & _NoDocEntry
                                        DatosGenereral.nMensaje = ""
                                        _SuccessSaved = True
                                    End If

                                    DatosGenereral.SiGrabo = 1
                                    '_Exception = False


                                    'Llamada comentarios--------------------------------------------------------------------------------
                                    Dim myarrayCommen(0) As String
                                    Dim fieldsComment As String = ""
                                    Dim sigrabo As Boolean = False
                                    Dim Linea As Integer = 0

                                    SQLstring = "Select TOP (1) U_Linea From [@BITAOPOR] WHERE U_DocEntry = '" & _NoDocEntry & "' ORDER BY U_Linea DESC "

                                    Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                                    If rSet2.RecordCount > 0 Then
                                        Linea = rSet2.Fields.Item(0).Value + 1
                                    End If

                                    fieldsComment = " Code= '" & "x" & "'"
                                    fieldsComment = fieldsComment & " Name='" & "S" & "'"
                                    fieldsComment = fieldsComment & " U_Fecha= '" & Now & "'"
                                    fieldsComment = fieldsComment & " U_HechoPor= '" & "Sha" & "'"
                                    fieldsComment = fieldsComment & " U_DocEntry='" & _NoDocEntry & "'"           '//RequesId
                                    fieldsComment = fieldsComment & " U_Linea= '" & Linea & "'"                     '//Control Interno de Linea
                                    fieldsComment = fieldsComment & " U_Comentario= '" & vOrder.Comments & "'"    '//Comentario

                                    myarrayCommen(0) = fieldsComment

                                    sigrabo = CommentsRq(myarrayCommen, _NoDocEntry, Linea, "A")

                                    If sigrabo = False Then
                                        DatosGenereral.nMensaje = "No Grabo Detalle"
                                    End If
                                    '------------------------------------------------------------------------------------------------------

                                Else

                                    DatosGenereral.SiGrabo = 35
                                    DatosGenereral.nMensaje = "[P01-035 SMR] Error SAP: " + SAPConnector.company.GetLastErrorDescription

                                    '_Exception = True
                                    _SuccessSaved = False
                                End If

                            Else
                                _Exception = True
                                _SuccessSaved = False
                            End If
                        Else
                            _SuccessSaved = False
                        End If
                    Else
                        _Exception = True
                        _SuccessSaved = False
                    End If

                    If _AApprove.ToUpper = "YES" Then

                    Else
                        If _SuccessSaved Then
                            _Exception = True
                            _WarningMsg = "Automatic Auto Approve Not allowed"
                            'Else
                            '   _WarningMsg = ""
                        End If
                    End If

                    'Graba el Mensaje para Control Interno
                    If vOrder.GetByKey(_NoDocEntry) Then

                        If DatosGenereral.nMensaje.Trim = "" Then
                            vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = DatosGenereral.nMensaje
                        Else
                            vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = _WarningMsg
                        End If
                        vOrder.Update()
                    End If

                    DatosGenereral.nMensaje = DatosGenereral.nMensaje & "RequesId: " & _NoDocEntry & " Exception: " & _Exception & " Success: " & _SuccessSaved & " Warning: " & _WarningMsg
                    Return DatosGenereral.nMensaje

            End Select
        End If  '//Selecciona RequestComeFrome

    End Function


    'POST de SMR Request Addrequest
    'Programador: Saul
    'Fecha: 02/10/2019
    'Solicitud de Daños
    Function SMRRequestEditrequest(ByVal content() As String, ByVal contentLine() As String, _Contrato As String, _Placa As String, _DealerReference As String, _CreatedBy As String, _MaxQuantity As Integer, _MaxAmount As Double, _MaxMonths As Integer, _MaxOutstanding As Integer, _PerMeleage As Integer, ByRef _Exception As Boolean, ByRef _NoDocEntry As Integer, _AApprove As String, _RequestComeFrom As String, _DaysInBetween As Integer, ByRef _SuccessSaved As Boolean, ByRef _WarningMsg As String) As String

        '"Call Center"
        '"SMR"
        '"Damage"
        '"Replacement"

        DatosGenereral.SiGrabo = 0
        DatosGenereral.nMensaje = ""
        Dim _Action As String = "Update"
        Dim matches2 As MatchCollection
        Dim TipoOrden As String = ""
        Dim EditCC As String = ""
        Dim VeriItem As Integer = 0

        'Busco el Nombre
        matches2 = Regex.Matches(content(0), fieldValuePattern)
        For Each Match As Match In matches2

            If Match.Groups.Item(1).Value.ToUpper = "U_TYPEOFREQUEST" Then
                TipoOrden = Match.Groups.Item(2).Value

                DatosGenereral.SiGrabo = 0
                DatosGenereral.nMensaje = ""
                Exit For
            End If

        Next Match

        If _RequestComeFrom = "CALL CENTER" Then
            'Verifica si la moficacion afecta Articulos -- (Esto significa que crearon una orden como tal)
            matches2 = Regex.Matches(contentLine(0), fieldValuePattern)
            For Each Match As Match In matches2

                If Match.Groups.Item(1).Value.ToUpper = "ITEMCODE" And Match.Groups.Item(2).Value <> "9600" Then

                    VeriItem = VeriItem + 1

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                End If

            Next Match

        End If

        If VeriItem = 0 Then
            EditCC = "CC"
        Else
            EditCC = "SMR"
        End If

        If _RequestComeFrom = "CALL CENTER" And EditCC = "CC" Then

            DatosGenereral.nMensaje = CallCenterAddrequest(content, contentLine, _Contrato, _Placa, _DealerReference, _CreatedBy, _Exception, _NoDocEntry, _RequestComeFrom, _Action, _SuccessSaved, _WarningMsg)
            Return DatosGenereral.nMensaje

        Else

            Select Case TipoOrden
                Case "P02"

                    DatosGenereral.nMensaje = RPLAddrequest(content, contentLine, _Contrato, _Placa, _MaxQuantity, _MaxAmount, _MaxOutstanding, _DealerReference, _CreatedBy, _Exception, _NoDocEntry, _AApprove, _RequestComeFrom, _Action, _SuccessSaved, _WarningMsg)

                    Return DatosGenereral.nMensaje

                Case "P03"
                    DatosGenereral.nMensaje = RRAddrequest(content, contentLine, _Contrato, _Placa, _MaxQuantity, _MaxAmount, _MaxOutstanding, _DealerReference, _CreatedBy, _Exception, _NoDocEntry, _AApprove, _RequestComeFrom, _Action, _SuccessSaved, _WarningMsg)

                    Return DatosGenereral.nMensaje

                Case "P01"

                    If Not connected Then connect()

                    Dim matches As MatchCollection

                    Dim nErr As Long = 0
                    Dim errMsg As String = ""
                    Dim chk As Integer = 0

                    'Create the BusinessPartners object
                    Dim vOrder As SAPbobsCOM.Documents

                    DatosGenereral.SiGrabo = 0
                    Dim paso As Boolean = False
                    Dim lngKey As String = ""

                    Dim nTotalLinea As Double = 0.00
                    Dim nPrecio As Double = 0.00
                    Dim nCanti As Integer = 0.00
                    Dim SQLstring As String = ""
                    Dim Proveedor As String = ""
                    Dim CanPo As Integer = 0
                    Dim CanOutPo As Integer = 0
                    Dim TotPo As Double = 0.00
                    Dim Meses As Integer = 0
                    Dim MileContrato As Integer = 0
                    Dim MileConsumidas As Integer = 0
                    Dim FechaPO As String = ""
                    Dim CuenConta As String = ""
                    Dim TipoCuenta As String = ""
                    Dim EstadoSolicita As String = ""
                    Dim Moneda As String = ""
                    Dim NomProvee As String = ""
                    Dim CodigoClie As String = ""
                    Dim Direccion As String = ""
                    Dim ActivoFijo As String = ""
                    Dim SerieUnida As String = ""
                    Dim CodItem As String = ""

                    'Crea el Objeto
                    vOrder = getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

                    If DatosGenereral.SiGrabo = 0 Then

                        If Not vOrder.GetByKey(_NoDocEntry) And _Action = "Update" Then
                            DatosGenereral.SiGrabo = 36
                            DatosGenereral.nMensaje = "[P01-036 SMR] The Purchase Order Does Not Exist "
                        End If

                    End If

                    'Busco el Nombre
                    If DatosGenereral.SiGrabo = 0 Then
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "U_ESTADOPO" Then
                                EstadoSolicita = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 27
                                DatosGenereral.nMensaje = "[P01-037 SMR] The U_EstadoPO field is not included"
                            End If

                        Next Match

                    End If

                    'Dato Cancelado
                    If DatosGenereral.SiGrabo = 0 Then

                        If EstadoSolicita = "0009" Then
                            DatosGenereral.SiGrabo = 27
                            DatosGenereral.nMensaje = "[P01-038 SMR] The Status is Canceled, it cann't be modified "
                        End If

                    End If

                    'Verifica si la cantidad de Ordenes de Compra llegaron a su limpite por Proveedor
                    'Condicion No. 1-------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el Nombre
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "CARDCODE" Then
                                Proveedor = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 25
                                DatosGenereral.nMensaje = "[P01-025 SMR] The CardCode field is not included.."
                            End If

                        Next Match

                        If DatosGenereral.SiGrabo = 0 Then
                            'Script para ver datos de Contrato
                            SQLstring = <sql>
                                  Select a.CardCode,  
                                    SUM((CASE when a.DocCur = 'QTZ' then DocTotal  else DocTotalFC END) *  a.DocRate) As TotPO,
                                    COUNT(a.CardCode) As CantPO
                                  From [OPOR] a
                                     inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                                     inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                                     where b.ObjectCode = 22 AND b.Series IN(130,12)
                                     and est.name != 'Cancelled'
                                     and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                                     and a.CANCELED = 'N'
                                     and a.DocStatus = 'O'
                                     AND CardCode = '<%= Proveedor %>'
                                     GROUP BY a.CardCode
                             </sql>.Value

                            Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                            If rSet.RecordCount > 0 Then
                                rSet.MoveFirst()
                                While Not rSet.EoF

                                    CanPo = rSet.Fields.Item(2).Value
                                    TotPo = rSet.Fields.Item(1).Value

                                    rSet.MoveNext()
                                End While
                            Else
                                DatosGenereral.SiGrabo = 26
                                DatosGenereral.nMensaje = "[P01-026 SMR] The Supplier don't Exist "
                            End If
                        End If
                    End If

                    If DatosGenereral.SiGrabo = 0 Then
                        If CanPo >= _MaxQuantity And _MaxQuantity > 0 Then
                            DatosGenereral.SiGrabo = 37
                            _WarningMsg = "[P01-037 SMR] Supplier exceeds maximum requests quantity of SMR " & " contact Leasing company"
                        End If
                    End If
                    'Fin de Condicion No. 1 ---------------------------------------------------------------


                    'Verifica si el Monto de Ordenes de Compra llegaron a su limpite por Proveedor
                    'Condicion No. 2-------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then
                        Dim TotActPo As Double = 0.00

                        'Busco el Nombre
                        matches = Regex.Matches(contentLine(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "PRICE" Then
                                TotActPo = Math.Round(Match.Groups.Item(2).Value * 1, 2) + TotActPo

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 27
                                DatosGenereral.nMensaje = "[P01-027 SMR] The Price field is not included.."
                            End If

                        Next Match

                        If DatosGenereral.SiGrabo = 0 Then

                            If TotActPo = 0 Then
                                DatosGenereral.SiGrabo = 28
                                DatosGenereral.nMensaje = "[P01-028 SMR] The requests need one Price "
                            Else

                                If TotPo >= _MaxAmount And _MaxAmount > 0 Then
                                    DatosGenereral.SiGrabo = 3
                                    _WarningMsg = "[P01-003 SMR] Supplier exceeds maximum total amount"
                                End If

                            End If
                        End If
                    End If
                    'Fin de Condicion No. 2 ---------------------------------------------------------------


                    'Verifica si el Max Asiganaciones Sobresalientes de Ordenes de Compra llegaron 
                    'a su limpite por Proveedor y por SMR
                    'Condicion No. 3-------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then
                        'Script para ver datos de Contrato
                        SQLstring = <sql>
                     Select COUNT(a.DocEntry) As CantOutPO
                     From OPOR a
                         inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                         inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                         where b.ObjectCode = 22 AND b.Series IN(130,12)
                         and est.name != 'Cancelled'
                         and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                         and a.CANCELED = 'N'
                         AND CardCode = '<%= Proveedor %>'
                         and a.DocStatus = 'O'
                         AND a.U_TypeofRequest = 'P01'  --SMR
                        GROUP BY a.CardCode
                 </sql>.Value

                        Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet2.RecordCount > 0 Then
                            rSet2.MoveFirst()
                            While Not rSet2.EoF

                                CanOutPo = rSet2.Fields.Item(0).Value

                                rSet2.MoveNext()
                            End While
                        Else
                            DatosGenereral.SiGrabo = 29
                            DatosGenereral.nMensaje = "[P01-029 SMR] The Supplier don't Exist "
                        End If

                        If DatosGenereral.SiGrabo = 0 Then
                            If CanOutPo >= _MaxOutstanding And _MaxOutstanding > 0 Then
                                DatosGenereral.SiGrabo = 4
                                _WarningMsg = "[P01-004 SMR] Supplier exceeds maximum requests Outstanding Assignment " & "contact Leasing company"
                            End If
                        End If
                    End If
                    'Fin de Condicion No. 3 ---------------------------------------------------------------


                    'Verifica si el Porcentaje de Millas ya se cumplio y/o esta llegando
                    'a su limpite por No. de Contrato (Placa)
                    'Condicion No. 4-------------------------------------------------------------------

                    'If DatosGenereral.SiGrabo = 0 Then

                    '    'Busco el Nombre
                    '    matches = Regex.Matches(content(0), fieldValuePattern)
                    '    For Each Match As Match In matches

                    '        If Match.Groups.Item(1).Value.ToUpper = "U_KILOMETRAJE" Then

                    '            MileConsumidas = Match.Groups.Item(2).Value

                    '            DatosGenereral.SiGrabo = 0
                    '            DatosGenereral.nMensaje = ""
                    '            Exit For
                    '        Else
                    '            DatosGenereral.SiGrabo = 38
                    '            DatosGenereral.nMensaje = "[P01-038 SMR] The U_Kilometraje field is not included.."
                    '        End If

                    '    Next Match
                    'End If

                    If DatosGenereral.SiGrabo = 0 Then

                        'Script para ver datos de Contrato
                        SQLstring = <sql>
                                       SELECT TOP (1) a.ItemCode,                      
                                            a.AttriTxt2,                     
                                            b.U_TipoServicio,
                                            d.CardName,
                                            ISNULL(d.Phone1,'')+' '+ISNULL(d.Phone2,'') As TelephoneNo,
                                            d.CardCode,
                                            e.Name,
                                            e.U_Identificacion,
                                            e.E_MailL,
                                            b.U_KilometrajeCon,
                                            b.U_TipoServicio
                                       FROM ITM13 a 
                                       INNER Join OITM b ON b.ItemCode = a.ItemCode 
                                       INNER JOIN [@MCONTRATO] c ON c.U_NumContrato = b.U_Contrato 
                                       INNER JOIN [OCRD] d ON d.U_CodigoMilenia = c.U_CodCliente 
                                       INNER JOIN [OCPR] e ON e.CardCode = d.CardCode 
                                       WHERE a.AttriTxt1 = '<%= _Placa %>'
                                       AND c.U_NumContrato = '<%= _Contrato %>'
                                       AND c.U_CodEstado = 'F'
                                </sql>.Value

                        Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet.RecordCount > 0 Then
                            rSet.MoveFirst()
                            While Not rSet.EoF

                                ''Campos Adicionales
                                NomProvee = rSet.Fields.Item(3).Value     '//Nombre Proveedor
                                CodigoClie = rSet.Fields.Item(5).Value    '//Codigo del Cliente
                                Direccion = rSet.Fields.Item(4).Value      '//Direccion
                                ActivoFijo = rSet.Fields.Item(0).Value    '//Codigo Articulo
                                SerieUnida = rSet.Fields.Item(1).Value    '//Serie Unidad

                                MileContrato = Convert.ToInt32(rSet.Fields.Item(9).Value)  '// Kilometraje Contratado

                                'Cuenta contable 
                                TipoCuenta = rSet.Fields.Item(10).Value

                                rSet.MoveNext()
                                Exit While
                            End While

                        Else
                            DatosGenereral.SiGrabo = 31
                            DatosGenereral.nMensaje = "[P01-031 SMR] The Contract or Plancese don't Exist "
                        End If
                    End If

                    ''Verifica el porcentaje de Kilometraje    
                    'If DatosGenereral.SiGrabo = 0 Then
                    '    If Convert.ToInt32(100 - (MileConsumidas / MileContrato) * 100) <= _PerMeleage And _PerMeleage > 0 Then
                    '        DatosGenereral.SiGrabo = 5
                    '        _WarningMsg = "[P01-005 SMR] Contract exceeds maximum mileage"
                    '    End If
                    'End If

                    'Obtiene el No. de Cuenta Contable
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el la Moneda
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "DOCCURRENCY" Then
                                Moneda = Match.Groups.Item(2).Value

                                If Moneda = "GTQ" Then
                                    Moneda = "QTZ"
                                End If

                                Exit For
                            End If

                        Next Match

                        If Moneda = "USD" Then
                            'Script para ver datos de Contrato
                            SQLstring = <sql>
                                    Select b.AcctCode, 'USD' As Mon 
                                       From [@LOP] a
                                       INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostosD,'-','')
                                       Where Code = '<%= TipoCuenta %>' 
                                </sql>.Value
                        Else
                            'Script para ver datos de Contrato
                            SQLstring = <sql>
                                    Select b.AcctCode, 'QTZ' As Mon 
                                       From [@LOP] a
                                       INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostos,'-','')
                                       Where Code = '<%= TipoCuenta %>' 
                                </sql>.Value
                        End If

                        Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet2.RecordCount > 0 Then
                            rSet2.MoveFirst()
                            While Not rSet2.EoF

                                CuenConta = rSet2.Fields.Item(0).Value

                                rSet2.MoveNext()
                            End While
                        End If
                    End If
                    'Fin de Condicion No. 4 ----------------------------------------------------------------------------


                    'Verifica si la cantidad de Meses ya se cumplio y/o esta llegando
                    'a su limpite por No. de Contrato (Ultima Cuota)
                    'Condicion No. 5-------------------------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el Nombre
                        matches = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "DOCDATE" Then
                                FechaPO = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 32
                                DatosGenereral.nMensaje = "[P01-032 SMR] The Docdate field is not included.."
                            End If

                        Next Match

                    End If

                    If DatosGenereral.SiGrabo = 0 Then
                        If _MaxMonths > 0 Then
                            SQLstring = "SELECT TOP (1) * FROM ( "
                            SQLstring = SQLstring & " Select TOP (" & _MaxMonths + 1 & ") U_FechaPago  As MesPago "
                            SQLstring = SQLstring & " From  admon.dbo.[@MCONTRATO] c "
                            SQLstring = SQLstring & "                     inner join admon.dbo.[@MCUOTA_CONTRATO] a ON a.U_NumContrato = c.U_NumContrato "
                            SQLstring = SQLstring & "                    where c.U_NumContrato = '" & _Contrato & "' "
                            SQLstring = SQLstring & "                     And c.U_CodEstado = 'F' "
                            SQLstring = SQLstring & "                     order by U_FechaPago desc) X "
                            SQLstring = SQLstring & "             ORDER BY 1 "

                            Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                            If rSet.RecordCount > 0 Then
                                rSet.MoveFirst()
                                While Not rSet.EoF

                                    If FechaPO >= rSet.Fields.Item(0).Value Then
                                        DatosGenereral.SiGrabo = 6
                                        _WarningMsg = "[P01-006 SMR] The request date is close to contract due date"
                                    End If

                                    Exit While
                                End While
                            Else
                                DatosGenereral.SiGrabo = 33
                                DatosGenereral.nMensaje = "[P01-033 SMR] The Contract need one Month "
                            End If
                        End If
                    End If
                    'Fin de Condicion No. 5 ----------------------------------------------------------------------------


                    'Verifica si Los días entre servicios DayInBetween
                    'a su limpite por Tipo de Servicio "M1"
                    'Condicion No. 6-------------------------------------------------------------------------------------
                    If DatosGenereral.SiGrabo = 0 Then

                        'Busco el Nombre
                        matches = Regex.Matches(contentLine(0), fieldValuePattern)
                        For Each Match As Match In matches

                            If Match.Groups.Item(1).Value.ToUpper = "ITEMCODE" Then
                                CodItem = Match.Groups.Item(2).Value

                                DatosGenereral.SiGrabo = 0
                                DatosGenereral.nMensaje = ""
                                Exit For
                            Else
                                DatosGenereral.SiGrabo = 39
                                DatosGenereral.nMensaje = "[P01-039 SMR] The ItemCode field is not included.."
                            End If

                        Next Match

                    End If


                    If DatosGenereral.SiGrabo = 0 Then

                        SQLstring = <sql>
                                         SELECT TOP (1) 
					                       a.DocEntry,
					                       a.DocNum,
					                       a.DocDate
                                         FROM OPOR a
                                             INNER JOIN POR1 ad ON ad.DocEntry = a.DocEntry
                                             inner join NNM1 b ON b.ObjectCode = a.objType AND b.Series = a.Series
                                             inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                                             INNER JOIN (SELECT	a.U_TipoReparacion,
												                    a.U_FamliaRep,
												                    a.ItemCode, 
												                    a.ItemName
										                    FROM OITM a
										                    INNER JOIN [@TIPOREPARACION] b ON a.U_TipoReparacion  = b.Code
										                    WHERE b.Code = 'M1'  --Tipo Servicio Preventivo
									                    ) ot ON ot.ItemCode = ad.ItemCode
                                         WHERE b.ObjectCode = 22 AND b.Series IN(130,12)
                                             and est.name != 'Cancelled'
                                             and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                                             and a.CANCELED = 'N'
                                             and a.DocStatus = 'O'
                                             AND CardCode = '<%= Proveedor %>'
                                             AND ot.ItemCode = '<%= CodItem %>'
                                        ORDER BY a.DocDate DESC
                                </sql>.Value

                        Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet.RecordCount > 0 Then

                            Dim days As Long = DateDiff(DateInterval.Day, Convert.ToDateTime(rSet.Fields.Item(2).Value), Now)

                            If days < _DaysInBetween And _DaysInBetween > 0 Then

                                DatosGenereral.SiGrabo = 7
                                _WarningMsg = "[P01-007 SMR] Your next service date is not yet met"

                            End If

                        End If
                    End If
                    'Fin de Condicion No. 6 ----------------------------------------------------------------------------


                    If _AApprove.ToUpper = "YES" Then

                        If DatosGenereral.SiGrabo = 0 Then  '//Grabada sin Errores
                            'EstadoSolicita = "0001"  '//Approved
                            _Exception = False
                        Else
                            If DatosGenereral.SiGrabo > 0 And DatosGenereral.SiGrabo < 25 Then
                                'EstadoSolicita = "0007"  '//Requested
                            End If
                            _Exception = True
                        End If

                    Else  '//Aprove = "NO"

                        'EstadoSolicita = "0007"  '//Requested
                        _Exception = False

                    End If


                    'Inicia la Recoleccion del Array con todos los campos para Aderirlos en SAP
                    If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then
                        Try

                            'Agrega la Cabecera
                            matches = Regex.Matches(content(0), fieldValuePattern)
                            For Each Match As Match In matches
                                'Match.Groups.Item(1).Value = Campo 
                                'Match.Groups.Item(2).Value = 'Valor

                                If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                                    vOrder.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                                Else
                                    CallByName(vOrder, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                                End If

                            Next Match

                            'If _Action = "Add" Then
                            'Agrega el Deatalle de Lines
                            matches = Regex.Matches(contentLine(0), fieldValuePattern)
                            For Each Match As Match In matches
                                'Match.Groups.Item(1).Value = Campo 
                                'Match.Groups.Item(2).Value = 'Valor

                                If Left(Match.Groups.Item(1).Value, 2) = "U_" Then

                                    If Match.Groups.Item(1).Value.Trim <> "U_LineNum" Then   '//Control de Lines
                                        vOrder.Lines.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                                    Else
                                        If Match.Groups.Item(2).Value.Trim <> "" Then

                                            'No se permite ingresar mas de un Articulo
                                            If Convert.ToInt32(Match.Groups.Item(2).Value.Trim) > 0 Then
                                                DatosGenereral.SiGrabo = 41
                                                DatosGenereral.nMensaje = "[P01-041 SMR] The Document only accepts one Item"
                                                Exit For
                                            End If

                                            'vOrder.Lines.Add()
                                        End If
                                    End If

                                Else
                                    'Para Poner el Numero de Cuenta
                                    If Match.Groups.Item(1).Value.Trim = "AccountCode" Then
                                        CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], CuenConta)
                                    Else
                                        CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                                    End If
                                End If
                            Next Match
                            ' End If

                            paso = True

                        Catch ex As Exception

                            DatosGenereral.SiGrabo = 34
                            DatosGenereral.nMensaje = "[P01-034 SMR] Error SAP: " + ex.Message

                            paso = False

                        End Try

                        'Actualiza datos
                        If paso Then
                            If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then

                                'Campos Adicionales
                                vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"
                                vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                                vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                                vOrder.UserFields.Fields.Item("U_Estado").Value = "FORMALIZADO"
                                vOrder.UserFields.Fields.Item("U_ModifiedBy").Value = _CreatedBy
                                vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                                vOrder.UserFields.Fields.Item("U_Placa").Value = _Placa
                                vOrder.UserFields.Fields.Item("U_NContrato").Value = _Contrato
                                vOrder.UserFields.Fields.Item("U_NumCredito").Value = _Contrato
                                vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita
                                vOrder.UserFields.Fields.Item("U_RequestComeFrom").Value = _RequestComeFrom
                                vOrder.UserFields.Fields.Item("U_OrdenServicio").Value = _DealerReference

                                vOrder.UserFields.Fields.Item("U_NomCredito").Value = NomProvee     '//Nombre Proveedor
                                vOrder.UserFields.Fields.Item("U_Cliente").Value = CodigoClie       '//Codigo del Cliente
                                vOrder.UserFields.Fields.Item("U_Direccion").Value = Direccion      '//Direccion
                                vOrder.UserFields.Fields.Item("U_CodUnidad").Value = ActivoFijo     '//Codigo Articulo
                                vOrder.UserFields.Fields.Item("U_SerieUnidad").Value = SerieUnida   '//Serie Unidad
                                vOrder.DocCurrency = Moneda

                                'Atividades Estatus
                                Select Case EstadoSolicita
                                    Case = "0007"       '//Requested
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0001"   '//Initial [@SMR_ESTADO_ACT]
                                    Case = "0001"       '//Approved
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0002"   '//Approved
                                    Case = "0006"       '//Rejected
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0003"   '//Approved
                                    Case = "0009"       '//Canceled
                                        vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0004"   '//Approved
                                End Select

                                vOrder.Series = 130
                                vOrder.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
                                vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = vOrder.UserFields.Fields.Item("U_Actualizaciones").Value + 1

                                chk = vOrder.Update
                                'End If

                                If (chk = 0) Then

                                    '¿If _Action = "Add" Then
                                    '_NoDocEntry = SAPConnector.company.GetNewObjectKey
                                    'DatosGenereral.nMensaje = "[P01-001 SMR] The Purcharse Order was successfully added, the No it's " & _NoDocEntry

                                    DatosGenereral.nMensaje = ""
                                    _SuccessSaved = True
                                    'End If

                                    DatosGenereral.SiGrabo = 1
                                    '_Exception = False

                                    'Llamada comentarios--------------------------------------------------------------------------------
                                    Dim myarrayCommen(0) As String
                                    Dim fieldsComment As String = ""
                                    Dim sigrabo As Boolean = False
                                    Dim Linea As Integer = 0

                                    SQLstring = "Select TOP (1) U_Linea From [@BITAOPOR] WHERE U_DocEntry = '" & _NoDocEntry & "' ORDER BY U_Linea DESC "

                                    Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                                    If rSet2.RecordCount > 0 Then
                                        Linea = rSet2.Fields.Item(0).Value + 1
                                    End If

                                    fieldsComment = " Code= '" & "x" & "'"
                                    fieldsComment = fieldsComment & " Name='" & "S" & "'"
                                    fieldsComment = fieldsComment & " U_Fecha= '" & Now & "'"
                                    fieldsComment = fieldsComment & " U_HechoPor= '" & "Sha" & "'"
                                    fieldsComment = fieldsComment & " U_DocEntry='" & _NoDocEntry & "'"           '//RequesId
                                    fieldsComment = fieldsComment & " U_Linea= '" & Linea & "'"                     '//Control Interno de Linea
                                    fieldsComment = fieldsComment & " U_Comentario= '" & vOrder.Comments & "'"    '//Comentario

                                    myarrayCommen(0) = fieldsComment

                                    sigrabo = CommentsRq(myarrayCommen, _NoDocEntry, Linea, "A")

                                    If sigrabo = False Then
                                        DatosGenereral.nMensaje = "No Grabo Detalle"
                                    End If
                                    '------------------------------------------------------------------------------------------------------

                                Else

                                    DatosGenereral.SiGrabo = 35
                                    DatosGenereral.nMensaje = "[P01-035 SMR] Error SAP: " + SAPConnector.company.GetLastErrorDescription

                                    '_Exception = True
                                    _SuccessSaved = False
                                End If

                            Else
                                _Exception = True
                                _SuccessSaved = False
                            End If
                        Else
                            _SuccessSaved = False
                        End If
                    Else
                        _Exception = True
                        _SuccessSaved = False
                    End If

                    If _AApprove.ToUpper = "YES" Then

                    Else
                        If _SuccessSaved Then
                            _Exception = True
                            _WarningMsg = "Automatic Auto Approve Not allowed"
                            'Else
                            '   _WarningMsg = ""
                        End If
                    End If

                    'Graba el Mensaje para Control Interno
                    If vOrder.GetByKey(_NoDocEntry) Then

                        If DatosGenereral.nMensaje.Trim = "" Then
                            vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = DatosGenereral.nMensaje
                        Else
                            vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = _WarningMsg
                        End If
                        vOrder.Update()
                    End If

                    DatosGenereral.nMensaje = DatosGenereral.nMensaje & "RequesId: " & _NoDocEntry & " Exception: " & _Exception & " Success: " & _SuccessSaved & " Warning: " & _WarningMsg
                    Return DatosGenereral.nMensaje

            End Select

        End If  '//RequestComeFrom

    End Function
#End Region

#Region "RR Req"
    'POST de SMR Request Addrequest
    'Programador: Saul
    'Fecha: 02/10/2019
    'Solicitud de Daños
    Function RRAddrequest(ByVal content() As String, ByVal contentLine() As String, _Contrato As String, _Placa As String, _MaxQuantity As Integer, _MaxAmount As Double, _MaxOutstanding As Integer, _DealerReference As String, _CreatedBy As String, ByRef _Exception As Boolean, ByRef _NoDocEntry As Integer, _AApprove As String, _RequestComeFrom As String, _Action As String, ByRef _SuccessSaved As Boolean, ByRef _WarningMsg As String) As String
        If Not connected Then connect()

        Dim matches As MatchCollection

        Dim nErr As Long = 0
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        'Create the BusinessPartners object
        Dim vOrder As SAPbobsCOM.Documents

        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim lngKey As String = ""

        Dim nTotalLinea As Double = 0.00
        Dim nPrecio As Double = 0.00
        Dim nCanti As Integer = 0.00
        Dim SQLstring As String = ""
        Dim Proveedor As String = ""
        Dim CanPo As Integer = 0
        Dim CanOutPo As Integer = 0
        Dim TotPo As Double = 0.00
        Dim Meses As Integer = 0
        Dim MileContrato As Integer = 0
        Dim MileConsumidas As Integer = 0
        Dim FechaPO As String = ""
        Dim CuenConta As String = ""
        Dim TipoCuenta As String = ""
        Dim EstadoSolicita As String = ""
        Dim Moneda As String = ""
        Dim NomProvee As String = ""
        Dim CodigoClie As String = ""
        Dim Direccion As String = ""
        Dim ActivoFijo As String = ""
        Dim SerieUnida As String = ""

        If _Action = "Add" Then
            _NoDocEntry = 0
        End If

        'Crea el Objeto
        vOrder = getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

        'Busco el Nombre
        If DatosGenereral.SiGrabo = 0 Then
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "U_ESTADOPO" Then
                    EstadoSolicita = Match.Groups.Item(2).Value

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                    Exit For
                Else
                    DatosGenereral.SiGrabo = 27
                    DatosGenereral.nMensaje = "[P03-034 Replacement Request] The U_EstadoPO field is not included"
                End If

            Next Match

        End If

        'Dato Cancelado
        If DatosGenereral.SiGrabo = 0 Then

            If EstadoSolicita = "0009" Then
                DatosGenereral.SiGrabo = 27
                DatosGenereral.nMensaje = "[P03-035 Replacement Request] The Status is Canceled, it cann't be modified "
            End If

        End If


        If DatosGenereral.SiGrabo = 0 Then

            If Not vOrder.GetByKey(_NoDocEntry) And _Action = "Update" Then
                DatosGenereral.SiGrabo = 33
                DatosGenereral.nMensaje = "[P03-033 Replacement Request]  The Purchase Order Does Not Exist "
            End If

        End If

        'Verifica si la cantidad de Ordenes de Compra llegaron a su limpite por Proveedor
        'Condicion No. 1-------------------------------------------------------------------
        If DatosGenereral.SiGrabo = 0 Then

            'Busco el Nombre
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "CARDCODE" Then
                    Proveedor = Match.Groups.Item(2).Value

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                    Exit For
                Else
                    DatosGenereral.SiGrabo = 25
                    DatosGenereral.nMensaje = "[P03-025 Replacement Request] The CardCode field is not included.."
                End If

            Next Match

            If DatosGenereral.SiGrabo = 0 Then
                'Script para ver datos de Contrato
                SQLstring = <sql>
                      Select a.CardCode,  
                        SUM((CASE when a.DocCur = 'QTZ' then DocTotal  else DocTotalFC END) *  a.DocRate) As TotPO,
                        COUNT(a.CardCode) As CantPO
                      From admon.dbo.OPOR a
                         inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                         inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                         where b.ObjectCode = 22 AND b.Series IN(130,12)
                         and est.name != 'Cancelled'
                         and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                         and a.CANCELED = 'N'
                         and a.DocStatus = 'O'
                         AND CardCode = '<%= Proveedor %>'
                         GROUP BY a.CardCode
                 </sql>.Value

                Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                If rSet.RecordCount > 0 Then
                    rSet.MoveFirst()
                    While Not rSet.EoF

                        CanPo = rSet.Fields.Item(2).Value
                        TotPo = rSet.Fields.Item(1).Value

                        rSet.MoveNext()
                    End While
                Else
                    DatosGenereral.SiGrabo = 26
                    DatosGenereral.nMensaje = "[P03-026 Replacement Request] The Supplier don't Exist "
                End If
            End If
        End If

        If DatosGenereral.SiGrabo = 0 Then

            If _Action = "Update" Then

                If (CanPo) >= _MaxQuantity And _MaxQuantity > 0 Then
                    DatosGenereral.SiGrabo = 37
                    _WarningMsg = "[P03-037 Replacement Request] Supplier exceeds maximum requests quantity " & " contact Leasing company"
                End If
            Else

                If (CanPo + 1) >= _MaxQuantity And _MaxQuantity > 0 Then
                    DatosGenereral.SiGrabo = 37
                    _WarningMsg = "[P03-037 Replacement Request] Supplier exceeds maximum requests quantity " & " contact Leasing company"
                End If

            End If

        End If
        'Fin de Condicion No. 1 ---------------------------------------------------------------


        'Verifica si el Monto de Ordenes de Compra llegaron a su limpite por Proveedor
        'Condicion No. 2-------------------------------------------------------------------
        If DatosGenereral.SiGrabo = 0 Then
            Dim TotActPo As Double = 0.00

            'Busco el Nombre
            matches = Regex.Matches(contentLine(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "PRICE" Then
                    TotActPo = Math.Round(Match.Groups.Item(2).Value * 1, 2) + TotActPo

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                    Exit For
                Else
                    DatosGenereral.SiGrabo = 27
                    DatosGenereral.nMensaje = "[P03-027 Replacement Request] The Price field is not included "
                End If

            Next Match


            If DatosGenereral.SiGrabo = 0 Then

                If _Action = "Update" Then

                    If (TotPo) >= _MaxAmount And _MaxAmount > 0 Then
                        DatosGenereral.SiGrabo = 3
                        _WarningMsg = "[P03-003 Replacement Request] Supplier exceeds maximum total amount "
                    End If

                Else

                    If TotActPo = 0 Then
                        DatosGenereral.SiGrabo = 28
                        DatosGenereral.nMensaje = "[P03-028 Replacement Request] The requests need one Price "
                    Else

                        If (TotPo + TotActPo) >= _MaxAmount And _MaxAmount > 0 Then
                            DatosGenereral.SiGrabo = 3
                            _WarningMsg = "[P03-003 Replacement Request] Supplier exceeds maximum total amount "
                        End If
                    End If
                End If

            End If
        End If
        'Fin de Condicion No. 2 ---------------------------------------------------------------

        'Verifica si el Max Asiganaciones Sobresalientes de Ordenes de Compra llegaron 
        'a su limpite por Proveedor y por SMR
        'Condicion No. 3-------------------------------------------------------------------
        If DatosGenereral.SiGrabo = 0 Then
            'Script para ver datos de Contrato
            SQLstring = <sql>
                     Select COUNT(a.DocEntry) As CantOutPO
                     From OPOR a
                         inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                         inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                         where b.ObjectCode = 22 AND b.Series IN(130,12)
                         and est.name != 'Cancelled'
                         and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                         and a.CANCELED = 'N'
                         AND CardCode = '<%= Proveedor %>'
                         and a.DocStatus = 'O'
                         AND a.U_TypeofRequest = 'P02'  --Damage
                        GROUP BY a.CardCode
                 </sql>.Value

            Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet2.RecordCount > 0 Then
                rSet2.MoveFirst()
                While Not rSet2.EoF

                    CanOutPo = rSet2.Fields.Item(0).Value

                    rSet2.MoveNext()
                End While
            Else
                DatosGenereral.SiGrabo = 29
                DatosGenereral.nMensaje = "[P03-029 Damage] The Supplier don't Exist "
            End If

            If DatosGenereral.SiGrabo = 0 Then
                If CanOutPo + 1 >= _MaxOutstanding And _MaxOutstanding > 0 Then
                    DatosGenereral.SiGrabo = 4
                    _WarningMsg = "[P03-004 Damage] Supplier exceeds maximum requests Outstanding Assignment " & "contact Leasing company"
                End If
            End If
        End If
        'Fin de Condicion No. 3 ---------------------------------------------------------------


        If DatosGenereral.SiGrabo = 0 Then
            'Script para ver datos de Contrato
            SQLstring = <sql>
                             SELECT TOP (1) a.ItemCode,                                    
                                a.AttriTxt2,                     
                                b.U_TipoServicio,
                                d.CardName,
                                ISNULL(d.Phone1,'')+' '+ISNULL(d.Phone2,'') As TelephoneNo,
                                d.CardCode,
                                e.Name,
                                e.U_Identificacion,
                                e.E_MailL,
                                b.U_KilometrajeCon,
                                b.U_TipoServicio
                            FROM ITM13 a 
                            INNER Join OITM b ON b.ItemCode = a.ItemCode 
                            INNER JOIN [@MCONTRATO] c ON c.U_NumContrato = b.U_Contrato 
                            INNER JOIN [OCRD] d ON d.U_CodigoMilenia = c.U_CodCliente 
                            INNER JOIN [OCPR] e ON e.CardCode = d.CardCode 
                            WHERE a.AttriTxt1 = '<%= _Placa %>'
                            AND c.U_NumContrato = '<%= _Contrato %>'
                            AND c.U_CodEstado = 'F'
                   </sql>.Value

            Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet.RecordCount > 0 Then
                rSet.MoveFirst()
                While Not rSet.EoF

                    'Campos Adicionales
                    NomProvee = rSet.Fields.Item(3).Value     '//Nombre Proveedor
                    CodigoClie = rSet.Fields.Item(5).Value    '//Codigo del Cliente
                    Direccion = rSet.Fields.Item(4).Value      '//Direccion
                    ActivoFijo = rSet.Fields.Item(0).Value    '//Codigo Articulo
                    SerieUnida = rSet.Fields.Item(1).Value    '//Serie Unidad

                    'Cuenta contable 
                    TipoCuenta = rSet.Fields.Item(10).Value

                    rSet.MoveNext()
                    Exit While
                End While

            Else
                DatosGenereral.SiGrabo = 29
                DatosGenereral.nMensaje = "[P03-029 Replacement Request] The Contract or Plancese don't Exist "
            End If
        End If


        'Obtiene el No. de Cuenta Contable
        If DatosGenereral.SiGrabo = 0 Then

            'Busco el la Moneda
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "DOCCURRENCY" Then
                    Moneda = Match.Groups.Item(2).Value

                    If Moneda = "GTQ" Then
                        Moneda = "QTZ"
                    End If

                    Exit For
                End If

            Next Match

            If Moneda = "USD" Then
                'Script para ver datos de Contrato
                SQLstring = <sql>
                            Select b.AcctCode, 'USD' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostosD,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
            Else
                'Script para ver datos de Contrato
                SQLstring = <sql>
                            Select b.AcctCode, 'QTZ' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostos,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
            End If

            Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet2.RecordCount > 0 Then
                rSet2.MoveFirst()
                While Not rSet2.EoF

                    CuenConta = rSet2.Fields.Item(0).Value

                    rSet2.MoveNext()
                End While
            End If
        End If

        If _Action = "Add" Then

            If _AApprove.ToUpper = "YES" Then

                If DatosGenereral.SiGrabo = 0 Then  '//Grabada sin Errores
                    EstadoSolicita = "0001"  '//Approved
                    _Exception = False
                Else
                    If DatosGenereral.SiGrabo > 0 And DatosGenereral.SiGrabo < 25 Then
                        EstadoSolicita = "0007"  '//Requested
                    End If
                    _Exception = True
                End If

            Else  '//Aprove = "NO"

                EstadoSolicita = "0007"  '//Requested
                _Exception = False

            End If

        Else

            If _AApprove.ToUpper = "YES" Then

                If DatosGenereral.SiGrabo = 0 Then  '//Grabada sin Errores
                    _Exception = False
                Else
                    _Exception = True
                End If

            Else  '//Aprove = "NO"

                _Exception = False

            End If

        End If

        'Inicia la Recoleccion del Array con todos los campos para Aderirlos en SAP
        If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then
            Try

                'Agrega la Cabecera
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches
                    'Match.Groups.Item(1).Value = Campo 
                    'Match.Groups.Item(2).Value = 'Valor

                    If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                        vOrder.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                    Else
                        CallByName(vOrder, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                    End If

                Next Match


                If _Action = "Add" Then

                    'Agrega el Deatalle de Lines
                    matches = Regex.Matches(contentLine(0), fieldValuePattern)
                    For Each Match As Match In matches
                        If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                            vOrder.Lines.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                        End If
                    Next Match

                    vOrder.Lines.ItemCode = "9001"
                    vOrder.Lines.Quantity = "1"
                    vOrder.Lines.AccountCode = CuenConta
                Else
                    vOrder.Lines.SetCurrentLine("0")
                    vOrder.Lines.Delete()

                    matches = Regex.Matches(contentLine(0), fieldValuePattern)
                    For Each Match As Match In matches
                        If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                            vOrder.Lines.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                        End If
                    Next Match

                    vOrder.Lines.ItemCode = "9001"
                    vOrder.Lines.Quantity = "1"
                    vOrder.Lines.AccountCode = CuenConta
                End If

                paso = True

            Catch ex As Exception

                DatosGenereral.SiGrabo = 31
                DatosGenereral.nMensaje = "[P03-031 Replacement Request] Error SAP: " + ex.Message

                paso = False

            End Try

            'Actualiza datos
            If paso Then
                If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then

                    vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"

                    If DatosGenereral.SiGrabo > 0 And DatosGenereral.SiGrabo < 25 Then
                        vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = DatosGenereral.nMensaje
                    End If

                    'Atividades Estatus
                    Select Case EstadoSolicita
                        Case = "0007"       '//Requested
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0001"   '//Initial [@SMR_ESTADO_ACT]
                        Case = "0001"       '//Approved
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0002"   '//Approved
                        Case = "0006"       '//Rejected
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0003"   '//Approved
                        Case = "0009"       '//Canceled
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0004"   '//Approved
                    End Select


                    If _Action = "Add" Then

                        vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                        vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                        vOrder.UserFields.Fields.Item("U_Estado").Value = "FORMALIZADO"
                        vOrder.UserFields.Fields.Item("U_CreatedBy").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_Placa").Value = _Placa
                        vOrder.UserFields.Fields.Item("U_NContrato").Value = _Contrato
                        vOrder.UserFields.Fields.Item("U_NumCredito").Value = _Contrato
                        vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = 0
                        vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita
                        vOrder.UserFields.Fields.Item("U_RequestComeFrom").Value = _RequestComeFrom
                        vOrder.UserFields.Fields.Item("U_OrdenServicio").Value = _DealerReference


                        vOrder.UserFields.Fields.Item("U_NomCredito").Value = NomProvee     '//Nombre Proveedor
                        vOrder.UserFields.Fields.Item("U_Cliente").Value = CodigoClie       '//Codigo del Cliente
                        vOrder.UserFields.Fields.Item("U_Direccion").Value = Direccion      '//Direccion
                        vOrder.UserFields.Fields.Item("U_CodUnidad").Value = ActivoFijo     '//Codigo Articulo
                        vOrder.UserFields.Fields.Item("U_SerieUnidad").Value = SerieUnida   '//Serie Unidad

                        vOrder.DocCurrency = Moneda

                        vOrder.Series = 130
                        vOrder.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders

                        chk = vOrder.Add()
                    Else
                        vOrder.DocCurrency = Moneda

                        vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita
                        vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = Convert.ToInt32(vOrder.UserFields.Fields.Item("U_Actualizaciones").Value) + 1

                        chk = vOrder.Update()
                    End If

                    If (chk = 0) Then

                        If _Action = "Add" Then
                            _NoDocEntry = SAPConnector.company.GetNewObjectKey
                            'DatosGenereral.nMensaje = "[P01-001 Replacement Request] The Purcharse Order was successfully added, the No it's " & _NoDocEntry
                            'DatosGenereral.nMensaje = "[P01-001 Replacement Request] The Purcharse Order was successfully added, the No it's " & _NoDocEntry
                        Else
                            'DatosGenereral.nMensaje = "[P01-001 Replacement Request] The Purcharse Order was successfully Updated, the No it's " & _NoDocEntry

                        End If
                        DatosGenereral.nMensaje = ""

                        DatosGenereral.SiGrabo = 1
                        _SuccessSaved = True


                        'Llamada comentarios--------------------------------------------------------------------------------
                        Dim myarrayCommen(0) As String
                        Dim fieldsComment As String = ""
                        Dim sigrabo As Boolean = False
                        Dim Linea As Integer = 0

                        SQLstring = "Select TOP (1) U_Linea From [@BITAOPOR] WHERE U_DocEntry = '" & _NoDocEntry & "' ORDER BY U_Linea DESC "

                        Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet2.RecordCount > 0 Then
                            Linea = rSet2.Fields.Item(0).Value + 1
                        End If

                        fieldsComment = " Code= '" & "x" & "'"
                        fieldsComment = fieldsComment & " Name='" & "S" & "'"
                        fieldsComment = fieldsComment & " U_Fecha= '" & Now & "'"
                        fieldsComment = fieldsComment & " U_HechoPor= '" & "Sha" & "'"
                        fieldsComment = fieldsComment & " U_DocEntry='" & _NoDocEntry & "'"           '//RequesId
                        fieldsComment = fieldsComment & " U_Linea= '" & Linea & "'"                     '//Control Interno de Linea
                        fieldsComment = fieldsComment & " U_Comentario= '" & vOrder.Comments & "'"    '//Comentario

                        myarrayCommen(0) = fieldsComment

                        sigrabo = CommentsRq(myarrayCommen, _NoDocEntry, Linea, "A")

                        If sigrabo = False Then
                            DatosGenereral.nMensaje = "No Grabo Detalle"
                        End If
                        '------------------------------------------------------------------------------------------------------

                    Else

                        DatosGenereral.SiGrabo = 32
                        DatosGenereral.nMensaje = "[P03-032 Replacement Request] Error SAP: " + SAPConnector.company.GetLastErrorDescription
                        _SuccessSaved = True
                    End If
                Else
                    _Exception = True
                    _SuccessSaved = False
                End If
            Else
                _SuccessSaved = False
            End If
        Else
            _Exception = True
            _SuccessSaved = False
        End If
        '-- Termina la grabación exitosa de los campos a SAP

        If _AApprove.ToUpper = "YES" Then

        Else
            If _SuccessSaved Then
                _Exception = True
                _WarningMsg = "Automatic Auto Approve Not allowed"
            End If
        End If

        'Graba el Mensaje para Control Interno
        If vOrder.GetByKey(_NoDocEntry) Then

            If DatosGenereral.nMensaje.Trim = "" Then
                vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = DatosGenereral.nMensaje
            Else
                vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = _WarningMsg
            End If
            vOrder.Update()
        End If

        DatosGenereral.nMensaje = DatosGenereral.nMensaje & "RequesId: " & _NoDocEntry & " Exception: " & _Exception & " Success: " & _SuccessSaved & " Warning: " & _WarningMsg
        Return DatosGenereral.nMensaje

    End Function

#End Region


#Region "RPL Req"
    'POST de Replacement Request Addrequest
    'Programador: Saul
    'Fecha: 27/01/201
    'Solicitud de Daños
    Function RPLAddrequest(ByVal content() As String, ByVal contentLine() As String, _Contrato As String, _Placa As String, _MaxQuantity As Integer, _MaxAmount As Double, _MaxOutstanding As Integer, _DealerReference As String, _CreatedBy As String, ByRef _Exception As Boolean, ByRef _NoDocEntry As Integer, _AApprove As String, _RequestComeFrom As String, _Action As String, ByRef _SuccessSaved As Boolean, ByRef _WarningMsg As String) As String
        If Not connected Then connect()

        Dim matches As MatchCollection

        Dim nErr As Long = 0
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        'Create the BusinessPartners object
        Dim vOrder As SAPbobsCOM.Documents

        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim lngKey As String = ""

        Dim nTotalLinea As Double = 0.00
        Dim nPrecio As Double = 0.00
        Dim nCanti As Integer = 0.00
        Dim SQLstring As String = ""
        Dim Proveedor As String = ""
        Dim CanPo As Integer = 0
        Dim CanOutPo As Integer = 0
        Dim TotPo As Double = 0.00
        Dim Meses As Integer = 0
        Dim MileContrato As Integer = 0
        Dim MileConsumidas As Integer = 0
        Dim FechaPO As String = ""
        Dim CuenConta As String = ""
        Dim TipoCuenta As String = ""
        Dim EstadoSolicita As String = ""
        Dim Moneda As String = ""
        Dim NomProvee As String = ""
        Dim CodigoClie As String = ""
        Dim Direccion As String = ""
        Dim ActivoFijo As String = ""
        Dim SerieUnida As String = ""

        If _Action = "Add" Then
            _NoDocEntry = 0
        End If

        'Crea el Objeto
        vOrder = getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

        'Busco el Nombre
        If DatosGenereral.SiGrabo = 0 Then
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "U_ESTADOPO" Then
                    EstadoSolicita = Match.Groups.Item(2).Value

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                    Exit For
                Else
                    DatosGenereral.SiGrabo = 27
                    DatosGenereral.nMensaje = "[P02-034 SMR] The U_EstadoPO field is not included"
                End If

            Next Match

        End If

        'Dato Cancelado
        If DatosGenereral.SiGrabo = 0 Then

            If EstadoSolicita = "0009" Then
                DatosGenereral.SiGrabo = 27
                DatosGenereral.nMensaje = "[P02-035 Damage] The Status is Canceled, it cann't be modified "
            End If

        End If


        If DatosGenereral.SiGrabo = 0 Then

            If Not vOrder.GetByKey(_NoDocEntry) And _Action = "Update" Then
                DatosGenereral.SiGrabo = 33
                DatosGenereral.nMensaje = "[P02-033 Damage]  The Purchase Order Does Not Exist "
            End If

        End If

        'Verifica si la cantidad de Ordenes de Compra llegaron a su limpite por Proveedor
        'Condicion No. 1-------------------------------------------------------------------
        If DatosGenereral.SiGrabo = 0 Then

            'Busco el Nombre
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "CARDCODE" Then
                    Proveedor = Match.Groups.Item(2).Value

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                    Exit For
                Else
                    DatosGenereral.SiGrabo = 25
                    DatosGenereral.nMensaje = "[P02-025 Damage] The CardCode field is not included.."
                End If

            Next Match

            If DatosGenereral.SiGrabo = 0 Then
                'Script para ver datos de Contrato
                SQLstring = <sql>
                      Select a.CardCode,  
                        SUM((CASE when a.DocCur = 'QTZ' then DocTotal  else DocTotalFC END) *  a.DocRate) As TotPO,
                        COUNT(a.CardCode) As CantPO
                      From admon.dbo.OPOR a
                         inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                         inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                         where b.ObjectCode = 22 AND b.Series IN(130,12)
                         and est.name != 'Cancelled'
                         and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                         and a.CANCELED = 'N'
                         and a.DocStatus = 'O'
                         AND CardCode = '<%= Proveedor %>'
                         GROUP BY a.CardCode
                 </sql>.Value

                Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                If rSet.RecordCount > 0 Then
                    rSet.MoveFirst()
                    While Not rSet.EoF

                        CanPo = rSet.Fields.Item(2).Value
                        TotPo = rSet.Fields.Item(1).Value

                        rSet.MoveNext()
                    End While
                Else
                    DatosGenereral.SiGrabo = 26
                    DatosGenereral.nMensaje = "[P02-026 Damage] The Supplier dosen't Exist "
                End If
            End If
        End If

        If DatosGenereral.SiGrabo = 0 Then

            If _Action = "Update" Then

                If (CanPo) >= _MaxQuantity And _MaxQuantity > 0 Then
                    DatosGenereral.SiGrabo = 36
                    _WarningMsg = "[P02-036 Damage] Supplier exceeds maximum requests quantity " & " contact Leasing company"
                End If
            Else

                If (CanPo + 1) >= _MaxQuantity And _MaxQuantity > 0 Then
                    DatosGenereral.SiGrabo = 37
                    _WarningMsg = "[P02-037 Damage] Supplier exceeds maximum requests quantity " & " contact Leasing company"
                End If

            End If

        End If
        'Fin de Condicion No. 1 ---------------------------------------------------------------


        'Verifica si el Monto de Ordenes de Compra llegaron a su limpite por Proveedor
        'Condicion No. 2-------------------------------------------------------------------
        If DatosGenereral.SiGrabo = 0 Then
            Dim TotActPo As Double = 0.00

            'Busco el Nombre
            matches = Regex.Matches(contentLine(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "PRICE" Then
                    TotActPo = Math.Round(Match.Groups.Item(2).Value * 1, 2) + TotActPo

                    DatosGenereral.SiGrabo = 0
                    DatosGenereral.nMensaje = ""
                    Exit For
                Else
                    DatosGenereral.SiGrabo = 27
                    DatosGenereral.nMensaje = "[P02-027 Damage] The Price field is not included "
                End If

            Next Match


            If DatosGenereral.SiGrabo = 0 Then

                If _Action = "Update" Then

                    If (TotPo) >= _MaxAmount And _MaxAmount > 0 Then
                        DatosGenereral.SiGrabo = 3
                        _WarningMsg = "[P02-003 Damage] Supplier exceeds maximum total amount "
                    End If

                Else

                    If TotActPo = 0 Then
                        DatosGenereral.SiGrabo = 28
                        DatosGenereral.nMensaje = "[P02-028 Damage] The requests need one Price "
                    Else

                        If (TotPo + TotActPo) >= _MaxAmount And _MaxAmount > 0 Then
                            DatosGenereral.SiGrabo = 3
                            _WarningMsg = "[P02-003 Damage] Supplier exceeds maximum total amount "
                        End If
                    End If
                End If

            End If
        End If
        'Fin de Condicion No. 2 ---------------------------------------------------------------

        'Verifica si el Max Asiganaciones Sobresalientes de Ordenes de Compra llegaron 
        'a su limpite por Proveedor y por SMR
        'Condicion No. 3-------------------------------------------------------------------
        If DatosGenereral.SiGrabo = 0 Then
            'Script para ver datos de Contrato
            SQLstring = <sql>
                     Select COUNT(a.DocEntry) As CantOutPO
                     From OPOR a
                         inner join NNM1 b on b.ObjectCode = a.objType AND b.Series = a.Series
                         inner join [@SMR_ESTADOS_PO] est on a.U_EstadoPO = est.Code
                         where b.ObjectCode = 22 AND b.Series IN(130,12)
                         and est.name != 'Cancelled'
                         and est.DocEntry NOT IN(6,11)   --6 = Rejeted, 9 = Cancelled
                         and a.CANCELED = 'N'
                         AND CardCode = '<%= Proveedor %>'
                         and a.DocStatus = 'O'
                         AND a.U_TypeofRequest = 'P02'  --Replacement
                        GROUP BY a.CardCode
                 </sql>.Value

            Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet2.RecordCount > 0 Then
                rSet2.MoveFirst()
                While Not rSet2.EoF

                    CanOutPo = rSet2.Fields.Item(0).Value

                    rSet2.MoveNext()
                End While
            Else
                DatosGenereral.SiGrabo = 29
                DatosGenereral.nMensaje = "[P02-029 Replacement] The Supplier don't Exist "
            End If

            If DatosGenereral.SiGrabo = 0 Then
                If CanOutPo + 1 >= _MaxOutstanding And _MaxOutstanding > 0 Then
                    DatosGenereral.SiGrabo = 4
                    _WarningMsg = "[P02-004 Replacement] Supplier exceeds maximum requests Outstanding Assignment " & "contact Leasing company"
                End If
            End If
        End If
        'Fin de Condicion No. 3 ---------------------------------------------------------------


        If DatosGenereral.SiGrabo = 0 Then
            'Script para ver datos de Contrato
            SQLstring = <sql>
                             SELECT TOP (1) a.ItemCode,                                    
                                a.AttriTxt2,                     
                                b.U_TipoServicio,
                                d.CardName,
                                ISNULL(d.Phone1,'')+' '+ISNULL(d.Phone2,'') As TelephoneNo,
                                d.CardCode,
                                e.Name,
                                e.U_Identificacion,
                                e.E_MailL,
                                b.U_KilometrajeCon,
                                b.U_TipoServicio
                            FROM ITM13 a 
                            INNER Join OITM b ON b.ItemCode = a.ItemCode 
                            INNER JOIN [@MCONTRATO] c ON c.U_NumContrato = b.U_Contrato 
                            INNER JOIN [OCRD] d ON d.U_CodigoMilenia = c.U_CodCliente 
                            INNER JOIN [OCPR] e ON e.CardCode = d.CardCode 
                            WHERE a.AttriTxt1 = '<%= _Placa %>'
                            AND c.U_NumContrato = '<%= _Contrato %>'
                            AND c.U_CodEstado = 'F'
                   </sql>.Value

            Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet.RecordCount > 0 Then
                rSet.MoveFirst()
                While Not rSet.EoF

                    'Campos Adicionales
                    NomProvee = rSet.Fields.Item(3).Value     '//Nombre Proveedor
                    CodigoClie = rSet.Fields.Item(5).Value    '//Codigo del Cliente
                    Direccion = rSet.Fields.Item(4).Value      '//Direccion
                    ActivoFijo = rSet.Fields.Item(0).Value    '//Codigo Articulo
                    SerieUnida = rSet.Fields.Item(1).Value    '//Serie Unidad

                    'Cuenta contable 
                    TipoCuenta = rSet.Fields.Item(10).Value

                    rSet.MoveNext()
                    Exit While
                End While

            Else
                DatosGenereral.SiGrabo = 29
                DatosGenereral.nMensaje = "[P02-029 Damage] The Contract or Plancese don't Exist "
            End If
        End If


        'Obtiene el No. de Cuenta Contable
        If DatosGenereral.SiGrabo = 0 Then

            'Busco el la Moneda
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches

                If Match.Groups.Item(1).Value.ToUpper = "DOCCURRENCY" Then
                    Moneda = Match.Groups.Item(2).Value

                    If Moneda = "GTQ" Then
                        Moneda = "QTZ"
                    End If

                    Exit For
                End If

            Next Match

            If Moneda = "USD" Then
                'Script para ver datos de Contrato
                SQLstring = <sql>
                            Select b.AcctCode, 'USD' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostosD,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
            Else
                'Script para ver datos de Contrato
                SQLstring = <sql>
                            Select b.AcctCode, 'QTZ' As Mon 
                               From [@LOP] a
                               INNER JOIN [OACT] b ON b.ActId = Replace(a.U_CuentaCostos,'-','')
                               Where Code = '<%= TipoCuenta %>' 
                        </sql>.Value
            End If

            Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet2.RecordCount > 0 Then
                rSet2.MoveFirst()
                While Not rSet2.EoF

                    CuenConta = rSet2.Fields.Item(0).Value

                    rSet2.MoveNext()
                End While
            End If
        End If

        If _Action = "Add" Then

            If _AApprove.ToUpper = "YES" Then

                If DatosGenereral.SiGrabo = 0 Then  '//Grabada sin Errores
                    EstadoSolicita = "0001"  '//Approved
                    _Exception = False
                Else
                    If DatosGenereral.SiGrabo > 0 And DatosGenereral.SiGrabo < 25 Then
                        EstadoSolicita = "0007"  '//Requested
                    End If
                    _Exception = True
                End If

            Else  '//Aprove = "NO"

                EstadoSolicita = "0007"  '//Requested
                _Exception = False

            End If

        Else

            If _AApprove.ToUpper = "YES" Then

                If DatosGenereral.SiGrabo = 0 Then  '//Grabada sin Errores
                    _Exception = False
                Else
                    _Exception = True
                End If

            Else  '//Aprove = "NO"

                _Exception = False

            End If

        End If


        'Inicia la Recoleccion del Array con todos los campos para Aderirlos en SAP
        If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then
            Try

                'Agrega la Cabecera
                matches = Regex.Matches(content(0), fieldValuePattern)
                For Each Match As Match In matches
                    'Match.Groups.Item(1).Value = Campo 
                    'Match.Groups.Item(2).Value = 'Valor

                    If Left(Match.Groups.Item(1).Value, 2) = "U_" Then

                        If Match.Groups.Item(1).Value = "U_PerdidaTotal" _
                           Or Match.Groups.Item(1).Value = "U_DanosTerceros" _
                           Or Match.Groups.Item(1).Value = "U_Robo" _
                           Or Match.Groups.Item(1).Value = "U_OtroCulpable" _
                           Or Match.Groups.Item(1).Value = "U_DanoBateria" _
                           Or Match.Groups.Item(1).Value = "U_DanoCubierto" Then

                            vOrder.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = IIf(Match.Groups.Item(2).Value = "True", "1", "0")
                        Else
                            vOrder.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        End If

                    Else
                        CallByName(vOrder, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                    End If

                Next Match

                If _Action = "Update" Then

                    'Borra todas las Lines 
                    Dim rSetII As SAPbobsCOM.Recordset = getRecordSet("Select LineNum from [POR1] WHERE DocEntry = '" & _NoDocEntry & "'")

                    If rSetII.RecordCount > 0 Then

                        rSetII.MoveFirst()
                        While Not rSetII.EoF

                            vOrder.Lines.SetCurrentLine("0")
                            vOrder.Lines.Delete()

                            rSetII.MoveNext()
                        End While
                    End If

                End If

                'Agrega el Deatalle de Lines
                matches = Regex.Matches(contentLine(0), fieldValuePattern)
                For Each Match As Match In matches
                    'Match.Groups.Item(1).Value = Campo 
                    'Match.Groups.Item(2).Value = 'Valor

                    If Left(Match.Groups.Item(1).Value, 2) = "U_" Then

                        If Match.Groups.Item(1).Value.Trim <> "U_LineNum" Then   '//Control de Lines
                            vOrder.Lines.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                        Else
                            If Match.Groups.Item(2).Value.Trim <> "" Then

                                'Suma mas de un Articulo
                                If Convert.ToInt32(Match.Groups.Item(2).Value.Trim) > 0 Then
                                    vOrder.Lines.Add()
                                End If

                            End If
                        End If

                    Else
                        'Para Poner el Numero de Cuenta
                        If Match.Groups.Item(1).Value.Trim = "AccountCode" Then
                            CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], CuenConta)
                        Else
                            CallByName(vOrder.Lines, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                        End If
                    End If
                Next Match

                paso = True

            Catch ex As Exception

                DatosGenereral.SiGrabo = 31
                DatosGenereral.nMensaje = "[P02-031 Damage] Error SAP: " + ex.Message

                paso = False

            End Try

            'Actualiza datos
            If paso Then
                If DatosGenereral.SiGrabo >= 0 And DatosGenereral.SiGrabo < 25 Then

                    vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"

                    If DatosGenereral.SiGrabo > 0 And DatosGenereral.SiGrabo < 25 Then
                        vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = DatosGenereral.nMensaje
                    End If

                    'Atividades Estatus
                    Select Case EstadoSolicita
                        Case = "0007"       '//Requested
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0001"   '//Initial [@SMR_ESTADO_ACT]
                        Case = "0001"       '//Approved
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0002"   '//Approved
                        Case = "0006"       '//Rejected
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0003"   '//Approved
                        Case = "0009"       '//Canceled
                            vOrder.Lines.UserFields.Fields.Item("U_EstadoActividad").Value = "0004"   '//Approved
                    End Select


                    If _Action = "Add" Then

                        'Campos Adicionales
                        vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"
                        vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                        vOrder.UserFields.Fields.Item("U_LugarPago").Value = "ARREND"
                        vOrder.UserFields.Fields.Item("U_Estado").Value = "FORMALIZADO"
                        vOrder.UserFields.Fields.Item("U_CreatedBy").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_Placa").Value = _Placa
                        vOrder.UserFields.Fields.Item("U_NContrato").Value = _Contrato
                        vOrder.UserFields.Fields.Item("U_NumCredito").Value = _Contrato
                        vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = 0
                        vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita
                        vOrder.UserFields.Fields.Item("U_RequestComeFrom").Value = _RequestComeFrom
                        vOrder.UserFields.Fields.Item("U_OrdenServicio").Value = _DealerReference

                        vOrder.UserFields.Fields.Item("U_NomCredito").Value = NomProvee     '//Nombre Proveedor
                        vOrder.UserFields.Fields.Item("U_Cliente").Value = CodigoClie       '//Codigo del Cliente
                        vOrder.UserFields.Fields.Item("U_Direccion").Value = Direccion      '//Direccion
                        vOrder.UserFields.Fields.Item("U_CodUnidad").Value = ActivoFijo     '//Codigo Articulo
                        vOrder.UserFields.Fields.Item("U_SerieUnidad").Value = SerieUnida   '//Serie Unidad

                        vOrder.DocCurrency = Moneda

                        vOrder.Series = 130
                        vOrder.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders

                        chk = vOrder.Add()
                    Else
                        vOrder.DocCurrency = Moneda

                        vOrder.UserFields.Fields.Item("U_TipoServicio").Value = TipoCuenta '"LTF"
                        vOrder.UserFields.Fields.Item("U_UsrActualizo").Value = _CreatedBy
                        vOrder.UserFields.Fields.Item("U_EstadoPO").Value = EstadoSolicita
                        vOrder.UserFields.Fields.Item("U_Actualizaciones").Value = Convert.ToInt32(vOrder.UserFields.Fields.Item("U_Actualizaciones").Value) + 1

                        chk = vOrder.Update()
                    End If

                    If (chk = 0) Then

                        If _Action = "Add" Then
                            _NoDocEntry = SAPConnector.company.GetNewObjectKey
                            'DatosGenereral.nMensaje = "[P01-001 Replacement Request] The Purcharse Order was successfully added, the No it's " & _NoDocEntry
                            'DatosGenereral.nMensaje = "[P01-001 Replacement Request] The Purcharse Order was successfully added, the No it's " & _NoDocEntry
                        Else
                            'DatosGenereral.nMensaje = "[P01-001 Replacement Request] The Purcharse Order was successfully Updated, the No it's " & _NoDocEntry
                        End If
                        DatosGenereral.nMensaje = ""

                        DatosGenereral.SiGrabo = 1
                        _SuccessSaved = True


                        'Llamada comentarios--------------------------------------------------------------------------------
                        Dim myarrayCommen(0) As String
                        Dim fieldsComment As String = ""
                        Dim sigrabo As Boolean = False
                        Dim Linea As Integer = 0

                        SQLstring = "Select TOP (1) U_Linea From [@BITAOPOR] WHERE U_DocEntry = '" & _NoDocEntry & "' ORDER BY U_Linea DESC "

                        Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

                        If rSet2.RecordCount > 0 Then
                            Linea = rSet2.Fields.Item(0).Value + 1
                        End If

                        fieldsComment = " Code= '" & "x" & "'"
                        fieldsComment = fieldsComment & " Name='" & "S" & "'"
                        fieldsComment = fieldsComment & " U_Fecha= '" & Now & "'"
                        fieldsComment = fieldsComment & " U_HechoPor= '" & "Sha" & "'"
                        fieldsComment = fieldsComment & " U_DocEntry='" & _NoDocEntry & "'"           '//RequesId
                        fieldsComment = fieldsComment & " U_Linea= '" & Linea & "'"                     '//Control Interno de Linea
                        fieldsComment = fieldsComment & " U_Comentario= '" & vOrder.Comments & "'"    '//Comentario

                        myarrayCommen(0) = fieldsComment

                        sigrabo = CommentsRq(myarrayCommen, _NoDocEntry, Linea, "A")

                        If sigrabo = False Then
                            DatosGenereral.nMensaje = "No Grabo Detalle"
                        End If
                        '------------------------------------------------------------------------------------------------------

                    Else

                        DatosGenereral.SiGrabo = 32
                        DatosGenereral.nMensaje = "[P02-032 Damage] Error SAP: " + SAPConnector.company.GetLastErrorDescription
                        _SuccessSaved = True
                    End If
                Else
                    _Exception = True
                    _SuccessSaved = False
                End If
            Else
                _SuccessSaved = False
            End If
        Else
            _Exception = True
            _SuccessSaved = False
        End If
        '-- Termina la grabación exitosa de los campos a SAP

        If _AApprove.ToUpper = "YES" Then

        Else
            If _SuccessSaved Then
                _Exception = True
                _WarningMsg = "Automatic Auto Approve Not allowed"
            End If
        End If

        'Graba el Mensaje para Control Interno
        If vOrder.GetByKey(_NoDocEntry) Then

            If DatosGenereral.nMensaje.Trim = "" Then
                vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = DatosGenereral.nMensaje
            Else
                vOrder.UserFields.Fields.Item("U_MOTIVO_RECHAZO").Value = _WarningMsg
            End If
            vOrder.Update()
        End If

        DatosGenereral.nMensaje = DatosGenereral.nMensaje & "RequesId: " & _NoDocEntry & " Exception: " & _Exception & " Success: " & _SuccessSaved & " Warning: " & _WarningMsg
        Return DatosGenereral.nMensaje

    End Function

#End Region

#Region "Comments"
    'POST de Comments
    'Programador: Saul
    'Fecha: 04/09/2019
    'Graba Datos Vales de Combustible
    Function CommentsRq(ByVal content() As String, ByVal _DocEntry As String, ByVal _Line As Integer, _Type As String) As Boolean
        'If Not connected Then connect()
        Dim matchesComen As MatchCollection

        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        'Create the BusinessPartners object
        Dim uTable As SAPbobsCOM.UserTable = getBObject("BITAOPOR")
        Dim SiGraboc As Integer = 0
        Dim paso As Boolean = False
        Dim MaxCorrela As String = CorreCode()
        Dim pasouno As Boolean = False
        DatosGenereral.nMensaje = ""

        'Verifica se Agrego el Name (es la llave del Driver)

        Dim Vehiculo As String = ""
        Dim FechaUltima As String = ""
        Dim SQLstring As String = ""
        Dim nDocume As String = ""


        SQLstring = "Select * from [OPOR] WHERE DocEntry = '" & _DocEntry & "'"

        Dim rSet As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

        If rSet.RecordCount = 0 Then

            SiGraboc = 2

        End If

        If SiGraboc = 0 And (_Type = "E" Or _Type = "D") Then

            SQLstring = "Select code from [@BITAOPOR] WHERE U_DocEntry = '" & _DocEntry & "' AND U_Linea = '" & _Line & "'"

            Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet2.RecordCount > 0 Then
                nDocume = rSet2.Fields.Item(0).Value
            Else
                nDocume = ""

                SiGraboc = 2
            End If

        Else

            SQLstring = "Select code from [@BITAOPOR] WHERE U_DocEntry = '" & _DocEntry & "' AND U_Linea = '" & _Line & "'"

            Dim rSet2 As SAPbobsCOM.Recordset = getRecordSet(SQLstring)

            If rSet2.RecordCount > 0 Then
                nDocume = ""
                SiGraboc = 2
            End If

        End If

        If uTable.GetByKey(nDocume) Then

            If _Type = "E" Then
                If SiGraboc = 0 Then

                    Try
                        matchesComen = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matchesComen
                            'Match.Groups.Item(1).Value = Campo 
                            'Match.Groups.Item(2).Value = 'Valor

                            If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                                uTable.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                            Else
                                If Match.Groups.Item(1).Value = "Code" Or Match.Groups.Item(1).Value = "Name" Then
                                    CallByName(uTable, Match.Groups.Item(1).Value, [Let], MaxCorrela)
                                Else
                                    CallByName(uTable, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                                End If
                            End If
                        Next Match

                        paso = True
                    Catch ex As Exception

                        SiGraboc = 2

                        paso = False

                    End Try

                    'Actualiza datos
                    If paso Then
                        If SiGraboc = 0 Then

                            'Actualiza la Placa para que sierva de relacion con el Vehiculo
                            uTable.UserFields.Fields.Item("U_DocEntry").Value = _DocEntry
                            uTable.UserFields.Fields.Item("U_Linea").Value = _Line

                            chk = uTable.Update()
                            If (chk = 0) Then
                                SiGraboc = 1
                                pasouno = True

                            Else
                                SAPConnector.company.GetLastError(nErr, errMsg)

                                If (0 <> nErr) Then
                                    ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                                    SiGraboc = 4
                                    pasouno = False

                                End If
                            End If
                        End If
                    End If
                End If

            Else

                If _Type = "D" Then
                    'Borra la linea del Log
                    chk = uTable.Remove()
                    If (chk = 0) Then
                        SiGraboc = 1
                        pasouno = True
                    Else
                        SAPConnector.company.GetLastError(nErr, errMsg)

                        If (0 <> nErr) Then
                            ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                            SiGraboc = 4
                            pasouno = False

                        End If
                    End If
                End If

            End If
        Else

            If _Type = "A" Then

                If SiGraboc = 0 Then
                    Try
                        matchesComen = Regex.Matches(content(0), fieldValuePattern)
                        For Each Match As Match In matchesComen
                            'Match.Groups.Item(1).Value = Campo 
                            'Match.Groups.Item(2).Value = 'Valor

                            If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                                uTable.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                            Else
                                If Match.Groups.Item(1).Value = "Code" Or Match.Groups.Item(1).Value = "Name" Then
                                    CallByName(uTable, Match.Groups.Item(1).Value, [Let], MaxCorrela)
                                Else
                                    CallByName(uTable, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                                End If
                            End If
                        Next Match


                        'fieldsComment = " Code= '" & "**" & "'"
                        'fieldsComment = fieldsComment & " Name='" & "**" & "'"
                        'fieldsComment = fieldsComment & " U_Fecha= '" & Now & "'"
                        'fieldsComment = fieldsComment & " U_HechoPor= '" & "Sha" & "'"
                        'fieldsComment = fieldsComment & " U_DocEntry='" & _NoDocEntry & "'"           '//RequesId
                        'fieldsComment = fieldsComment & " U_Linea= '" & "0" & "'"                     '//Control Interno de Linea
                        'fieldsComment = fieldsComment & " U_Comentario= '" & vOrder.Comments & "'"    '//Comentario

                        paso = True
                    Catch ex As Exception

                        SiGraboc = 2

                        paso = False
                    End Try

                    'Actualiza datos
                    If paso Then
                        If SiGraboc = 0 Then

                            'Actualiza la Placa para que sierva de relacion con el Vehiculo
                            uTable.UserFields.Fields.Item("U_DocEntry").Value = _DocEntry
                            uTable.UserFields.Fields.Item("U_Linea").Value = _Line

                            chk = uTable.Add()
                            If (chk = 0) Then
                                SiGraboc = 1
                                pasouno = True
                            Else
                                SAPConnector.company.GetLastError(nErr, errMsg)

                                If (0 <> nErr) Then
                                    ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                                    SiGraboc = 4
                                    pasouno = False

                                End If
                            End If
                        End If
                    End If
                End If
            End If

        End If

        Return pasouno
    End Function
#End Region

#Region "Otros"

    Function BorraLinea(ByVal _Key As String, ByVal _Lin As Integer) As String
        If Not connected Then connect()
        Dim matches As MatchCollection

        Dim nErr As Long = 0
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        'Create the BusinessPartners object
        Dim vOrder As SAPbobsCOM.Documents

        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim lngKey As String = ""

        'Crea el Objeto
        vOrder = getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

        If vOrder.GetByKey(_Key) Then
            ''Para Modificar
            Dim rSetII As SAPbobsCOM.Recordset = getRecordSet("Select LineNum from [POR1] WHERE DocEntry = '" & _Key & "' AND LineNum = '" & _Lin & "'")

            If rSetII.RecordCount > 0 Then
                Dim nLin2 As String = rSetII.Fields.Item(0).Value

                vOrder.Lines.SetCurrentLine(nLin2)      '//Modifica
                vOrder.Lines.Delete()
                paso = True
            End If

            'Actualiza datos
            If paso Then
                If DatosGenereral.SiGrabo = 0 Then

                    'vOrder.Series = 130   'Numero Asignado para Orden de Compra

                    chk = vOrder.Update()
                    If (chk = 0) Then

                        DatosGenereral.SiGrabo = 1
                        DatosGenereral.nMensaje = "Deleted Row " & _Lin

                    Else

                        'If (0 <> nErr) Then
                        If chk <> 0 Then

                            ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                            DatosGenereral.SiGrabo = 8
                            DatosGenereral.nMensaje = "Error SAP: " + SAPConnector.company.GetLastErrorDescription

                        End If
                    End If
                End If

            End If
        End If
        Return DatosGenereral.nMensaje

    End Function


    'Catalogos
    'POST de Fuel Card
    'Programador: Saul
    'Fecha: 04/09/2019
    'Graba Datos Vales de Combustible
    Function EditCatalogo(ByVal content() As String, ByVal _Tabla As String) As String
        If Not connected Then connect()
        Dim matches As MatchCollection

        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        'Create the BusinessPartners object
        Dim uTable As SAPbobsCOM.UserTable = getBObject(_Tabla)
        DatosGenereral.SiGrabo = 0
        Dim paso As Boolean = False
        Dim pasouno As Boolean = False
        DatosGenereral.nMensaje = ""

        Try
            matches = Regex.Matches(content(0), fieldValuePattern)
            For Each Match As Match In matches
                'Match.Groups.Item(1).Value = Campo 
                'Match.Groups.Item(2).Value = 'Valor

                If Left(Match.Groups.Item(1).Value, 2) = "U_" Then
                    uTable.UserFields.Fields.Item(Match.Groups.Item(1).Value).Value = Match.Groups.Item(2).Value
                Else
                    CallByName(uTable, Match.Groups.Item(1).Value, [Let], Match.Groups.Item(2).Value)
                End If
            Next Match

            paso = True
        Catch ex As Exception

            DatosGenereral.SiGrabo = 2
            DatosGenereral.nMensaje = "Error SAP: " + ex.Message

            paso = False

        End Try

        'Actualiza datos
        If paso Then
            If DatosGenereral.SiGrabo = 0 Then

                chk = uTable.Update()
                If (chk = 0) Then
                    DatosGenereral.SiGrabo = 1
                    DatosGenereral.nMensaje = "The Catalog was successfully updated "
                Else
                    SAPConnector.company.GetLastError(nErr, errMsg)

                    If (0 <> nErr) Then
                        ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

                        DatosGenereral.SiGrabo = 4
                        DatosGenereral.nMensaje = "Error SAP:" + Str(nErr) + "," + errMsg

                    End If
                End If

            End If
        End If

        Return DatosGenereral.nMensaje
    End Function
#End Region

End Class

