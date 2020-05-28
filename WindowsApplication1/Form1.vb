Imports ArrendUtils
Imports System.Data.SqlClient
Imports System.IO

Public Class Form1
    Public connx As New ArrendUtils.SAPConnector

    'Contenedor Conexion SAP 
    Private Shared _datosData As ArrendUtils.SAPConnector
    Public Shared Property DatosData() As ArrendUtils.SAPConnector
        Get
            Return _datosData
        End Get
        Set(ByVal value As ArrendUtils.SAPConnector)
            _datosData = value
        End Set
    End Property

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'connx.ConnString = New ArrendUtils.RegConnector("Software\MyDB").tuConexion
            'connx.transactional = False
            'testConn()

            'Dim tbl As DataTable = connx.execQuery("select Code from [@MCONTRATO] WHERE U_NumEmpresa = 500").ToTable
            'actualizaMail()
            'Throw New System.Exception("Excepción de prueba.")

            cambiarEM()
        Catch ex As Exception
            MsgBox(ex.HResult & vbCrLf & ex.Message & vbCrLf & ex.StackTrace)
            'Dim bita As New ArrendUtils.Bitacora()
            'bita.excepcion(ex, "BLeiva", "tester", "form1")
        Finally
            If connx.connected Then connx.disconnect()
        End Try
    End Sub

    Private Sub importaTasas()
        Dim connx2 As New SAPConnector("ARRENDDB\ARRENDDB:Milenia;sysadm,julio1;noimporta,noimporta")
        Dim dtable As DataTable = connx2.execQuery(<sql>
                             select codigo_moneda	U_CodigoMoneda
	                            , fecha_hora		U_FechaHora
	                            , tasa_compra		U_TasaCompra
	                            , tasa_venta		U_TasaVenta
	                            , num_empresa	U_NumEmpresa
                            from finasql.caja_tasa_cambio
                            where fecha_hora between '2015-05-01' and '2015-05-31'
                         </sql>.Value).Table

        Dim uTable As SAPbobsCOM.UserTable = connx.getBObject("MCAJA_TASA_CAMBIO")
        For Each row In dtable.Rows
            connx.insert(uTable, {""}, True)
            For column As Integer = 0 To 4
                uTable.UserFields.Fields.Item(dtable.Columns(column).ColumnName).Value = row.item(column)
            Next column
            connx.commit()
        Next row
    End Sub
    Private Sub llenaFechas()
        Dim failed As Integer = 0
        For Each linea In File.ReadLines("C:\Users\bleiva.LEASING\Documents\codes.txt")
            Try
                connx.updateOnKey(connx.getBObject("MCUOTA_CONTRATO"), {"U_FechaCalcMora=' '"}, linea)
            Catch ex As Exception
                failed += 1
                Continue For
            End Try
        Next linea
        MsgBox(failed)
    End Sub

    Sub testConn()
        connx.connect()
        MsgBox(connx.connected)
        connx.disconnect()
    End Sub

    Sub testPC()
        Dim pc As SAPbobsCOM.JournalEntries = connx.getBObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    End Sub

    Private Sub insertaPC()
        Dim msg As String = "", rtn As Long = 0
        Dim pc As SAPbobsCOM.JournalEntries = connx.getBObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        pc.Lines.ShortName = "C003262"
        'pc.Lines.ControlAccount = "_SYS00000001962"
        pc.Lines.Debit = 1000

        pc.Lines.Add()

        pc.Lines.AccountCode = "_SYS00000001962"
        pc.Lines.Credit = 1000

        pc.Lines.Add()

        connx.insert(pc, {""})
    End Sub

    Private Sub actualizaMail()
        For Each row As DataRow In connx.execQuery("SELECT CardCode From OCRD WHERE E_Mail <> 'migracionsap@leasing.com.gt' AND CardType = 'C'", _
                                                   New Dictionary(Of String, Object)).Table.Rows
            connx.updateOnKey(connx.getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), {""}, row(0), True)
            connx.OnHold.EmailAddress = "migracionsap@leasing.com.gt"
            connx.commit()
        Next
    End Sub

    Private Sub probar_mora()
        Dim calculador As New CalculadorMora
        Dim cadena As String = New RegConnector("Software\MyDB").tuConexion

        cadena = cadena.Substring(0, cadena.LastIndexOf(";"))

        With calculador
            .connString = cadena
            MsgBox(.mora(4, 16596, 15, "", "N"))
        End With
    End Sub

    Private Sub cambiarEM()
        'Dim connx As New SAPConnector

        'Hacemos conexión con SAPX
        Dim connstring As String = New ArrendUtils.RegConnector("Software\MyDB").tuConexion

        DatosData = New ArrendUtils.SAPConnector(connstring)
        DatosData.transactional = False
        DatosData.connect()

        Dim Data As String = connstring
        Data = Data.Substring(Data.IndexOf(":") + 1, Data.IndexOf(";") - Data.IndexOf(":") - 1)


        MsgBox("DB:  " & Data.Trim.ToUpper)
        'With connx
        ' .ConnString = New RegConnector("Software\MyDB").tuConexion
        ' .transactional = False

        Dim sn As SAPbobsCOM.BusinessPartners = DatosData.getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        Dim rSet As SAPbobsCOM.Recordset = DatosData.getRecordSet("SELECT top 5 CardName FROM OCRD order by CardName desc")

        If rSet.RecordCount = 0 Then Exit Sub Else rSet.MoveFirst()
        While Not rSet.EoF
            'If sn.GetByKey(rSet.Fields.Item(0).Value) Then
            '    sn.EmailAddress = "migracionsap@leasing.com.gt"
            '    sn.Update()
            'End If
            MsgBox("Cliente: " & rSet.Fields.Item("CardName").Value)

            rSet.MoveNext()
        End While

        DatosData.disconnect()
        ' End With
    End Sub

    Private Sub inserta_cheque()

        Dim connx As New ArrendUtils.SAPConnector

        With connx

            .transactional = False
            .ConnString = New RegConnector("Software\MyDB").tuConexion

            Dim pago As SAPbobsCOM.Payments = .getBObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)



            Dim contenidoChk() As String = {<keys
                                                CheckSum='100' DueDate='2015-05-18' CheckAccount='_SYS00000001260'

                                            />.ToString.Replace("""", "'")}
            Dim contenidoRCT() As String = {<keys
                                                CardCode='C002150' DocDate='2015-05-18' DueDate='2015-05-18'

                                            />.ToString.Replace("""", "'")}

            .insert(pago.Checks, contenidoChk)
            .insert(pago, contenidoRCT)

            .disconnect()
        End With

    End Sub

    Private Sub inserta_NC()
        Dim msg As String = "", rtn As Long = 0
        'Dim connx As New SAPConnector
        With connx
            .transactional = False
            .ConnString = New RegConnector("Software\MyDB").tuConexion
            Dim fac As SAPbobsCOM.Documents = .getBObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            Dim nc As SAPbobsCOM.Documents = .getBObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)

            fac.GetByKey(152953)


            For i = 0 To fac.Lines.Count - 1
                fac.Lines.SetCurrentLine(i)

                nc.Lines.BaseType = fac.Lines.LineType
                nc.Lines.BaseEntry = fac.DocEntry
                nc.Lines.BaseLine = fac.Lines.LineNum
                nc.Lines.ItemCode = fac.Lines.ItemCode
                nc.Lines.Quantity = fac.Lines.Quantity
                nc.Lines.TaxCode = fac.Lines.TaxCode
                nc.Lines.UnitPrice = fac.Lines.UnitPrice

                nc.Lines.Add()

                'connx.Company.GetLastError(rtn, msg)

                If rtn <> 0 Then MsgBox(msg)
            Next i


            nc.DocDate = fac.DocDate
            nc.DocDueDate = fac.DocDueDate
            nc.TaxDate = fac.TaxDate
            nc.DiscountPercent = fac.DiscountPercent
            nc.CardCode = fac.CardCode
            nc.DocTotal = fac.DocTotal

            nc.Add()

            'connx.Company.GetLastError(rtn, msg)

            If rtn <> 0 Then MsgBox(msg)


            '.insert(nc, {""})
            .disconnect()
        End With
    End Sub

    Private Sub inserta_recibo()
        Dim connx As New ArrendUtils.SAPConnector
        With connx
            .ConnString = New RegConnector("Software\MyDB").tuConexion

            Dim rec As SAPbobsCOM.Payments = .getBObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments) 'Recibo

            'Agregarle un Checke
            Dim contCheque() As String = {<keys


                                          />.ToString.Replace("""", "'")}

            'Agregar la cabecera
            'Insertarlo

        End With
    End Sub

    Private Sub testReader()
        Dim table As DataTable = connx.execQuery("SELECT Top 7 DocEntry FROM OINV", New Dictionary(Of String, Object)).Table
        For Each row In table.Rows
            MsgBox(row(0))
        Next
    End Sub


End Class
