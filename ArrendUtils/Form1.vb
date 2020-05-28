Imports System.Data.SqlClient
Imports System.IO

Public Class Form1

    Dim Datos As New DatosGenereral()
    Dim dTable As New DataTable
    Dim dTable2 As New DataTable
    Dim Mensaje As String = ""
    Dim voDocEntry As Integer = 0
    Dim voException As Boolean = False
    Dim voSuccessSaved As Boolean = False
    Dim voWarningMsg As String = ""


    Public connx As New ArrendUtils_AF.SAPConnector

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'llamadas()
        'SMR()           '//SMR
        Reemplazo()     '//Replacement
        'Daño()          '//Damage
        'FuelCard()
        'Drivers()
        'AsignVeh()

    End Sub

    Private Sub llamadas()

        Try
            connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion
            connx.transactional = False

            Dim nLicencia As String = "1001006805"
            Dim myarray(0) As String
            Dim fields As String = ""
            Dim voException As Boolean = False
            Dim docentryline As String = ""
            Dim MaxAmount As Double  '-Todo Tipo de Solicitud y por Supplier
            Dim MaxQuantity As Integer    '-Todo Tipo de Solicitud y por Supplier
            Dim MaxOutstanding As Integer '-Todo Tipo de Solicitud y por Supplier
            Dim MaxMonths As Integer '-Por Contrato
            Dim PerMeleage As Integer '-Por Contrato
            Dim myarray_linea(0) As String
            Dim fieldsline As String = ""
            Dim DaysInBetween As Integer = 0

            'U_Estado_PO
            '0001    Approved
            '0002    Initial
            '0003    New
            '0004    Partially Approved
            '0005    Planned
            '0006    Rejected
            '0007    Requested
            '0008    Rescheduled
            '0009    Cancelled

            '- = CALL CENTER
            fields = " U_TypeofRequest= '" & "P01" & "'"                        '//RequestTypeCode (P01,P02,P03)
            fields = fields & " U_EstadoPO='" & "0003" & "'"                    '//RequestStatusCode
            Dim nContrato As String = "30342"
            Dim nPlaca As String = "P0556GRZ"
            'Mileage
            fields = fields & " U_ContactName= '" & "GONZALO MORALES" & "'"             '//ContactName
            fields = fields & " U_FechaInicio='" & "2020-05-20" & "'"           '//StartDate
            fields = fields & " U_TelephoneNo= '" & "12345678" & "'"             '//TelephoneNo
            Dim dealerreference As String = "REF001"
            fields = fields & " U_CustomerContactNo= '" & "12345678" & "'"       '//CustomerContactNo
            fields = fields & " U_FechaCliente='" & "2020-05-20" & "'"          '//Customer Request Date
            fields = fields & " U_FechaProveedor='" & "2020-05-20" & "'"        '//Supplier proposed Date
            fields = fields & " U_ReplacementReason = '" & "RR1" & "'"          '//Replacement Reason 
            fields = fields & " CardCode='" & "P001186" & "'"                   '//SupplierId
            'Total amount
            fields = fields & " DocCurrency = '" & "GTQ" & "'"                  '//Currency Code

            '//Control interno de la DLL
            fields = fields & " DocDate='" & Now & "'"                          '//Fecha del día (para control de la fecha de la Orden de C)
            '//TotalAmout es calculado automaticamente por SAP y se muestra en la GET

            'Luego de este insert se debe llamar la otra rutina para agregar otra Lines de Comentarios 
            'en la tabla [@]
            fields = fields & " Comments = '" & "we are adding purchase order" & "'"

            Dim CreateBy As String = "Gonzalo Morales"
            '//Para Post y Put y el LastSeenBy es el mimo valor que Modifiedby y CreateBy

            '// El NoOfTime se trabaja de forma y se actualiza en el Campo 
            '// U_Actualizaciones tomando como base la Post = 0 y Put se incrementa en 1

            '// El detalle se lleva contralado en con un ITEM Dummy
            'fieldsline = fieldsline & " U_LineNum='" & "0" & "'"
            'fieldsline = fieldsline & " ItemCode= '" & "9600" & "'"  //CALL CENTER
            'fieldsline = fieldsline & " Quantity= '" & 1 & "'"
            'fieldsline = fieldsline & " Price='" & 0.0 & "'"
            'fieldsline = fieldsline & " AccountCode ='" & "_SYS00000002493" & "'"

            myarray(0) = fields
            myarray_linea(0) = fieldsline

            voDocEntry = 15682
            voException = False

            MaxAmount = 0
            MaxQuantity = 100
            MaxOutstanding = 0
            MaxMonths = 0
            PerMeleage = 0
            DaysInBetween = 0

            Dim AApprove = "Yes"
            Dim RequestComeFrom As String = "CALL CENTER"
            'P01 ="SMR"
            'P02 ="Damage"
            'P03 ="Replacement"

            Mensaje = connx.SMRRequestAddrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMeleage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)
            'Mensaje = connx.SMRRequestEditrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMeleage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)


        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
        Finally

            If DatosGenereral.SiGrabo > 1 Then
                'MsgBox(DatosGenereral.nMensaje & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
                MsgBox(Mensaje)
            Else
                If DatosGenereral.SiGrabo = 1 Then
                    MsgBox(Mensaje)
                End If
            End If
            If connx.connected Then connx.disconnect()
        End Try

    End Sub

    Private Sub SMR()

        Try
            connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion

            connx.transactional = False

            Dim nLicencia As String = "1001006805"
            Dim myarray(0) As String
            Dim fields As String = ""
            Dim docentryline As String = ""
            Dim MaxAmount As Double  '-Todo Tipo de Solicitud y por Supplier
            Dim MaxQuantity As Integer    '-Todo Tipo de Solicitud y por Supplier
            Dim MaxOutstanding As Integer '-Todo Tipo de Solicitud y por Supplier
            Dim MaxMonths As Integer '-Por Contrato
            Dim PerMileage As Integer '-Por Contrato
            Dim myarray_linea(0) As String
            Dim fieldsline As String = ""
            Dim DaysInBetween As Integer = 0

            'U_Estado_PO
            '0001    Approved
            '0002    Initial
            '0003    New
            '0004    Partially Approved
            '0005    Planned
            '0006    Rejected
            '0007    Requested
            '0008    Rescheduled
            '0009    Cancelled

            '-- "P01" = SMR
            '//Cabecera SMR
            fields = " U_TypeofRequest= '" & "P01" & "'"                        '//RequestTypeCode
            fields = fields & " U_EstadoPO='" & "0002" & "'"                    '//RequestStatusCode
            Dim nContrato As String = "30342"                                   '//ContractID
            Dim nPlaca As String = "P0556GRZ"                                   '//LIcensePlate
            fields = fields & " U_Kilometraje= '" & "9000000" & "'"             '//Mileage
            fields = fields & " U_ContactName= '" & "GONZALO MORALES" & "'"     '//ContactName
            fields = fields & " U_FechaInicio='" & "2020-05-20" & "'"           '//StartDate
            fields = fields & " U_TelephoneNo= '" & "12345678" & "'"            '//TelephoneNo
            Dim dealerreference As String = "REF001"                            '//DealerReferenceNo
            fields = fields & " U_CustomerContactNo= '" & "22181818" & "'"       '//CustomerContactNo
            fields = fields & " DocDate='" & Now & "'"                          '//CustomerRequestDate
            fields = fields & " U_FechaFinal='" & "2020-06-20" & "'"            '//SupplierProposedDate
            fields = fields & " U_ReplacementReason='" & "" & "'"               '//ReplacmentReasonCode
            fields = fields & " CardCode='" & "P001186" & "'"                   '//SupplierId
            'total amount is calculated by SAP
            fields = fields & " DocCurrency = '" & "GTQ" & "'"                  '//Currency Code

            '//Control interno de la DLL
            '//TotalAmout es calculado automaticamente por SAP y se muestra en la GET

            '// El NoOfTime se trabaja de forma y se actualiza en el Campo 
            '// U_Actualizaciones tomando como base la Post = 0 y Put se incrementa en 1

            'Detalle de Actividades SMR 
            'En SMR solo se permite tener 1 Articulo
            fieldsline = fieldsline & " U_LineNum='" & "0" & "'"
            fieldsline = fieldsline & " ItemCode= '" & "100" & "'"              '//Activity Code
            fieldsline = fieldsline & " U_Defecto= '" & "01" & "'"              '//Buscar Tabla de relacion.  DefectCode
            fieldsline = fieldsline & " U_TipoOper= '" & "A" & "'"              '//Buscar Tabla de relacion.  OperationCode
            fieldsline = fieldsline & " U_Ubicacion= '" & "GDI" & "'"           '//Buscar Tabla de relacion.  LocationCode
            'fieldsline = fieldsline & " U_RepairTime= '" & "60" & "'" create field in SAP
            fieldsline = fieldsline & " U_ManoDeObra='" & 3.0 & "'"             '// Labour Cost
            fieldsline = fieldsline & " U_CostoMateriales= '" & "100.00" & "'"  '//Material Cost
            fieldsline = fieldsline & " Price='" & 30.0 & "'"                   '//Amount
            fieldsline = fieldsline & " U_EstadoActividad= '" & "0002" & "'"    '//activityStatusCode
            fieldsline = fieldsline & " U_PrecioLimite= '" & "100.00" & "'"     '//Threshold
            fieldsline = fieldsline & " U_MotivoRechazo= '" & "" & "'"          '//RejectionReasonCode

            '//Control Interno de la DLL
            fieldsline = fieldsline & " Quantity= '" & 1 & "'"                  '//Cantidad
            fieldsline = fieldsline & " AccountCode ='" & "_SYS00000002493" & "'"
            'fieldsline = fieldsline & " U_LineNum='" & "0" & "'"

            'Luego de este insert se debe llamar la otra rutina para agregar otra Lines de Comentarios 
            'en la tabla [@]
            fields = fields & " Comments = '" & "ADD SMR" & "'"

            myarray(0) = fields
            myarray_linea(0) = fieldsline

            Dim CreateBy As String = "Nikita Solanky"
            '//Para Post y Put y el LastSeenBy es el mimo valor que Modifiedby y CreateBy
            MaxQuantity = 400
            MaxAmount = 300000
            MaxMonths = 15
            MaxOutstanding = 350
            PerMileage = 5
            voException = False 'voException definido global
            voDocEntry = 15696 'voDocEntry definido global
            Dim AApprove = "Yes"
            Dim RequestComeFrom As String = "NORMAL"
            DaysInBetween = 30
            voSuccessSaved = False 'SuccessSaved definido global
            voWarningMsg = "" 'Warningmsg definido global
            'Dim voAction = "Add" this field was removed because we use differnt methods for add and edit

            '"SMR"
            '"Damage"
            '"Replacement"

            'Mensaje = connx.SMRRequestAddrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMileage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)
            Mensaje = connx.SMRRequestEditrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMileage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
        Finally

            If DatosGenereral.SiGrabo > 1 Then
                'MsgBox(DatosGenereral.nMensaje & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
                'MsgBox(Mensaje)
            Else
                'If DatosGenereral.SiGrabo = 1 Then

                'MsgBox(Mensae & "RequesId " & voDocEntry & "Exception " & voException & "Success " & voSuccessSaved & "Warning " & voWarningMsg)

                'End If

                If connx.connected Then connx.disconnect()
            End If

            MsgBox(Mensaje)

            'MsgBox("RequesId " & voDocEntry)
            'MsgBox("Exception " & voException)
            'MsgBox("Success " & voSuccessSaved)
            'MsgBox("Warning " & voWarningMsg)
        End Try

    End Sub

    'Damage
    'Cambiar la cuenta 
    Private Sub Daño()

        Try
            connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion

            connx.transactional = False

            Dim nLicencia As String = "1001006805"
            Dim myarray(0) As String
            Dim fields As String = ""
            Dim docentryline As String = ""
            Dim MaxAmount As Double  '-Todo Tipo de Solicitud y por Supplier
            Dim MaxQuantity As Integer    '-Todo Tipo de Solicitud y por Supplier
            Dim MaxOutstanding As Integer '-Todo Tipo de Solicitud y por Supplier
            Dim MaxMonths As Integer '-Por Contrato
            Dim PerMeleage As Integer '-Por Contrato
            Dim myarray_linea(0) As String
            Dim fieldsline As String = ""
            Dim DaysInBetween As Integer = 0

            'U_Estado_PO
            '0001    Approved
            '0002    Initial
            '0003    New
            '0004    Partially Approved
            '0005    Planned
            '0006    Rejected
            '0007    Requested
            '0008    Rescheduled
            '0009    Cancelled

            '-- "P02" = Damage
            '//Cabecera Damage
            fields = " U_TypeofRequest= '" & "P02" & "'"                        '//RequestTypeCode
            fields = fields & " U_EstadoPO='" & "0007" & "'"                    '//RequestStatusCode
            Dim nContrato As String = "30342"                                   'ContractId
            Dim nPlaca As String = "P0556GRZ"                                   'License plate
            fields = fields & " U_Kilometraje= '" & "90000" & "'"               '//Mileage
            fields = fields & " U_ContactName= '" & "GONZALO MORALES" & "'"     '//ContactName
            fields = fields & " U_FechaInicio='" & "2020-05-20" & "'"           '//StartDate
            fields = fields & " U_TelephoneNo= '" & "12345678" & "'"            '//TelephoneNo
            Dim dealerreference As String = "REF001"                            'DealerReferenceNo
            fields = fields & " U_CustomerContactNo= '" & "12345678" & "'"      '//CustomerContactNo
            fields = fields & " DocDate='" & Now & "'"                          '//CustomerRequestDate
            fields = fields & " U_FechaFinal='" & "2020-05-20" & "'"           '//SupplierProposedDate
            fields = fields & " U_ReplacementReason='" & "" & "'"               '//ReplacmentReasonCode
            fields = fields & " CardCode='" & "P001186" & "'"                   '//SupplierId
            'Total amount is calculated in SAP
            fields = fields & " DocCurrency = '" & "GTQ" & "'"                  '//Currency Code

            '//Details Damage.' 
            'En SAP es Cabecera
            fields = fields & " U_DamageCode= '" & "0001" & "'"         'DamageTypeCode
            fields = fields & " U_PerdidaTotal= '" & "False" & "'"       'IsTotalLoss
            fields = fields & " U_DanosTerceros= '" & "False" & "'"      'TwoPartyDamage
            fields = fields & " U_Robo= '" & "False" & "'"               'Stolen
            fields = fields & " U_OtroCulpable= '" & "False" & "'"       'IsOtherPartyGuilty
            fields = fields & " U_DanoBateria= '" & "False" & "'"        'IsBatteryDamage
            fields = fields & " U_DanoCubierto= '" & "False" & "'"       'IsDamageCovered

            'INSERTAR
            'Detalle de Actividades Damage 
            'En Damage este permite tener mas de 1 Articulo y el control de U_LineNum
            'fieldsline = fieldsline & " U_LineNum='" & "0" & "'"
            fieldsline = fieldsline & " ItemCode= '" & "1146" & "'"             '//ActivityCode
            fieldsline = fieldsline & " U_Defecto= '" & "0005" & "'"              '//DefectCode
            fieldsline = fieldsline & " U_TipoOper= '" & "0003" & "'"              '//OperationCode
            fieldsline = fieldsline & " U_Ubicacion= '" & "0005" & "'"           '//LocationCode
            'Warranty Code is not managed by supplier                            'WarrantyCode
            'Tyre brand code was removed
            'fieldsline = fieldsline & " U_RepairTime= '" & "60" & "'"            '/RepairTime create field in Sap
            fieldsline = fieldsline & " U_ManoDeObra='" & "100" & "'"             '//Labour Cost
            fieldsline = fieldsline & " U_CostoMateriales= '" & "200" & "'"     '//Material Cost
            fieldsline = fieldsline & " Price='" & "300" & "'"                   '//Amount
            fieldsline = fieldsline & " U_EstadoActividad= '" & "0001" & "'"       '//activityStatusCode
            fieldsline = fieldsline & " U_MotivoRechazo= '" & "" & "'"          '//RejectionReasonCode
            fieldsline = fieldsline & " U_PrecioLimite= '" & "100.00" & "'"     '//Threshold

            '//Control Interno de la DLL
            fieldsline = fieldsline & " Quantity= '" & "1" & "'"                  '//Cantidad
            fieldsline = fieldsline & " AccountCode ='" & "_SYS00000002492" & "'"
            fieldsline = fieldsline & " U_LineNum='" & "0" & "'"

            ''fieldsline = fieldsline & " U_LineNum='" & "0" & "'"
            'fieldsline = fieldsline & " ItemCode= '" & "1146" & "'"             '//Codigo de Actividad, y en SAP es codigo del Articulo (OITM)
            'fieldsline = fieldsline & " U_Defecto= '" & "01" & "'"              '//Buscar Tabla de relacion.  DefectCode
            'fieldsline = fieldsline & " U_TipoOper= '" & "A" & "'"              '//Buscar Tabla de relacion.  OperationCode
            'fieldsline = fieldsline & " U_Ubicacion= '" & "GDI" & "'"           '//Buscar Tabla de relacion.  LocationCode
            'fieldsline = fieldsline & " U_RepairTime= '" & "60" & "'"
            'fieldsline = fieldsline & " U_ManoDeObra='" & 3.0 & "'"             '//Labour Cost
            'fieldsline = fieldsline & " U_CostoMateriales= '" & "100.00" & "'"  '//Material Cost
            'fieldsline = fieldsline & " Price='" & 30.0 & "'"                   '//Amount
            'fieldsline = fieldsline & " U_EstadoActividad= '" & "2" & "'"       '//activityStatusCode
            'fieldsline = fieldsline & " U_MotivoRechazo= '" & "" & "'"          '//RejectionReasonCode
            'fieldsline = fieldsline & " U_PrecioLimite= '" & "100.00" & "'"     '//Threshold

            ''//Control Interno de la DLL
            'fieldsline = fieldsline & " Quantity= '" & 1 & "'"                  '//Cantidad
            'fieldsline = fieldsline & " AccountCode ='" & "_SYS00000002493" & "'"
            'fieldsline = fieldsline & " U_LineNum='" & "1" & "'"


            ''fieldsline = fieldsline & " U_LineNum='" & "0" & "'"
            'fieldsline = fieldsline & " ItemCode= '" & "1146" & "'"             '//Codigo de Actividad, y en SAP es codigo del Articulo (OITM)
            'fieldsline = fieldsline & " U_Defecto= '" & "01" & "'"              '//Buscar Tabla de relacion.  DefectCode
            'fieldsline = fieldsline & " U_TipoOper= '" & "A" & "'"              '//Buscar Tabla de relacion.  OperationCode
            'fieldsline = fieldsline & " U_Ubicacion= '" & "GDI" & "'"           '//Buscar Tabla de relacion.  LocationCode
            'fieldsline = fieldsline & " U_RepairTime= '" & "60" & "'"
            'fieldsline = fieldsline & " U_ManoDeObra='" & 3.0 & "'"             '//Labour Cost
            'fieldsline = fieldsline & " U_CostoMateriales= '" & "100.00" & "'"  '//Material Cost
            'fieldsline = fieldsline & " Price='" & 30.0 & "'"                   '//Amount
            'fieldsline = fieldsline & " U_EstadoActividad= '" & "2" & "'"       '//activityStatusCode
            'fieldsline = fieldsline & " U_MotivoRechazo= '" & "" & "'"          '//RejectionReasonCode
            'fieldsline = fieldsline & " U_PrecioLimite= '" & "100.00" & "'"     '//Threshold

            ''//Control Interno de la DLL
            'fieldsline = fieldsline & " Quantity= '" & 1 & "'"                  '//Cantidad
            'fieldsline = fieldsline & " AccountCode ='" & "_SYS00000002493" & "'"
            'fieldsline = fieldsline & " U_LineNum='" & "2" & "'"

            'Luego de este insert se debe llamar la otra rutina para agregar otra Lines de Comentarios 
            'en la tabla [@]
            fields = fields & " Comments = '" & "Test The Record Modify damage" & "'"
            Dim CreateBy As String = "Gonzalo Morales"
            '//Para Post y Put y el LastSeenBy es el mimo valor que Modifiedby y CreateBy
            '// El NoOfTime se trabaja de forma y se actualiza en el Campo 
            '// U_Actualizaciones tomando como base la Post = 0 y Put se incrementa en 1

            myarray(0) = fields
            myarray_linea(0) = fieldsline

            MaxQuantity = 400
            MaxAmount = 280000.0
            MaxMonths = 0
            MaxOutstanding = 0
            PerMeleage = 0
            voException = False
            voDocEntry = 15703
            Dim AApprove = "Yes"
            Dim RequestComeFrom As String = "NORMAL"
            DaysInBetween = 0
            voSuccessSaved = False
            voWarningMsg = ""

            'Dim voAction = "Add"

            '"SMR"
            '"Damage"
            '"Replacement"

            'Mensaje = connx.SMRRequestAddrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMeleage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)
            Mensaje = connx.SMRRequestEditrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMeleage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
        Finally

            If DatosGenereral.SiGrabo > 1 Then
                'MsgBox(DatosGenereral.nMensaje & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
                'MsgBox(Mensaje)
            Else
                'If DatosGenereral.SiGrabo = 1 Then

                'MsgBox(Mensae & "RequesId " & voDocEntry & "Exception " & voException & "Success " & voSuccessSaved & "Warning " & voWarningMsg)

                'End If

                If connx.connected Then connx.disconnect()
            End If

            MsgBox(Mensaje)

            'MsgBox("RequesId " & voDocEntry)
            'MsgBox("Exception " & voException)
            'MsgBox("Success " & voSuccessSaved)
            'MsgBox("Warning " & voWarningMsg)
        End Try

    End Sub

    'Reemplazo
    Private Sub Reemplazo()

        Try
            connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion

            connx.transactional = False
            ''testConn()

            Dim nLicencia As String = "1001006805"
            Dim myarray(0) As String
            Dim fields As String = ""
            Dim voException As Boolean = False
            Dim docentryline As String = ""
            Dim MaxAmount As Double  '-Todo Tipo de Solicitud y por Supplier
            Dim MaxQuantity As Integer    '-Todo Tipo de Solicitud y por Supplier
            Dim MaxOutstanding As Integer '-Todo Tipo de Solicitud y por Supplier
            Dim MaxMonths As Integer '-Por Contrato
            Dim PerMeleage As Integer '-Por Contrato
            Dim myarray_linea(0) As String
            Dim fieldsline As String = ""
            Dim DaysInBetween As Integer = 0

            'U_Estado_PO
            '0001    Approved
            '0002    Initial
            '0003    New
            '0004    Partially Approved
            '0005    Planned
            '0006    Rejected
            '0007    Requested
            '0008    Rescheduled
            '0009    Cancelled

            '-- "P03" = Replacement
            '//Cabecera Replacement
            fields = " U_TypeofRequest= '" & "P03" & "'"                        '//RequestTypeCode
            fields = fields & " U_EstadoPO='" & "0002" & "'"                    '//RequestStatusCode
            Dim nContrato As String = "30342"                                   'ContractID
            Dim nPlaca As String = "P0556GRZ"                                   'LicensePlate
            fields = fields & " U_Kilometraje= '" & "90000" & "'"             '//Mileage
            fields = fields & " U_ContactName= '" & "GONZALO" & "'"             '//ContactName
            fields = fields & " U_FechaInicio='" & "2020-05-04" & "'"           '//StartDate
            fields = fields & " U_TelephoneNo= '" & "12345678" & "'"             '//TelephoneNo
            Dim dealerreference As String = "REF001"                              'dealerreference
            fields = fields & " U_CustomerContactNo= '" & "12345678" & "'"       '//CustomerContactNo
            fields = fields & " DocDate='" & Now & "'"                          '//CustomerRequestDate
            fields = fields & " U_FechaFinal='" & "2020-06-20" & "'"            '//SupplierProposedDate
            fields = fields & " U_ReplacementReason = '" & "RR1" & "'"          '//Replacement Reason 
            fields = fields & " CardCode='" & "P001186" & "'"                   '//SupplierId
            'total amount is calculated by SAP
            fields = fields & " DocCurrency = '" & "GTQ" & "'"                  '//Currency Code
            fields = fields & " U_InitialDuration='" & "1" & "'"            '//InitialDuration
            fields = fields & " U_ActualDuration= '" & "2" & "'"             '//ActualDuration
            fields = fields & " U_InitialRate= '" & "3" & "'"                '//InitialRate
            fields = fields & " U_ActualRate= '" & "4" & "'"                 '//ActualRate
            fields = fields & " U_InitialExtraDayCos='" & "5" & "'"         '//InitialExtraDayCost
            fields = fields & " U_ActualExtraDayCost= '" & "6" & "'"         '//ActualExtraDayCost

            'INSERTAR
            'Activity Code 9001
            fieldsline = fieldsline & " U_TipoAutoSust= '" & "RT01" & "'"        '//ReplacementTypeCode
            fieldsline = fieldsline & " U_DurationReasonCode = '" & "1" & "'"            '//DurationReasonCode
            fieldsline = fieldsline & " U_DurationActivityCo= '" & "2" & "'"             '//DurationActivityCode
            fieldsline = fieldsline & " U_RateReasonCode= '" & "3" & "'"                '//RateReasonCode
            fieldsline = fieldsline & " U_RateActivityCode= '" & "4" & "'"              '//RateActivityCode
            fieldsline = fieldsline & " U_ExtraDayReasonCode = '" & "5" & "'"         '//ExtraDayReasonCode
            fieldsline = fieldsline & " U_ExtraDayActivityCo= '" & "6" & "'"         '//ExtraDayActivityCode
            fieldsline = fieldsline & " U_MotivoRechazo= '" & "" & "'"          '//RejectionReasonCode
            fieldsline = fieldsline & " U_EstadoActividad= '" & "0001" & "'"          '//ActualStatusCode
            fieldsline = fieldsline & " Price='" & 30.0 & "'"  'hay que arreglar el campo del precio 

            'Luego de este insert se debe llamar la otra rutina para agregar otra Lines de Comentarios 
            'en la tabla [@]
            fields = fields & " Comments = '" & "Test The Record Modify Replacement" & "'"

            Dim CreateBy As String = "Gonzalo Morales"
            '//Para Post y Put y el LastSeenBy es el mimo valor que Modifiedby y CreateBy
            '// El NoOfTime se trabaja de forma y se actualiza en el Campo 
            '// U_Actualizaciones tomando como base la Post = 0 y Put se incrementa en 1


            '//Control Interno de la DLL
            ' ItemCode= '" & "9600" & "'"                   '//Codigo de Actividad, y en SAP es codigo del Articulo (OITM)
            'fieldsline = fieldsline & " Quantity= '" & 1 & "'"                  '//Cantidad
            'fieldsline = fieldsline & " AccountCode ='" & "_SYS00000002493" & "'"
            'fieldsline = fieldsline & " U_LineNum='" & "2" & "'"

            myarray(0) = fields
            myarray_linea(0) = fieldsline


            MaxQuantity = 400
            MaxAmount = 100000.0
            MaxMonths = 0
            MaxOutstanding = 0
            PerMeleage = 0
            voException = False
            voDocEntry = 15707
            Dim AApprove = "Yes"
            Dim RequestComeFrom As String = "NORMAL"
            DaysInBetween = 0
            voSuccessSaved = False
            voWarningMsg = ""
            'voAction  "Add"

            '"SMR"
            '"Damage"
            '"Replacement"

            Mensaje = connx.SMRRequestAddrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMeleage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)
            'Mensaje = connx.SMRRequestEditrequest(myarray, myarray_linea, nContrato, nPlaca, dealerreference, CreateBy, MaxQuantity, MaxAmount, MaxMonths, MaxOutstanding, PerMeleage, voException, voDocEntry, AApprove, RequestComeFrom, DaysInBetween, voSuccessSaved, voWarningMsg)

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
        Finally

            If DatosGenereral.SiGrabo > 1 Then
                '    'MsgBox(DatosGenereral.nMensaje & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
                '    MsgBox(Mensaje)
                'Else
                '    If DatosGenereral.SiGrabo = 1 Then
                '        MsgBox(Mensaje)
                '    End If

                If connx.connected Then connx.disconnect()
            End If

            MsgBox(Mensaje)

        End Try

    End Sub

    Private Sub ValidaClave()
        dTable = Datos.consulta_reader("Select * from dbo.tablasdll")

        If dTable.Rows.Count = 0 Then
            MsgBox("No Existen datos !!!!")
        Else

            'Setea Campos
            For Each dRow As DataRow In dTable.Rows
                MsgBox(dRow.Item("Id").ToString & " " & dRow.Item("Tabla").ToString & " " & dRow.Item("Campo").ToString & " " & dRow.Item("Tipo").ToString & " " & Convert.ToDouble(dRow.Item("Tamano").ToString) & " " & dRow.Item("cPut").ToString & " " & dRow.Item("cPost").ToString & " " & dRow.Item("pKey").ToString)
            Next

            'Me.Close()

        End If

    End Sub



    Private Sub Catalogos()
        'Dim connx2 As New SAPConnector("ARRENDDB\ARRENDDB: Milenia;sysadm,julio1;noimporta,noimporta")

        'Dim nErr As Long = 0
        'Dim errMsg As String = ""
        'Dim chk As Integer = 0

        'Dim uTable As SAPbobsCOM.UserTable = connx.getBObject("SMR_UBICACION")

        'uTable.Code = "FX"
        'uTable.Name = "Dato1"


        ''        oUserTable = SAPConnector.company.UserTables.Item("SMR_UBICACION");
        ''Int iRet = 0;

        ''Try
        ''{
        ''	oCompany.StartTransaction();
        ''	oUserTable.Code = code;
        ''    oUserTable.Name = Name;
        ''    oUserTable.UserFields.Fields.Item("U_Field").Value = Valor;
        ''	iRet = oUserTable.Add();



        'chk = uTable.Add()
        'If (chk = 0) Then
        '    DatosGenereral.SiGrabo = 1
        '    DatosGenereral.nMensaje = "The Catalog was successfully recorded !!!"
        'Else
        '    SAPConnector.company.GetLastError(nErr, errMsg)

        '    If (0 <> nErr) Then
        '        ' MsgBox("Error SAP:" + Str(nErr) + "," + errMsg)

        '        DatosGenereral.SiGrabo = 4
        '        DatosGenereral.nMensaje = "Error SAP:" + Str(nErr) + "," + errMsg

        '    End If
        'End If

        '        '1. To Update any UDO Data:
        '        Dim oGeneralService As SAPbobsCOM.GeneralService
        '        Dim oGeneralData As SAPbobsCOM.GeneralData
        '        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        '        Dim sCmp As SAPbobsCOM.CompanyService
        '        SAPbobsCOM.GeneralData oChild = null;
        'SAPbobsCOM.GeneralDataCollection oChildren = null;
        'sCmp = SBO_Company.GetCompanyService();
        'oGeneralService = sCmp.GetGeneralService("UDOCODE");
        'oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));
        'oGeneralParams.SetProperty("Code", ANYCode);
        'oGeneralData = oGeneralService.GetByParams(oGeneralParams);
        'oChildren = oGeneralData.Child("DETAILTABLEOFUDO");
        'oChild = oChildren.Item(LineID - 1);
        'oChild.SetProperty("U_Field1", VALUETOSET1);
        'oChild.SetProperty("U_Field2", VALUETOSET2);

        'oGeneralService.Update(oGeneralData);

        '        '1. To Add any UDO Data:
        '        Dim oGeneralService As SAPbobsCOM.GeneralService
        '        Dim oGeneralData As SAPbobsCOM.GeneralData
        '        Dim oChild As SAPbobsCOM.GeneralData
        '        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        '        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams

        '        oCompService = SAPConnector.company.GetCompanyService()

        '        SBO_Company.StartTransaction();
        'oGeneralService = oCompService.GetGeneralService("UDOCODE");
        'oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
        'oGeneralData.SetProperty("Code", ANYCode);
        'oGeneralData.SetProperty("U_Field1", ProjectCode);
        '// Adding data to Detail Line
        'oChildren = oGeneralData.Child("ACTDETAIL1");
        'oChild = oChildren.Add();
        'oChild.SetProperty("U_Code", VALUE);
        'oChild.SetProperty("U_Name", VALUE);
        'oGeneralService.Add(oGeneralData);
        'Hope it helps.
        'Thanks & Regards
        'Ankit Chauhan

    End Sub
    Private Sub AsignVeh()
        connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion
        connx.transactional = False

        Dim Customer As String = "C002723"
        Dim License As String = "24681357"
        Dim Plate As String = "P0835HGL"

        Try
            Mensaje = connx.AsignVehicle(Customer, License, Plate)
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
        MsgBox(Mensaje)
    End Sub

    Private Sub Drivers()

        connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion
        connx.transactional = False

        Dim SiGrabo As Integer = 0
        Dim nCliente As String = "C002723"
        Dim nContrato As String = ""
        Dim nLicencia As String = "24681357"
        Dim nPlaca As String = ""

        'Pilotos
        Dim myarray(0) As String
        Dim fields As String = ""

        fields = "Name='" & "JUAN BURRION" & " '"
        fields = fields & " FirstName= '" & "JUAN MANUEL" & "'"
        fields = fields & " LastName= '" & "BURRION FLORES" & "'"
        fields = fields & " Gender= '" & "M" & "'"
        fields = fields & " DateOfBirth= '" & "1988-11-15" & "'"
        fields = fields & " Address= '" & "NUEVA MONTSERRAT ZONA 7" & "'"
        fields = fields & " Phone1= '" & "22181818" & "'"
        fields = fields & " Phone2= '" & "22181830" & "'"
        fields = fields & " MobilePhone= '" & "52003011" & "'"
        fields = fields & " E_Mail= '" & "jburrion@gmail.com" & "'"
        fields = fields & " U_CreatedBy= '" & "Gonzalo Morales" & "'"
        fields = fields & " U_ModifiedBy= '" & "Gonzalo Morales" & "'"
        fields = fields & " Active= '" & "N" & "'"

        myarray(0) = fields

        Try
            'Mensaje = connx.AddDriver(myarray, nCliente, nLicencia, nContrato, nPlaca)    'Driver Agregar
            Mensaje = connx.EditDriver(myarray, nCliente, nLicencia, nContrato, nPlaca)    'Driver Agregar

            SiGrabo = 1
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
        End Try

        If SiGrabo = 1 Then

            If connx.connected Then connx.disconnect()
        End If

        MsgBox(Mensaje)

    End Sub

    Private Sub FuelCard()

        connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion
        connx.transactional = False

        Dim nContrato As String = "30544"
        Dim nLicencia As String = "P0228HJT"


        'CardFuel
        Dim myarray(0) As String
        Dim fields As String = ""
        Dim SiGrabo As Integer = 0

        fields = fields & " Code= '" & "x" & "'"
        fields = fields & " Name= '" & "x" & "'"
        fields = fields & " U_ContractID= '" & nContrato & "'"
        fields = fields & " U_International= '" & "1" & "'"
        fields = fields & " U_CardType= '" & "I" & "'"
        fields = fields & " U_Suppliers= '" & "S" & "'"
        fields = fields & " U_Brand= '" & "T" & "'"
        fields = fields & " U_FuelCardNumber= '" & "12345678" & "'"
        fields = fields & " U_LicensePlate= '" & "P0579GCX" & "'"
        fields = fields & " U_DriverLicenseNo= '" & "21204568048" & "'"
        fields = fields & " U_CreatedBy= '" & "gonzalo.morales@arrendleasing.com" & "'"
        fields = fields & " U_Reason= 'R'"
        fields = fields & " U_LimitPerCard= '1000'"
        fields = fields & " U_Currency= 'QTZ'"
        fields = fields & " U_UsageFrecuency= 'Daily'"
        fields = fields & " U_Comments= 'Comentarios'"
        fields = fields & " U_CreateDate= '2020-03-26'"

        myarray(0) = fields

        Try
            Mensaje = connx.AddFuelCard(myarray, nContrato, nLicencia)    'Driver Agregar
            SiGrabo = 1
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
        End Try

        If DatosGenereral.SiGrabo > 1 Then
            '    'MsgBox(DatosGenereral.nMensaje & vbCrLf & "Por Favor consultar con el Administrador del Sistema")
            '    MsgBox(Mensaje)
            'Else
            '    If DatosGenereral.SiGrabo = 1 Then
            '        MsgBox(Mensaje)
            '    End If

            If connx.connected Then connx.disconnect()
        End If

        MsgBox(Mensaje)


    End Sub

    'Private Sub importaTasas()
    '    Dim connx2 As New SAPConnector("ARRENDDB\ARRENDDB: Milenia;sysadm,julio1;noimporta,noimporta")
    '    Dim dtable As DataTable = connx2.execQuery(<sql>
    '                         select codigo_moneda	U_CodigoMoneda
    '                         , fecha_hora		U_FechaHora
    '                         , tasa_compra		U_TasaCompra
    '                         , tasa_venta		U_TasaVenta
    '                         , num_empresa	U_NumEmpresa
    '                        from finasql.caja_tasa_cambio
    '                        where fecha_hora between '2015-05-01' and '2015-05-31'
    '                     </sql>.Value).Table

    '    Dim uTable As SAPbobsCOM.UserTable = connx.getBObject("MCAJA_TASA_CAMBIO")
    '    For Each row In dtable.Rows
    '        connx.insert(uTable, {""}, True)
    '        For column As Integer = 0 To 4
    '            uTable.UserFields.Fields.Item(dtable.Columns(column).ColumnName).Value = row.item(column)
    '        Next column
    '        connx.commit()
    '    Next row
    'End Sub

    ''Private Sub llenaFechas()
    ''    Dim failed As Integer = 0
    ''    For Each linea In File.ReadLines("C:\Users\bleiva.LEASING\Documents\codes.txt")
    ''        Try
    ''            connx.updateOnKey(connx.getBObject("MCUOTA_CONTRATO"), {"U_FechaCalcMora=' '"}, linea)
    ''        Catch ex As Exception
    ''            failed += 1
    ''            Continue For
    ''        End Try
    ''    Next linea
    ''    MsgBox(failed)
    ''End Sub

    Sub testConn()
        connx.connect()
        MsgBox(connx.connected)
        connx.disconnect()
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Me.Close()
    'End Sub

    Private Sub ServiceCallActivities()

        Dim sboServiceCall As SAPbobsCOM.ServiceCalls
        Dim sboContact As SAPbobsCOM.Contacts
        Dim lngDocEntry As Long
        Dim lngKey As String

        'oServiceCall = connx.getBObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        sboServiceCall = connx.getBObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        sboContact = connx.getBObject(SAPbobsCOM.BoObjectTypes.oContacts)

        'Crea la Sevice Call
        Dim retval As Integer = 0

        Try
            sboServiceCall.CustomerCode = "C002756"        'CardCode
            sboServiceCall.Subject = " Created : " & Now  'Asunto
            sboServiceCall.InternalSerialNum = 32

            retval = sboServiceCall.Add()
            If retval <> 0 Then
                MsgBox(SAPConnector.company.GetLastErrorDescription)
            Else
                MsgBox("Grabo Service Call " & SAPConnector.company.GetNewObjectKey)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            'GC.Collect()
        End Try

        'Servicio
        lngKey = SAPConnector.company.GetNewObjectKey

        'Create a new Activity
        sboContact.CardCode = "P001186"
        sboContact.Closed = 0
        sboContact.ContactDate = Now
        'sboContact.Notes = "Esta es mi Prueba Exitosa Sha, Agregando Ativities Service Call"
        sboContact.Details = "Details...."
        sboContact.Notes = "Esta es mi Prueba Exitosa Sha, Agregando Ativities Service Call"

        'sboContact.DocType = 22            'Para Agregar Pedidos
        'sboContact.DocEntry = lngKey       'Enlaza El pedido con Orden de Compra

        retval = sboContact.Add()
        If retval <> 0 Then
            MsgBox(SAPConnector.company.GetLastErrorDescription)
        Else
            MsgBox("Agrego la Actividad " & SAPConnector.company.GetNewObjectKey)
        End If

        'Get Activity Code
        lngDocEntry = SAPConnector.company.GetNewObjectKey
        Try
            'Assign Activity to Service Call
            'Set sboServiceCall = sboCompany.GetBusinessObject(oServiceCalls)
            sboServiceCall.GetByKey(lngKey)

            sboServiceCall.Subject = "Prueba Actividad - SHA" & Now
            sboServiceCall.Activities.ActivityCode = lngDocEntry
            sboServiceCall.Activities.Add()


            retval = sboServiceCall.Update()
            If retval <> 0 Then
                MsgBox(SAPConnector.company.GetLastErrorDescription)
            Else
                MsgBox("Modifico la Service Call " & SAPConnector.company.GetNewObjectKey)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            GC.Collect()
        End Try


    End Sub

    Private Sub m_UpdatePurcherOrder()

        'Dim oAtt As SAPbobsCOM.Attachments2
        'oAtt = Con.GetBusinessObject(BoObjectTypes.oAttachments2)

        'oAtt.Lines.Add()
        'Dim fileOK As Boolean = False 'Creo una variable de tipo boolean 
        'Dim path As String = "C:/Anexos" 'Debes crear una carpeta vacia "Anexos" con un archivo en la solucion de tu proyecto .Net 
        'oAtt.Lines.SourcePath = path 'Aqui le doy la ruta de origen, que exige el objeto de la DI-API Para poder hacer el Attachment
        'Dim FileName = "MiArchivo.PDF"

        'Dim fileExtension As String 'Creo la Variable FileExtencion
        'fileExtension = System.IO.Path.GetExtension(FileName).ToLower() 'Obtengo la extencion de mi archivo y la convierto a Minuscula
        'Dim allowedExtensions As String() = {".jpg", ".jpeg", ".png", ".gif", ".pdf"} 'Extenciones validas 
        '    'Recorro las extenciones y las comparo con las de mi archivo
        '    For i As Integer = 0 To allowedExtensions.Length - 1
        '        If fileExtension = allowedExtensions(i) Then
        '            fileOK = True
        '        End If
        '    Next
        ''si fue correcta la extencion entonces poga la variable FileOk en verdadero
        'If fileOK Then
        '    Try
        '        oAtt.Lines.FileName = FileName 'Entreguele al objeto de la Di-api el Nombre del Archivo
        '        oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES 'Sobre escriba el metodo

        '        '//ERROR Necesito tener varios Adjuntos y no permite
        '        oAtt.Lines.Add()
        '        oAtt.Lines.FileName = FileName
        '        oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES 'Sobre escriba el metodo

        '        oAtt.Lines.Add()
        '        oAtt.Lines.FileName = FileName
        '        oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES 'Sobre escriba el metodo
        '    Catch ex As Exception
        '        MsgBox("A ocurrido un error: " & ex.Message)
        '    End Try
        'Else
        '    MsgBox("El Sistema no acepta este tipo de archivos")
        'End If

        'Dim iAttEntry As Integer = -1
        'Dim ErrAtt As Integer = oAtt.Add
        'Dim ErrCode As String = Con.GetLastErrorCode
        'Dim ErrDescrip As String = Con.GetLastErrorDescription

        ''If (oAtt.Update() = 0) '// Para Modificar

        'If (oAtt.Add() = 0) Then

        '    'So, you need get the id of the attachment that was created 
        '    iAttEntry = Int32.Parse(Con.GetNewObjectKey())   '//ERROR Necesito Saber cual fue el ultimo Numero de Attarchmen y no permite


        '    '-- Lo solucione de otra manera

        '    'Dim rSet As SAPbobsCOM.Recordset = connx.getRecordSet("Select top 1 AbsEntry from [OATC] order by AbsEntry desc")

        '    'If rSet.RecordCount > 0 Then
        '    '    rSet.MoveFirst()
        '    '    iAttEntry = rSet.Fields.Item(0).Value
        '    'End If
        '    ''


        '    'Adjunta el Attacmen a la Orden de Compra
        '    Dim RetVal As Integer = 0
        '    Dim vOrder As SAPbobsCOM.Documents
        '    vOrder = Con.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
        '    Dim nDocEntry As Integer = 220

        '    If (vOrder.GetByKey(nDocEntry)) Then
        '        'test = oContact.CardCode
        '        vOrder.AttachmentEntry = iAttEntry
        '        'vOrder.Attachments.Add
        '        RetVal = vOrder.Update()

        '        If RetVal <> 0 Then
        '            MsgBox(SAPConnector.company.GetLastErrorDescription)
        '        Else
        '            MsgBox("Se Actulizao Orden de Compra ")
        '        End If
        '    End If

        'Else
        '    MsgBox("A ocurrido un error: " & ErrCode & ":" & ErrDescrip)
        'End If

    End Sub

    Private Sub OrdenCompra()

        connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion
        connx.transactional = False
        connx.connect()


        ''--Agregar Lines
        Dim RetVal As Integer
        Dim vOrder As SAPbobsCOM.Documents
        vOrder = connx.getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
        Dim fec As DateTime = DateTime.Now

        'If vOrder.GetByKey("15257") Then

        vOrder.CardCode = "P001186"
        vOrder.DocDate = Now
        vOrder.DocDueDate = Now
        vOrder.NumAtCard = "PRUEBA"
        vOrder.Comments = "PRUEBA DE OPCION DE COMPRA SHA"
        vOrder.UserFields.Fields.Item("U_FechaCliente").Value = fec

        'vOrder.Lines.SetCurrentLine("0")
        'vOrder.Lines.ItemCode = "1145"
        'vOrder.Lines.Quantity = "1"
        'vOrder.Lines.Price = "15.00"
        'vOrder.Lines.UserFields.Fields.Item("U_Descripcionii").Value = "Prueba10"
        'vOrder.Lines.Delete()

        vOrder.Lines.ItemCode = "1146"
        vOrder.Lines.Quantity = "1"
        vOrder.Lines.Price = "10.00"
        vOrder.Lines.UserFields.Fields.Item("U_Descripcionii").Value = "Prueba13"
        vOrder.Lines.UserFields.Fields.Item("U_Descripcionii").Value = "Prueba13"

        ''vOrder.Lines.SetCurrentLine("1")
        'vOrder.Lines.ItemCode = "1148"
        'vOrder.Lines.Quantity = "5"
        'vOrder.Lines.Price = "10.5"
        'vOrder.Lines.UserFields.Fields.Item("U_Descripcionii").Value = "Prueba3"
        ''vOrder.Lines.Delete()
        ''vOrder.Lines.Add()

        ''vOrder.Lines.SetCurrentLine("2")
        'vOrder.Lines.ItemCode = "1152"
        'vOrder.Lines.Quantity = "8"
        'vOrder.Lines.Price = "10.5"
        'vOrder.Lines.UserFields.Fields.Item("U_Descripcionii").Value = "Prueba2"
        ''vOrder.Lines.Delete()
        ''vOrder.Lines.Add()

        'vOrder.Series = 93
        'La linea siguente es muy importante ya que le indicas el tipo de docuemnto que se va a crear, en este caso puse de ejemplo un Orden de Compra
        'vOrder.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
        'RetVal = vOrder.Update

        Dim nErr As Long = 0
            Dim errMsg As String = ""

            Try

            'SAPConnector.company.GetLastError(nErr, errMsg)
            RetVal = vOrder.Add()

            'RetVal = vOrder.Update()

            SAPConnector.company.GetLastError(nErr, errMsg)

                If RetVal <> 0 Then
                    MsgBox("Ocurrio un error " & MsgBox("Error SAP:" + Str(nErr) + "," + errMsg))
                    ' Exit Sub
                Else
                    MsgBox("Grabao con Existo 1 ")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        'End If


        '        docPO = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

        'If (docPO.GetByKey(System.Convert.ToInt32(ss.docEntry))) Then
        '            {
        '// Add a New PO line that makes up the difference in what was shipped And the original quantity
        'docPO.Lines.ItemCode = ssi.itemCode;
        'docPO.Lines.Quantity = System.Convert.ToDouble(poLineRow["OpenQty"].ToString()) - draft.Lines.Quantity;
        'docPO.Lines.WarehouseCode = ssi.whsCode;
        'docPO.Lines.Price = ssi.price;
        'docPO.Lines.Add();
        '// Now modify the existing, "original" PO line item to reflect the quantity change (this will show what actually shipped)
        'docPO.Lines.SetCurrentLine(System.Convert.ToInt32(poLineRow["LineNum"].ToString()));
        'docPO.Lines.Quantity = draft.Lines.Quantity;
        'Result = docPO.Update();
        'If (Result!= 0) Then
        '                {
        'emailSupport("An SAP error at draft.Add() for Purchase Document " + tempDocEntry + "\n\n" + company.GetLastErrorDescription(), tempCardCode);
        'Return;
        '}
        '}
        connx.disconnect()

    End Sub

    Private Sub Servicio()
        Dim nErr As Long = 0
        Dim errMsg As String = ""
        Dim chk As Integer = 0
        'Dim connx As New SAPConnector

        Dim nCliente As String = "C000005"
        Dim nPiloto As String = "COO5-P4"


        'Create the BusinessPartners object
        Dim oServiceCall As SAPbobsCOM.ServiceCalls
        Dim act As SAPbobsCOM.Contacts


        'vBP = connx.GetBusinessObject(oBusinessPartners) 'Calls BusinessPartners object

        'oServiceCall = connx.getBObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        oServiceCall = connx.getBObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        act = connx.getBObject(SAPbobsCOM.BoObjectTypes.oContacts)

        Dim retval As Integer = 0

        Try
            oServiceCall.CustomerCode = "C002756"        'CardCode
            oServiceCall.Subject = " Created : " & Now  'Asunto
            oServiceCall.InternalSerialNum = 32

            retval = oServiceCall.Add()
            If retval <> 0 Then
                MsgBox(SAPConnector.company.GetLastErrorDescription)
            Else
                MsgBox(SAPConnector.company.GetNewObjectKey)
            End If



            'act.CardCode = "P001186"
            'act.Details = "API TEST FROM C#"
            'act.Notes = "Test from C#"

            'retval = act.Add()
            'If retval <> 0 Then
            '    MsgBox(SAPConnector.company.GetLastErrorDescription)
            'Else
            '    MsgBox(SAPConnector.company.GetNewObjectKey)
            'End If

            ''Dim clgCode As Integer = SAPConnector.company.GetNewObjectKey

            'oServiceCall.GetByKey("195")
            'oServiceCall.Activities.ActivityCode = 195
            'oServiceCall.Activities.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            GC.Collect()
        End Try



        'If (chk = 0) Then
        '        MsgBox("Exitoso")
        '    Else

        '        SAPConnector.company.GetLastError(nErr, errMsg)
        '        If (0 <> nErr) Then
        '            MsgBox("Found error:" + Str(nErr) + "," + errMsg)
        '        End If
        '    End If
        '
    End Sub


    'Sub testPC()
    '    Dim pc As SAPbobsCOM.JournalEntries = connx.getBObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    'End Sub

    ''Private Sub insertaPC()
    ''    Dim msg As String = "", rtn As Long = 0
    ''    Dim pc As SAPbobsCOM.JournalEntries = connx.getBObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    ''    pc.Lines.ShortName = "C003262"
    ''    'pc.Lines.ControlAccount = "_SYS00000001962"
    ''    pc.Lines.Debit = 1000

    ''    pc.Lines.Add()

    ''    pc.Lines.AccountCode = "_SYS00000001962"
    ''    pc.Lines.Credit = 1000

    ''    pc.Lines.Add()

    ''    connx.insert(pc, {""})
    ''End Sub

    ''Private Sub actualizaMail()
    ''    For Each row As DataRow In connx.execQuery("SELECT CardCode From OCRD WHERE E_Mail <> 'migracionsap@leasing.com.gt' AND CardType = 'C'", _
    ''                                               New Dictionary(Of String, Object)).Table.Rows
    ''        connx.updateOnKey(connx.getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), {""}, row(0), True)
    ''        connx.OnHold.EmailAddress = "migracionsap@leasing.com.gt"
    ''        connx.commit()
    ''    Next
    ''End Sub

    ''Private Sub probar_mora()
    ''    Dim calculador As New CalculadorMora
    ''    Dim cadena As String = New RegConnector("Software\MyDB").tuConexion

    ''    cadena = cadena.Substring(0, cadena.LastIndexOf(";"))

    ''    With calculador
    ''        .connString = cadena
    ''        MsgBox(.mora(4, 16596, 15, "", "N"))
    ''    End With
    ''End Sub

    Private Sub cambiarEM()
        connx.ConnString = New ArrendUtils_AF.RegConnector("ConfigSQL.xml").tuconexion

        connx.transactional = False
        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0

        Dim mail As String = ""

        Dim sn As SAPbobsCOM.BusinessPartners = connx.getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        Dim rSet As SAPbobsCOM.Recordset = connx.getRecordSet("Select * from OCRD WHERE E_Mail = 'marisol.contreras@arrendleasing.com' and CardType = 'S'")

        If rSet.RecordCount = 0 Then Exit Sub Else rSet.MoveFirst()
        While Not rSet.EoF


            If sn.GetByKey(rSet.Fields.Item(0).Value) Then

                sn.EmailAddress = mail

                sn.Update()
                If (chk = 0) Then

                    'MsgBox("Exitoso")
                Else
                    SAPConnector.company.GetLastError(nErr, errMsg)
                    If (0 <> nErr) Then
                        MsgBox("Found error:" + Str(nErr) + "," + errMsg)
                    End If
                End If

            End If


            rSet.MoveNext()
        End While

        MsgBox("Termino Exitosomente !!")
    End Sub

    'Private Sub inserta_cheque()

    '    Dim connx As New ArrendUtils.SAPConnector

    '    With connx

    '        .transactional = False
    '        .ConnString = New RegConnector("Software\MyDB").tuConexion

    '        Dim pago As SAPbobsCOM.Payments = .getBObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)



    '        Dim contenidoChk() As String = {<keys
    '                                            CheckSum='100' DueDate='2015-05-18' CheckAccount='_SYS00000001260'

    '                                        />.ToString.Replace("""", "'")}
    '        Dim contenidoRCT() As String = {<keys
    '                                            CardCode='C002150' DocDate='2015-05-18' DueDate='2015-05-18'

    '                                        />.ToString.Replace("""", "'")}

    '        .insert(pago.Checks, contenidoChk)
    '        .insert(pago, contenidoRCT)

    '        .disconnect()
    '    End With

    'End Sub

    'Private Sub inserta_NC()
    '    Dim msg As String = "", rtn As Long = 0
    '    'Dim connx As New SAPConnector
    '    With connx
    '        .transactional = False
    '        .ConnString = New RegConnector("Software\MyDB").tuConexion
    '        Dim fac As SAPbobsCOM.Documents = .getBObject(SAPbobsCOM.BoObjectTypes.oInvoices)
    '        Dim nc As SAPbobsCOM.Documents = .getBObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)

    '        fac.GetByKey(152953)


    '        For i = 0 To fac.Lines.Count - 1
    '            fac.Lines.SetCurrentLine(i)

    '            nc.Lines.BaseType = fac.Lines.LineType
    '            nc.Lines.BaseEntry = fac.DocEntry
    '            nc.Lines.BaseLine = fac.Lines.LineNum
    '            nc.Lines.ItemCode = fac.Lines.ItemCode
    '            nc.Lines.Quantity = fac.Lines.Quantity
    '            nc.Lines.TaxCode = fac.Lines.TaxCode
    '            nc.Lines.UnitPrice = fac.Lines.UnitPrice

    '            nc.Lines.Add()

    '            'connx.Company.GetLastError(rtn, msg)

    '            If rtn <> 0 Then MsgBox(msg)
    '        Next i


    '        nc.DocDate = fac.DocDate
    '        nc.DocDueDate = fac.DocDueDate
    '        nc.TaxDate = fac.TaxDate
    '        nc.DiscountPercent = fac.DiscountPercent
    '        nc.CardCode = fac.CardCode
    '        nc.DocTotal = fac.DocTotal

    '        nc.Add()

    '        'connx.Company.GetLastError(rtn, msg)

    '        If rtn <> 0 Then MsgBox(msg)


    '        '.insert(nc, {""})
    '        .disconnect()
    '    End With
    'End Sub

    'Private Sub inserta_recibo()
    '    Dim connx As New ArrendUtils.SAPConnector
    '    With connx
    '        .ConnString = New RegConnector("Software\MyDB").tuConexion

    '        Dim rec As SAPbobsCOM.Payments = .getBObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments) 'Recibo

    '        'Agregarle un Checke
    '        Dim contCheque() As String = {<keys


    '                                      />.ToString.Replace("""", "'")}

    '        'Agregar la cabecera
    '        'Insertarlo

    '    End With
    'End Sub

    'Private Sub testReader()
    '    Dim table As DataTable = connx.execQuery("SELECT Top 7 DocEntry FROM OINV", New Dictionary(Of String, Object)).Table
    '    For Each row In table.Rows
    '        MsgBox(row(0))
    '    Next
    'End Sub

    Private Sub Piloto()
        Dim nErr As Long
        Dim errMsg As String = ""
        Dim chk As Integer = 0
        'Dim connx As New SAPConnector

        Dim nCliente As String = "C005004"
        Dim nPiloto As String = "SAUL HERNANDEZ"

        Dim DatosIngresados(13) As String
        Dim fecha As DateTime = Now()

        DatosIngresados(0) = nCliente
        DatosIngresados(1) = nPiloto
        DatosIngresados(2) = "SAUL"
        DatosIngresados(3) = "HERNANDEZ"
        DatosIngresados(4) = 1
        DatosIngresados(5) = fecha
        DatosIngresados(6) = "010101"  'Codigo Postal
        DatosIngresados(7) = "Direccion Correcta"
        DatosIngresados(8) = "Guatemala"
        DatosIngresados(9) = "Ciudad"
        DatosIngresados(10) = "Villa Nueva"
        DatosIngresados(11) = "12313"
        DatosIngresados(12) = "59459446"
        DatosIngresados(13) = "V001645"


        'Create the BusinessPartners object
        Dim vBP As SAPbobsCOM.BusinessPartners
        Dim sboContacts As SAPbobsCOM.ContactEmployees
        'vBP = connx.GetBusinessObject(oBusinessPartners) 'Calls BusinessPartners object

        vBP = connx.getBObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        sboContacts = vBP.ContactEmployees

        If vBP.GetByKey(nCliente) Then

            If (vBP.ContactEmployees.Count = 0) Then
                vBP.ContactEmployees.Add()
            Else
                vBP.ContactEmployees.Add()
                'lse
                '    If (Not vBP.ContactEmployees.Name = "") Then

                '        vBP.ContactEmployees.Name = "SAUL"
                '        vBP.ContactEmployees.Add()

                '        chk = vBP.Update()
                '    End If
            End If

            vBP.ContactEmployees.Name = DatosIngresados(1)
                vBP.ContactEmployees.FirstName = DatosIngresados(2)
                vBP.ContactEmployees.LastName = DatosIngresados(3)
                vBP.ContactEmployees.Gender = DatosIngresados(4)
                vBP.ContactEmployees.DateOfBirth = DatosIngresados(5)
                vBP.ContactEmployees.UserFields.Fields.Item("U_LugarEmision").Value = DatosIngresados(6)
                vBP.ContactEmployees.Address = DatosIngresados(7)
                vBP.ContactEmployees.UserFields.Fields.Item("U_PaisDireccion").Value = DatosIngresados(8)
                vBP.ContactEmployees.UserFields.Fields.Item("U_DeptoDir").Value = DatosIngresados(9)
                vBP.ContactEmployees.UserFields.Fields.Item("U_MuniDir").Value = DatosIngresados(10)
                vBP.ContactEmployees.Phone1 = DatosIngresados(11)
                vBP.ContactEmployees.Phone1 = DatosIngresados(12)
                vBP.ContactEmployees.UserFields.Fields.Item("U_LugarTrabajo").Value = DatosIngresados(13)
                chk = vBP.Update()

                If (chk = 0) Then
                    MsgBox("Exitoso")
                Else

                    SAPConnector.company.GetLastError(nErr, errMsg)
                    If (0 <> nErr) Then
                        MsgBox("Found error:" + Str(nErr) + "," + errMsg)
                    End If
                End If
            End If

        'vBP.ContactEmployees.Name = DatosIngresados(1)
        'vBP.ContactEmployees.FirstName = DatosIngresados(2)
        'vBP.ContactEmployees.LastName = DatosIngresados(3)
        'vBP.ContactEmployees.Gender = DatosIngresados(4)
        'vBP.ContactEmployees.DateOfBirth = DatosIngresados(5)
        'vBP.ContactEmployees.UserFields.Fields.Item("U_LugarEmision").Value = DatosIngresados(6)
        'vBP.ContactEmployees.Address = DatosIngresados(7)
        'vBP.ContactEmployees.UserFields.Fields.Item("U_PaisDireccion").Value = DatosIngresados(8)
        'vBP.ContactEmployees.UserFields.Fields.Item("U_DeptoDir").Value = DatosIngresados(9)
        'vBP.ContactEmployees.UserFields.Fields.Item("U_MuniDir").Value = DatosIngresados(10)
        'vBP.ContactEmployees.Phone1 = DatosIngresados(11)
        'vBP.ContactEmployees.Phone1 = DatosIngresados(12)
        'vBP.ContactEmployees.UserFields.Fields.Item("U_LugarTrabajo").Value = DatosIngresados(13)

        'chk = vBP.Add() '// used BP.Add();
        'If (chk = 0) Then

        '    MsgBox("Exitoso")
        'Else

        '    SAPConnector.company.GetLastError(nErr, errMsg)
        '    If (0 <> nErr) Then
        '        MsgBox("Found error:" + Str(nErr) + "," + errMsg)
        '    End If
        'End If


    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'Form1
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Name = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Private Sub adjuntar2()

        Try
            Dim oB As SAPbobsCOM.Documents
            'oB = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders), SAPbobsCOM.Documents)
            oB = connx.getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

            ''oB.Series = 15
            'oB.DocDate = "2019-11-13"
            'oB.CardCode = "P000971"
            'oB.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items

            'oB.Lines.ItemCode = "1147"
            'oB.Lines.Quantity = 1
            'oB.Lines.TaxCode = "IVA"
            'oB.Lines.LineTotal = "20.00"

            'ACA AGREGO LOS ANEXOS
            'En este caso se agrega manualmente los archivos, sin embargo puedes hacer un for o while si quieres hacerlo de alguna manera mas automatica.

            Dim oATT As SAPbobsCOM.Attachments2
            oATT = connx.getBObject(SAPbobsCOM.BoObjectTypes.oAttachments2)


            If oATT.GetByKey(4988) Then

                'Dim rSetII As SAPbobsCOM.Recordset = connx.getRecordSet("Select Line from [ATC1]  where AbsEntry = 4988")

                'If rSetII.RecordCount > 0 Then
                ''rSetII.MoveFirst()

                'While Not rSetII.EoF
                'Dim nLin2 As String = rSetII.Fields.Item(0).Value

                oATT.Lines.SetCurrentLine("0")      '//Modifica

                oATT.Lines.FileName = "NCE-64-NCA-001_190000000016" 'Nombre del documento
                oATT.Lines.FileExtension = "pdf" 'Extension del archivo
                        oATT.Lines.SourcePath = "C:\tmp" 'Ruta del archivo
                        oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES 'Sobre escribir o no

                'rSetII.MoveNext()

                'Dim nLin3 As String = rSetII.Fields.Item(0).Value
                oATT.Lines.SetCurrentLine("1")      '//Modifica

                oATT.Lines.FileName = "NCE-64-NCA-001_190000000017"
                oATT.Lines.FileExtension = "pdf"
                    oATT.Lines.SourcePath = "C:\tmp"
                    oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES

                'End While
                'End If

                ''oATT.Lines.SetCurrentLine("2")
                ''oATT.Lines.Add()
                ''oATT.Lines.FileName = "363674_99B8969F_3220130776" 'Nombre del documento
                'oATT.Lines.FileName = " NCE-64-NCA-001_190000000016" 'Nombre del documento
                'oATT.Lines.FileExtension = "pdf" 'Extension del archivo
                'oATT.Lines.SourcePath = "C:\tmp" 'Ruta del archivo
                'oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES 'Sobre escribir o no

                ''oATT.Lines.Add()
                'oATT.Lines.FileName = "363675_F6535FE4_3423683311"
                'oATT.Lines.FileExtension = "pdf"
                'oATT.Lines.SourcePath = "C:\tmp"
                'oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES

                'oATT.Lines.Add()
                'oATT.Lines.FileName = "FAC_82_548290"
                'oATT.Lines.FileExtension = "xml"
                'oATT.Lines.SourcePath = "C:/temp"
                'oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES
                ''Finaliza el agregar anexos



                'If oATT.Add() <> 0 Then
                If oATT.Update() <> 0 Then
                    'Dim Respuesta = SAPConnector.company.GetLastErrorCode 'Esto si es incorrecto
                    Dim Respuesta = SAPConnector.company.GetLastErrorDescription 'Esto si es incorrecto
                    MsgBox(Respuesta)
                Else
                    ' oB.AttachmentEntry = SAPConnector.company.GetNewObjectKey 'Este es el entry que genera
                    MsgBox("Actualizado")
                End If
            End If

            'If oB.Add() <> 0 Then
            '    Dim Respuesta = SAPConnector.company.GetLastErrorDescription 'Esto si es incorrecto
            '    MsgBox(Respuesta)
            'Else
            '    Dim Respuesta = SAPConnector.company.GetNewObjectKey()
            '    MsgBox(Respuesta)
            'End If
        Catch ex As Exception
            Dim respuesta = ex.ToString
        End Try

        'Dim oAtt As SAPbobsCOM.Attachments2
        'oAtt = connx.getBObject(SAPbobsCOM.BoObjectTypes.oAttachments2)
        'oAtt.Lines.Add()

        'Dim path As String = "C:\tmp" 'Debes crear una carpeta vacia "Anexos" con un archivo en la solucion de tu proyecto .Net
        'oAtt.Lines.SourcePath = path   'Aqui le doy la ruta de origen, que exige el objeto de la DI-API Para poder hacer el Attachment

        ''If FileCargarArchivosc.HasFile Then
        'Dim fileName As String = "363674_99B8969F_3220130776.PDF"
        'Dim fileName2 As String = "363675_F6535FE4_3423683311.PDF"
        'Dim fileExtension As String 'Creo la Variable FileExtencion
        'Dim fileOK As Boolean = False 'Creo una variable de tipo boolean 
        'Dim nDocEntry As Integer = 15040

        'fileExtension = System.IO.Path.GetExtension(fileName).ToLower() 'Obtengo la extencion de mi archivo y la convierto a Minuscula
        'Dim allowedExtensions As String() = {".jpg", ".jpeg", ".png", ".gif", ".pdf"} 'Extenciones validas

        ''Recorro las extenciones y las comparo con las de mi archivo
        'For i As Integer = 0 To allowedExtensions.Length - 1
        '    If fileExtension = allowedExtensions(i) Then
        '        fileOK = True
        '    End If
        'Next

        ''si fue correcta la extencion entonces poga la variable FileOk en verdadero
        'If fileOK Then

        '    If oAtt.GetByKey(4981) Then

        '        Try

        '            'oAtt.Lines.SetCurrentLine(nlin)
        '            oAtt.Lines.FileName = fileName2 'Entreguele al objeto de la Di-api el Nombre del Archivo
        '            'oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tNO 'Sobre escriba el metodo
        '            'oAtt.Lines.FileExtension = "PDF"
        '            oAtt.Lines.Add()

        '        Catch ex As Exception

        '            'lblErrorDeffer.Text = "El Archivo no pudo ser cargado." + ex.ToString
        '            MsgBox(ex.Message)
        '        End Try

        '        Dim iAttEntry As Integer = -1
        '        Dim ErrAtt As Integer = oAtt.Update()

        '        If ErrAtt <> 0 Then
        '            MsgBox(SAPConnector.company.GetLastErrorDescription)
        '        Else
        '            'Dim rSet As SAPbobsCOM.Recordset = connx.getRecordSet("Select top 1 AbsEntry from [OATC] order by AbsEntry desc")

        '            'If rSet.RecordCount > 0 Then
        '            '    rSet.MoveFirst()
        '            '    iAttEntry = rSet.Fields.Item(0).Value
        '            'End If

        '            MsgBox("Se Adjunto Actualizacion Exitosamente ")

        '            'Dim RetVal As Integer = 0
        '            'Dim vOrder As SAPbobsCOM.Documents
        '            'vOrder = connx.getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

        '            'If (vOrder.GetByKey(nDocEntry)) Then
        '            '    'test = oContact.CardCode
        '            '    vOrder.AttachmentEntry = iAttEntry
        '            '    'vOrder.Attachments.Add
        '            '    RetVal = vOrder.Update()

        '            '    If RetVal <> 0 Then
        '            '        MsgBox(SAPConnector.company.GetLastErrorDescription)
        '            '    Else
        '            '        MsgBox("Se Actulizao Orden de Compra ")
        '            '    End If
        '            'End If
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub adjuntar()

        Dim oAtt As SAPbobsCOM.Attachments2
        oAtt = connx.getBObject(SAPbobsCOM.BoObjectTypes.oAttachments2)
        oAtt.Lines.Add()

        Dim path As String = "C:\tmp" 'Debes crear una carpeta vacia "Anexos" con un archivo en la solucion de tu proyecto .Net
        oAtt.Lines.SourcePath = path   'Aqui le doy la ruta de origen, que exige el objeto de la DI-API Para poder hacer el Attachment

        'If FileCargarArchivosc.HasFile Then
        Dim fileName As String = "363675_F6535FE4_3423683311.PDF"
        'Dim fileName2 As String = "363675_F6535FE4_3423683311.PDF"
        Dim fileExtension As String 'Creo la Variable FileExtencion
        Dim fileOK As Boolean = False 'Creo una variable de tipo boolean 
        Dim nDocEntry As Integer = 15058

        fileExtension = System.IO.Path.GetExtension(fileName).ToLower() 'Obtengo la extencion de mi archivo y la convierto a Minuscula
        Dim allowedExtensions As String() = {".jpg", ".jpeg", ".png", ".gif", ".pdf"} 'Extenciones validas

        'Recorro las extenciones y las comparo con las de mi archivo
        For i As Integer = 0 To allowedExtensions.Length - 1
            If fileExtension = allowedExtensions(i) Then
                fileOK = True
            End If
        Next

        'si fue correcta la extencion entonces poga la variable FileOk en verdadero
        If fileOK Then

            If oAtt.GetByKey(4985) Then

                Try
                    Dim xlen As Integer = fileName.Length
                    oAtt.Lines.Add()
                    oAtt.Lines.FileName = fileName 'Entreguele al objeto de la Di-api el Nombre del Archivo
                    oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES 'Sobre escriba el metodo
                    'Att.Lines.FileExtension = "PDF"

                    'oAtt.Lines.FileName = fileName2 'Entreguele al objeto de la Di-api el Nombre del Archivo
                    'oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES 'Sobre escriba el metodo
                    ''Att.Lines.FileExtension = "PDF"
                Catch ex As Exception

                    'lblErrorDeffer.Text = "El Archivo no pudo ser cargado." + ex.ToString
                    MsgBox(ex.Message)
                End Try

            End If
        End If

            Dim iAttEntry As Integer = -1
        'Dim ErrAtt As Integer = oAtt.Add()
        Dim ErrAtt As Integer = oAtt.Update()

        If ErrAtt <> 0 Then
            MsgBox(SAPConnector.company.GetLastErrorDescription)
        Else
            'Dim rSet As SAPbobsCOM.Recordset = connx.getRecordSet("Select top 1 AbsEntry from [OATC] order by AbsEntry desc")

            'If rSet.RecordCount > 0 Then
            '    rSet.MoveFirst()
            '    iAttEntry = rSet.Fields.Item(0).Value
            'End If

            MsgBox("Se Adjunto Exitosamente ")

            'Dim RetVal As Integer = 0
            'Dim vOrder As SAPbobsCOM.Documents
            'vOrder = connx.getBObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)

            'If (vOrder.GetByKey(nDocEntry)) Then
            '    'test = oContact.CardCode
            '    vOrder.AttachmentEntry = iAttEntry
            '    'vOrder.Attachments.Add
            '    RetVal = vOrder.Update()

            '    If RetVal <> 0 Then
            '        MsgBox(SAPConnector.company.GetLastErrorDescription)
            '    Else
            '        MsgBox("Se Actulizao Orden de Compra ")
            '    End If
            'End If
        End If
    End Sub


    '    oAttach.Lines.Add
    'oAttach.Lines.FileName = "filename";     //put here actual file name
    'oAttach.Lines.FileExtension= ".ext";        //put here actual file extension
    'oAttach.Lines.SourcePath= "path";        //put here actual file path

    'oAttach.Update();


    'Public Function InsertarCliente(ByVal modeloCliente As ModelCliente) As List(Of String)


    '    Try

    '        For Each valueEncabe As ModelEncaCliente In modeloCliente.cliente
    '            'variables
    '            Dim oB As SAPbobsCOM.BusinessPartners
    '            oB = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
    '            oB.CardCode = valueEncabe.Codigo
    '            oB.CardName = valueEncabe.Nombrecompleto
    '            oB.CardType = SAPbobsCOM.BoCardTypes.cCustomer
    '            oB.Series = 1
    '            oB.FederalTaxID = "000000000000"
    '            oB.GroupCode = valueEncabe.Grupo
    '            oB.Currency = valueEncabe.Moneda

    '            'ACA AGREGO LOS ANEXOS
    '            Dim oATT As SAPbobsCOM.Attachments2 = GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2)
    '            For Each value As clsAnexos In valueEncabe.Anexos
    '                oATT.Lines.Add()
    '                oATT.Lines.FileName = value.FileName
    '                oATT.Lines.FileExtension = value.FileExtension
    '                oATT.Lines.SourcePath = value.SourcePath
    '                oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES

    '                If oATT.Add() <> 0 Then
    '                    Dim Respuesta = oCompany.GetLastErrorCode 'Esto si es incorrecto
    '                Else
    '                    oB.AttachmentEntry = oCompany.GetNewObjectKey 'Este es el entry que genera
    '                End If
    '            Next



    '            If oB.Add() <> 0 Then
    '                Dim Respuesta = oCompany.GetLastErrorCode 'Esto si es incorrecto
    '            Else
    '                Dim Respuesta = oCompany.GetNewObjectKey()
    '            End If
    '        Next

    '    Catch ex As Exception
    '        Dim respuesta = ex.ToString
    '    End Try
    'End Function

    'Public Function ActualizaCliente(ByVal cardcode As String, ByVal modeloCliente As ModelCliente) As List(Of String)

    '    Try

    '        For Each valueEncabe As ModelEncaCliente In modeloCliente.cliente
    '            Dim OB As SAPbobsCOM.BusinessPartners
    '            OB = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
    '            OB.GetByKey(cardcode)
    '            OB.CardName = valueEncabe.Nombrecompleto
    '            OB.CardType = SAPbobsCOM.BoCardTypes.cCustomer
    '            OB.Series = 1
    '            OB.FederalTaxID = "000000000000" 'xnDoc.SelectSingleNode("federaltaxid").InnerText
    '            OB.GroupCode = valueEncabe.Grupo


    '            Dim oATT As SAPbobsCOM.Attachments2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2)
    '            For Each value As clsAnexos In valueEncabe.Anexos
    '                oATT.Lines.Add()
    '                'oATT.Lines.SetCurrentLine(1)
    '                oATT.Lines.FileName = value.FileName
    '                oATT.Lines.FileExtension = value.FileExtension
    '                oATT.Lines.SourcePath = value.SourcePath
    '                oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES

    '                If oATT.Add() <> 0 Then
    '                    Dim respuesta = oCompany.GetLastErrorCode
    '                Else
    '                    OB.AttachmentEntry = oCompany.GetNewObjectKey
    '                End If
    '            Next

    '            If OB.Update() <> 0 Then
    '                Dim Respuesta = oCompany.GetLastErrorCode 'Esto si es incorrecto
    '            Else
    '                Dim Respuesta = oCompany.GetNewObjectKey()
    '            End If
    '        Next

    '    Catch ex As Exception
    '        Dim respuesta = ex.ToString
    '    End Try
    'End Function

End Class