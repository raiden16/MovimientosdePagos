Friend Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF

    Public Sub New()
        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        setFilters()

    End Sub

    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        End Try
    End Sub

    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
            'Finally
        End Try
    End Sub

    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        Finally
            loRecSet = Nothing
        End Try
    End Sub

    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try

            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            lofilter.AddEx(142) '// FORMA PEDIDO DE COMPRAS
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx("tekPagos") '// FORMA PAGOS
            lofilter.AddEx(142) '// FORMA PEDIDO DE COMPRAS

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub

    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        ''SBOApplication.MessageBox("Action: " & pVal.Before_Action & "  Type: " & pVal.FormTypeEx)
        If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then
        Else
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then
                Select Case pVal.FormTypeEx
                    Case "tekPagos"                     '////// FORMA RESERVA DE PEDIDOS
                        'frmPagosControllerAfter(FormUID, pVal)
                    Case 142                     '////// FORMA RESERVA DE PEDIDOS
                        frmPOControllerAfter(FormUID, pVal)

                End Select
            End If
        End If

    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA PEDIDOS DE COMPRAS
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmPOControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        Dim oPO As PO
        Dim otekPagos As FrmtekPagos
        Dim coForm As SAPbouiCOM.Form
        Dim DocNum, DocEntry, stTabla As String
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim oDatatable As SAPbouiCOM.DBDataSource
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType
                            '///// Carga de formas
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    oPO = New PO
                    oPO.addFormItems(FormUID)

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                                    '--- Boton Movimientos del Pedido
                        Case "btMov"

                            stTabla = "OPOR"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            DocNum = oDatatable.GetValue("DocNum", 0)

                            MsgBox(DocNum)

                            If (DocNum Is Nothing) Or (DocNum = "") Then
                                DocNum = "0"
                            End If

                            stQueryH = "Select T1.""DocEntry"" from ""OPOR"" T1 where T1.""DocNum""=" & DocNum
                            oRecSetH.DoQuery(stQueryH)

                            If oRecSetH.RecordCount > 0 Then

                                oRecSetH.MoveFirst()
                                DocEntry = oRecSetH.Fields.Item("DocEntry").Value

                                MsgBox(DocEntry)

                                otekPagos = New FrmtekPagos
                                otekPagos.openForm(csDirectory)
                                otekPagos.cargarMovimientos(DocEntry)

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally
            oPO = Nothing
        End Try
    End Sub

End Class
