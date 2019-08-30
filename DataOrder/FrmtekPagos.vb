Public Class FrmtekPagos

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openForm(ByVal psDirectory As String, ByVal psDocEntry As String, ByVal psTotal As Double)
        Dim stQueryH, stQueryH2, Status As String
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        'Dim Monto As Integer

        Try

            oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            csFormUID = "tekPagos"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If

            '--- Referencia de Forma
            setForm(csFormUID)

            '" & psDocEntry & "

            stQueryH = "call ""Total_de_Pedidos"" ('" & psDocEntry & "')"

            oRecSetH.DoQuery(stQueryH)

            stQueryH2 = "Select ""DocStatus"" from ""OPOR"" where ""DocEntry""=" & psDocEntry

            oRecSetH2.DoQuery(stQueryH2)

            If oRecSetH.RecordCount > 0 Or oRecSetH2.RecordCount > 0 Then
                oRecSetH.MoveFirst()
                oRecSetH2.MoveFirst()

                Monto = oRecSetH.Fields.Item("Monto").Value
                Status = oRecSetH2.Fields.Item("DocStatus").Value


                If Monto >= psTotal Or Status = "C" Then

                    coForm.Items.Item("4").Enabled = False

                End If

            End If

            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

            Return Monto

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmTratamientoPedidos. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources() As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources()
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources() As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid

        Try
            bindUserDataSources = 0

            loDS = coForm.DataSources.UserDataSources.Add("dsMonto", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            loText = coForm.Items.Item("3").Specific  'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsMonto")   ' uno mi userdatasources a mi caja de texto

            oGrid = coForm.Items.Item("1").Specific
            oDataTable = coForm.DataSources.DataTables.Add("Movimientos")
            oGrid.DataTable = oDataTable

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    '----- carga los procesos de carga
    Public Function cargarMovimientos(stDocEntry As String)
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset

        Try

            coForm.DataSources.UserDataSources.Item("dsMonto").Value = Nothing
            oGrid = coForm.Items.Item("1").Specific
            oGrid.DataTable.Clear()
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            stQuery = "call ""Movimientos_de_Pedidos"" ('" & stDocEntry & "')"

            oGrid.DataTable.ExecuteQuery(stQuery)

            Return 0

        Catch ex As Exception

            MsgBox("FrmTratamientoPedidos. cargarDetalle: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function

End Class
