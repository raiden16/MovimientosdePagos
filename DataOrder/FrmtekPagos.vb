Public Class FrmtekPagos

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Sub openForm(ByVal psDirectory As String)

        Try
            csFormUID = "tekPagos"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")
            End If

            '--- Referencia de Forma
            setForm(csFormUID)

            coForm.Title = coForm.Title & ". Pedido: " & stDocNum
            '---- refresca forma
            coForm.Refresh()
            coForm.Visible = True

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmTratamientoPedidos. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Sub

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

            oGrid = coForm.Items.Item("1").Specific
            oGrid.DataTable.Clear()
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            stQuery = "Select 

                       case when ifnull(T5.""DocNum"",0)=0 then 'Factura de Anticipo' else 'Pago Efectuado' end as ""Movimiento"",

                       case when ifnull(T5.""DocNum"",0)=0 then case when T3.""DocCur""<>'MXN' then T3.""DocTotalFC"" else T3.""DocTotal"" end else  
                       case when T5.""DocCurr""<>'MXN' then T5.""DocTotalFC"" else T5.""DocTotal"" end
                       end as ""Monto"",

                       case when ifnull(T5.""DocNum"",0)=0 then T3.""DocCur"" else T5.""DocCurr"" end as ""Moneda""

                       From ""OPOR"" T0 
                       Inner Join ""POR1"" T1 on T1.""DocEntry""=T0.""DocEntry"" 
                       Left Outer Join ""DPO1"" T2 on T2.""BaseEntry""=T1.""DocEntry"" and T2.""BaseType""=T1.""ObjType"" and T2.""BaseLine""=T1.""LineNum"" and T2.""ItemCode""=T1.""ItemCode"" 
                       Left Outer Join ""ODPO"" T3 on T3.""DocEntry""=T2.""DocEntry"" 
                       Left Outer Join ""VPM2"" T4 on T4.""DocEntry""=T3.""DocEntry"" and T4.""InvType""=T3.""ObjType""
                       Left Outer Join ""OVPM"" T5 on T5.""DocEntry""=T4.""DocNum""
                       Where T0.""DocEntry"" = " & stDocEntry & " 
                       group by T3.""DocNum"",T3.""ObjType"",T3.""DocEntry"",T5.""DocNum"",T3.""DocTotal"",T5.""DocCurr"",T5.""DocTotalFC"",T5.""DocTotal"",T3.""DocCur"",T3.""DocTotalFC"",T3.""DocTotal""
                      "

            oGrid.DataTable.ExecuteQuery(stQuery)

            Return 0

        Catch ex As Exception

            MsgBox("FrmTratamientoPedidos. cargarDetalle: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function

End Class
