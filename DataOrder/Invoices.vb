Public Class Invoices

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oInvoice As SAPbobsCOM.Documents

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub

    Public Sub dataInvoice(ByVal DocEntry As String, ByVal DocTotal As Double)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim DocNum, CardCode, DocCur, ItemCode, Quantity, Price, DiscPrcnt, TaxCode, WhsCode, Currency, ObjType, LineNum As String
        Dim llError As Long
        Dim lsError As String

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oInvoice = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments)

        Try

            stQueryH = "Select T0.""DocNum"",T0.""CardCode"",T0.""DocCur"" from OPOR T0 where T0.""DocEntry""=" & DocEntry
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                DocNum = oRecSetH.Fields.Item("DocNum").Value
                CardCode = oRecSetH.Fields.Item("CardCode").Value
                DocCur = oRecSetH.Fields.Item("DocCur").Value

                oInvoice.CardCode = CardCode
                oInvoice.DocCurrency = DocCur
                oInvoice.DocDate = DateTime.Now.ToString("dd/MM/yyyy")
                oInvoice.Comments = "Basado en Pedidos " & DocNum & "."
                oInvoice.DocTotal = DocTotal
                oInvoice.DownPaymentType = 1

                stQueryH = "Select T0.""ItemCode"",T0.""Quantity"",T0.""Price"",T0.""DiscPrcnt"",T0.""TaxCode"",T0.""WhsCode"",T0.""Currency"",T0.""DocEntry"",T0.""ObjType"",T0.""LineNum"" from POR1 T0 where T0.""DocEntry""=" & DocEntry
                oRecSetH.DoQuery(stQueryH)

                If oRecSetH.RecordCount > 0 Then

                    oRecSetH.MoveFirst()

                    For i = 0 To oRecSetH.RecordCount - 1

                        ItemCode = oRecSetH.Fields.Item("ItemCode").Value
                        Quantity = oRecSetH.Fields.Item("Quantity").Value
                        Price = oRecSetH.Fields.Item("Price").Value
                        DiscPrcnt = oRecSetH.Fields.Item("DiscPrcnt").Value
                        TaxCode = oRecSetH.Fields.Item("TaxCode").Value
                        WhsCode = oRecSetH.Fields.Item("WhsCode").Value
                        Currency = oRecSetH.Fields.Item("Currency").Value
                        ObjType = oRecSetH.Fields.Item("ObjType").Value
                        LineNum = oRecSetH.Fields.Item("LineNum").Value

                        oInvoice.Lines.ItemCode = ItemCode
                        oInvoice.Lines.Quantity = Quantity
                        oInvoice.Lines.Price = Price
                        oInvoice.Lines.DiscountPercent = DiscPrcnt
                        oInvoice.Lines.TaxCode = TaxCode
                        oInvoice.Lines.WarehouseCode = WhsCode
                        oInvoice.Lines.Currency = Currency
                        oInvoice.Lines.BaseEntry = DocEntry
                        oInvoice.Lines.BaseType = ObjType
                        oInvoice.Lines.BaseLine = LineNum

                        oInvoice.Lines.Add()

                        If i < oRecSetH.RecordCount - 1 Then
                            oRecSetH.MoveNext()
                        End If

                    Next

                End If

                If oInvoice.Add() <> 0 Then

                    SBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)

                End If

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en el evento sobre Agregar facturas de anticpo. " & ex.Message)

        Finally

        End Try

    End Sub

End Class
