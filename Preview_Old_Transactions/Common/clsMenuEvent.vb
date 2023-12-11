Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace Preview_Old_Transactions

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "133", "139"
                        'Default_Sample_MenuEvent(pVal, BubbleEvent)
                        If pVal.BeforeAction = True Then Exit Sub
                        objform = objaddon.objapplication.Forms.ActiveForm
                        If pVal.MenuUID = "VPRICE" Then
                            If objform.Items.Item("3").Specific.Selected.Value = "S" Then Exit Sub
                            Dim objmatrix As SAPbouiCOM.Matrix
                            objmatrix = objform.Items.Item("38").Specific
                            Dim MatRow As Integer
                            Try
                                MatRow = objmatrix.GetCellFocus().rowIndex
                            Catch ex As Exception
                                If RowIndex <> 0 Then MatRow = RowIndex
                                RowIndex = 0
                            End Try

                            Dim txtDate As SAPbouiCOM.EditText
                            Dim DocDate As Date
                            Dim StrQuery As String = ""
                            If Not objaddon.FormExist("PRETRAN") Then
                                If objaddon.objapplication.Forms.ActiveForm.TypeEx = "139" Then ' Sales Order
                                    txtDate = objform.Items.Item("10").Specific
                                    DocDate = Date.ParseExact(txtDate.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                    StrQuery = Get_Recent_Transactions(objform.UniqueID, objform.Items.Item("4").Specific.String, objmatrix.Columns.Item("1").Cells.Item(MatRow).Specific.String, DocDate.ToString("yyyyMMdd"), objform.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0))
                                ElseIf objaddon.objapplication.Forms.ActiveForm.TypeEx = "133" Then ' Sales Invoice
                                    txtDate = objform.Items.Item("10").Specific
                                    DocDate = Date.ParseExact(txtDate.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                    StrQuery = Get_Recent_Transactions(objform.UniqueID, objform.Items.Item("4").Specific.String, objmatrix.Columns.Item("1").Cells.Item(MatRow).Specific.String, DocDate.ToString("yyyyMMdd"), objform.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0))
                                End If
                                Dim objRs As SAPbobsCOM.Recordset
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRs.DoQuery(StrQuery)
                                If objRs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Sub
                                Dim activeform As New FrmPreviewTransactions
                                activeform.Show()
                                activeform.UIAPIRawForm.Left = objform.Left + 100
                                activeform.UIAPIRawForm.Top = objform.Top + 100
                                activeform.Load_Grid(StrQuery, 5) ' objRs.Fields.Item("ObjType").Value.ToString,
                            End If
                        End If
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Function Get_Recent_Transactions(ByVal FormUID As String, ByVal CardCode As String, ByVal ItemCode As String, ByVal DocDate As String, ByVal DocEntry As String) As String
            Try
                'Dim objmatrix As SAPbouiCOM.Matrix
                Dim strSQL As String
                'Dim objRs As SAPbobsCOM.Recordset
                objform = objaddon.objapplication.Forms.Item(FormUID)
                'objmatrix = objform.Items.Item("38").Specific
                Dim HeaderTable, LineTable As String
                If objform.TypeEx = "133" Then ' Sales Invoice
                    HeaderTable = "OINV" : LineTable = "INV1"
                ElseIf objform.TypeEx = "139" Then ' Sales Order
                    HeaderTable = "ORDR" : LineTable = "RDR1"
                End If
                If objaddon.HANA Then
                    strSQL = "select Top 5  T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"", T1.""ItemCode"",T1.""Quantity"",T1.""Price"",T1.""U_Scheme"" ""Scheme"",T0.""ObjType"""
                    strSQL += vbCrLf + "from " & LineTable & " T1 Left Join " & HeaderTable & " T0 on T0.""DocEntry""=T1.""DocEntry"" where T0.""CardCode""='" & CardCode & "' and T1.""ItemCode""='" & ItemCode & "'"
                    strSQL += vbCrLf + "and T0.""DocDate""<='" & DocDate & "' "
                    If DocEntry <> "" Then
                        strSQL += vbCrLf + " And T0.""DocEntry""<>" & DocEntry & "  "
                    End If
                    strSQL += vbCrLf + "Order by T0.""DocDate"" Desc"
                Else
                    strSQL = "select Top 5  T0.DocEntry,T0.DocNum,T0.DocDate, T1.ItemCode,T1.Quantity,T1.Price,T1.U_Scheme Scheme,T0.ObjType"
                    strSQL += vbCrLf + "from " & LineTable & " T1 Left Join " & HeaderTable & " T0 on T0.DocEntry=T1.DocEntry where T0.CardCode='" & CardCode & "' and T1.ItemCode='" & ItemCode & "'"
                    strSQL += vbCrLf + "and T0.DocDate<='" & DocDate & "' "
                    If DocEntry <> "" Then
                        strSQL += vbCrLf + " And T0.DocEntry<>" & DocEntry & "  "
                    End If
                    strSQL += vbCrLf + "Order by T0.DocDate Desc"
                End If
                Return strSQL
            Catch ex As Exception

            End Try
        End Function

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim oUDFForm As SAPbouiCOM.Form
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID

                        Case "6913"

                    End Select
                Else
                    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Select Case pval.MenuUID
                        Case "1284" 'Cancel

                        Case "1281" 'Find

                        Case "1287" 'Duplicate

                        Case Else

                    End Select
                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Production Order"

        Private Sub Production_Order_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OAPI")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("tposdate").Enabled = True
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("tduedate").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                        Case "1282" ' Add Mode
                            objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource, "MIAPSI")

                        Case "1288", "1289", "1290", "1291"

                        Case "1293"
                            DeleteRow(Matrix0, "@MIPL_API1")
                        Case "1292"
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                        Case "1304" 'Refresh
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

#End Region



        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub
    End Class
End Namespace