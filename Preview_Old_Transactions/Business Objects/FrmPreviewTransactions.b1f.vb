Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Preview_Old_Transactions
    <FormAttribute("PRETRAN", "Business Objects/FrmPreviewTransactions.b1f")>
    Friend Class FrmPreviewTransactions
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub
        Dim DTQuery As String = ""
        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("grdtran").Specific, SAPbouiCOM.Grid)
            Me.StaticText0 = CType(Me.GetItem("lblrec").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("cmbrec").Specific, SAPbouiCOM.ComboBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("PRETRAN", 0)
                ComboBox0.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                bModal = True
            Catch ex As Exception
            End Try
        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub

        Public Sub Load_Grid(ByVal Query As String, ByVal EntryCount As Integer)
            Try
                DTQuery = Query
                If EntryCount = 5 Or EntryCount = 10 Then
                    Query = Query.Remove(11, 1)
                    Query = Query.Insert(11, ComboBox0.Selected.Value)
                Else
                    Query = Query.Remove(7, 5)
                End If
                objform.DataSources.DataTables.Item("DT_0").ExecuteQuery(Query)
                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_0")
                Grid0.RowHeaders.TitleObject.Caption = "#"
                objform.Freeze(True)
                Dim ObjectType As String
                For i As Integer = 0 To Grid0.Columns.Count - 1
                    If Grid0.Columns.Item(i).UniqueID = "ObjType" Then Grid0.Columns.Item(i).Visible = False
                    Grid0.Columns.Item(i).Editable = False
                Next
                ObjectType = Grid0.DataTable.GetValue("ObjType", 0)
                'Grid0.Rows.SelectedRows.Add(0)
                Grid0.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                Dim col As SAPbouiCOM.EditTextColumn
                col = Grid0.Columns.Item(0)
                col.LinkedObjectType = ObjectType
                Grid0.AutoResizeColumns()
                objaddon.objapplication.StatusBar.SetText("...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Grid0.AutoResizeColumns()
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            Try
                If pVal.InnerEvent = True Or DTQuery = "" Then Exit Sub
                objform.Freeze(True)
                Load_Grid(DTQuery, ComboBox0.Selected.Value)
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

    End Class
End Namespace
