Namespace Preview_Old_Transactions

    Public Class clsRightClickEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Dim ocombo As SAPbouiCOM.ComboBox
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "133", "139"
                        objform = objaddon.objapplication.Forms.ActiveForm
                        If objform.Items.Item("3").Specific.Selected.Value = "S" Then Exit Sub
                        If eventInfo.BeforeAction Then
                            If objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If eventInfo.ColUID = 0 Then
                                    If objform.Items.Item("4").Specific.String = "" Then Exit Sub
                                    RowIndex = eventInfo.Row
                                    If Not objaddon.objapplication.Menus.Item("1280").SubMenus.Exists("VPRICE") Then RightClickMenu_Add("1280", "VPRICE", "Preview Recent Price", 0)
                                Else
                                    If objaddon.objapplication.Menus.Item("1280").SubMenus.Exists("VPRICE") Then RightClickMenu_Delete("1280", "VPRICE")
                                End If
                            End If
                        Else
                            If objaddon.objapplication.Menus.Item("1280").SubMenus.Exists("VPRICE") Then RightClickMenu_Delete("1280", "VPRICE")
                        End If
                End Select
            Catch ex As Exception
            End Try
        End Sub

        Private Sub RightClickMenu_Add(ByVal MainMenu As String, ByVal NewMenuID As String, ByVal NewMenuName As String, ByVal position As Integer)
            Dim omenus As SAPbouiCOM.Menus
            Dim omenuitem As SAPbouiCOM.MenuItem
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If Not omenuitem.SubMenus.Exists(NewMenuID) Then
                oCreationPackage.UniqueID = NewMenuID
                oCreationPackage.String = NewMenuName
                oCreationPackage.Position = position
                oCreationPackage.Enabled = True
                omenus = omenuitem.SubMenus
                omenus.AddEx(oCreationPackage)
            End If
        End Sub

        Private Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Dim omenuitem As SAPbouiCOM.MenuItem
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If omenuitem.SubMenus.Exists(NewMenuID) Then
                objaddon.objapplication.Menus.RemoveEx(NewMenuID)
            End If
        End Sub



    End Class

End Namespace
