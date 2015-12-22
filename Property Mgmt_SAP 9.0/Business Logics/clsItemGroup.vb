Public Class clsItemGroup
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_ItemGroup, frm_Itemgroup)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        databind(oForm)
        oForm.Freeze(False)
    End Sub




#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim strCardCode As String
            strCardCode = oApplication.Utilities.getEdittextvalue(aForm, "4")

            oGrid = aForm.Items.Item("6").Specific
            oGrid.DataTable.ExecuteQuery("Select * from [@DABT_OITB] where U_CardCode='" & strCardCode & "' order by Convert(Numeric,Code)")
            FormatGrids(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

#Region "FormatGrids"
    Private Sub FormatGrids(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("6").Specific
        oGrid.Columns.Item(0).Visible = False
        oGrid.Columns.Item(0).TitleObject.Caption = "Code"
        oGrid.Columns.Item(1).TitleObject.Caption = "Name"
        oGrid.Columns.Item(1).Visible = False
        oGrid.Columns.Item(2).TitleObject.Caption = "Card Code"
        oGrid.Columns.Item(2).Visible = False
        oGrid.Columns.Item(3).TitleObject.Caption = "Item Group"
        oGrid.Columns.Item(3).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oRecset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Select ItmsGrpCod,ItmsGrpnam from OITB order by ItmsGrpCod")
        oCombobox = oGrid.Columns.Item(3)
        For intRow As Integer = 0 To oRecset.RecordCount - 1
            oCombobox.ValidValues.Add(oRecset.Fields.Item(0).Value, oRecset.Fields.Item(1).Value)
            oRecset.MoveNext()
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)

        oGrid = aForm.Items.Item("6").Specific
        If oGrid.DataTable.GetValue("U_ITMSGRPCOD", oGrid.DataTable.Rows.Count - 1) <> "" Then
            oGrid.DataTable.Rows.Add()
        End If
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Dim orec As SAPbobsCOM.Recordset
        Dim strTablename As String
        strTablename = ""
        oGrid = aForm.Items.Item("6").Specific
        strTablename = "[@DABT_OITB]"

        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                orec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orec.DoQuery("Update" & strTablename & " set Name = 'D' +Name where code='" & oGrid.DataTable.GetValue("Code", intRow) & "'")
                oGrid.DataTable.Rows.Remove(intRow)
                Exit For
            End If
        Next
    End Sub
#End Region

#Region "AddtoUDT"
    Private Function AddToUDT(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strItem, strCode, strWhs, strBin, strwhsdesc, strbindesc, strTo, strHeaderRef, strConditionType, strfromdate, strCardCode As String
        Dim oTempRec, otemp As SAPbobsCOM.Recordset
        Dim ousertable As SAPbobsCOM.UserTable
        Dim ocheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oedittext As SAPbouiCOM.EditTextColumn
        Dim dblPercentage As Double
        Dim dtFrom, dtTo As Date
        Dim oBPGrid As SAPbouiCOM.Grid

        oBPGrid = aform.Items.Item("6").Specific
        strCardCode = oApplication.Utilities.getEdittextvalue(aform, "4")
        ousertable = oApplication.Company.UserTables.Item("DABT_OITB")
        For intLoop As Integer = 0 To oBPGrid.DataTable.Rows.Count - 1
            strCode = oBPGrid.DataTable.GetValue(0, intLoop)
            oCombobox = oBPGrid.Columns.Item("U_ITMSGRPCOD")
            strfromdate = oCombobox.GetSelectedValue(intLoop).Value
            If strfromdate <> "" Then
                otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemp.DoQuery("Select * from [@DABT_OITB] where U_CardCode='" & strCardCode & "' and U_ITMSGRPCOD='" & strfromdate & "'")
                If otemp.RecordCount > 0 Then
                    strCode = otemp.Fields.Item("Code").Value
                Else
                    strCode = oBPGrid.DataTable.GetValue(0, intLoop)
                End If
                If strCode <> "" Then
                    ousertable.GetByKey(strCode)
                    ousertable.Name = strCode
                    ousertable.UserFields.Fields.Item("U_CardCode").Value = strCardCode

                    oCombobox = oBPGrid.Columns.Item("U_ITMSGRPCOD")
                    ousertable.UserFields.Fields.Item("U_ITMSGRPCOD").Value = oCombobox.GetSelectedValue(intLoop).Value
                    If ousertable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Else
                    strCode = oApplication.Utilities.getMaxCode("@DABT_OITB", "Code")
                    ousertable.Code = strCode
                    ousertable.Name = strCode
                    ousertable.UserFields.Fields.Item("U_CardCode").Value = strCardCode
                    oCombobox = oBPGrid.Columns.Item("U_ITMSGRPCOD")
                    ousertable.UserFields.Fields.Item("U_ITMSGRPCOD").Value = oCombobox.GetSelectedValue(intLoop).Value
                    If ousertable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            End If
        Next
        oRecset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Delete from [@DABT_OITB] where name like 'D%'")
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_None)
        Return True
    End Function
    
#End Region




#End Region

#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_ItemGroup
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        'EnableControls(oForm)
                    End If

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
        End Try
    End Sub
#End Region


    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
               
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Itemgroup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "2" Then
                                    Dim oDelrec As SAPbobsCOM.Recordset
                                    oDelrec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oDelrec.DoQuery("Update [@DABT_OITB] set name=code where name like 'D%'")
                                End If
                        End Select


                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        If AddToUDT(oForm) Then
                                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            databind(oForm)
                                        End If
                                    Case "12"
                                        AddRow(oForm)
                                    Case "13"
                                        deleterow(oForm)

                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        'If pVal.ItemUID = "11" And (pVal.ColUID = "U_FreeItem1" Or pVal.ColUID = "U_FreeItem2" Or pVal.ColUID = "U_FreeItem3" Or pVal.ColUID = "U_FreeItem4") Then
                                        '    val = oDataTable.GetValue("ItemCode", 0)
                                        '    oGrid = oForm.Items.Item("11").Specific
                                        '    oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        'End If
                                        'If pVal.ItemUID = "16" And (pVal.ColUID = "U_ItemCode") Then
                                        '    val = oDataTable.GetValue("ItemCode", 0)
                                        '    oGrid = oForm.Items.Item("16").Specific
                                        '    oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                        'End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If

        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#End Region
End Class
