Public Class clsInstallment
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBoxColumn
    Private oCombo As SAPbouiCOM.ComboBox
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
    Public Sub LoadForm(ByVal aContractNo As String)
        oForm = oApplication.Utilities.LoadForm(xml_Installments, frm_Installments)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
       
        oTest.DoQuery("Select * from [@Z_CONTRACT] where DocEntry=" & CInt(aContractNo))
        If oTest.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(oForm, "3", oTest.Fields.Item("DocEntry").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "5", oTest.Fields.Item("U_Z_CntNo").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "7", oTest.Fields.Item("U_Z_StartDate").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "9", oTest.Fields.Item("U_Z_EndDate").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "11", oTest.Fields.Item("U_Z_ChgMonth").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "13", oTest.Fields.Item("U_Z_Annualrent").Value)
            oGrid = oForm.Items.Item("15").Specific
            oGrid.DataTable.ExecuteQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where U_Z_ConId=" & CInt(aContractNo))
        Else
            oGrid = oForm.Items.Item("15").Specific
            oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where 1=2")
        End If
        Dim strUser As String = oApplication.Company.UserName
        Dim aChoice As Boolean = False
        oTest.DoQuery("Select isnull(U_Z_Install,'N'),* from OUSR where USER_CODE='" & strUser & "'")
        If oTest.Fields.Item(0).Value = "Y" Then
            If AddtoUDT(oForm) = True Then
                oGrid = oForm.Items.Item("15").Specific
                oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS]  where U_Z_ConId=" & CInt(aContractNo))
            End If
            oForm.Items.Item("14").Visible = True
            aChoice = True
        Else
            If AddtoUDT(oForm) = True Then
                oGrid = oForm.Items.Item("15").Specific
                oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS]  where U_Z_ConId=" & CInt(aContractNo))
            End If
            oForm.Items.Item("14").Visible = True
            aChoice = False
        End If

        Formatgrid(oGrid, aChoice)
        'oForm.DataBrowser.BrowseBy = "1000002"
        'oForm.EnableMenu(mnu_DELETE_ROW, True)
        'oForm.EnableMenu(mnu_ADD_ROW, True)
        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        'Dim oTest As SAPbobsCOM.Recordset
        'oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Dim st As String
        'st = "Update [@Z_CONTRACT] set U_Z_CntNo= U_Z_ConNo +'_'+ convert(varchar,isnull(U_Z_SeqNo,'1'))"
        'oTest.DoQuery(st)
        'oForm.Items.Item("18").Visible = False
        'oForm.Items.Item("19").Visible = False
        'oForm.Items.Item("120").Visible = False
        'oForm.Items.Item("60").Visible = False
        'oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
#End Region

    Private Sub Formatgrid(ByVal aGrid As SAPbouiCOM.Grid, ByVal aflag As Boolean)
        aGrid.Columns.Item(0).Visible = False
        aGrid.Columns.Item(1).Visible = False
        aGrid.Columns.Item(2).TitleObject.Caption = "Month"
        aGrid.Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oCombobox = aGrid.Columns.Item(2)
        For intRow As Integer = 1 To 12
            oCombobox.ValidValues.Add(intRow, MonthName(intRow))
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        aGrid.Columns.Item(3).TitleObject.Caption = "Year"
        aGrid.Columns.Item(3).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oCombobox = aGrid.Columns.Item(3)
        For intRow As Integer = 2005 To 2050
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        aGrid.Columns.Item(2).Editable = False
        aGrid.Columns.Item(3).Editable = False
        aGrid.Columns.Item(4).Editable = aflag
        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
    End Sub


    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim dtFrom, dtTo As Date
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode As String
        Dim dblAnnualRent, dblNoofMonths As Double
        dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "7"))
        dtTo = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "9"))
        Dim otest As SAPbobsCOM.Recordset
        dblAnnualRent = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "13"))
        dblNoofMonths = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "11"))
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where U_Z_ConId=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
        If 1 = 1 Then ' otest.RecordCount <= 0 Then
            oUserTable = oApplication.Company.UserTables.Item("Z_CONINS")
            oGrid = aForm.Items.Item("15").Specific
            oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where 1=2")
            While dtFrom <= dtTo
                otest.DoQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where U_Z_Month=" & Month(dtFrom) & " and U_Z_Year=" & Year(dtFrom) & " and   U_Z_ConId=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
                If otest.RecordCount <= 0 Then
                    strCode = oApplication.Utilities.getMaxCode("@Z_CONINS", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_ConID").Value = CInt(oApplication.Utilities.getEdittextvalue(aForm, "3"))
                    oUserTable.UserFields.Fields.Item("U_Z_ConNo").Value = (oApplication.Utilities.getEdittextvalue(aForm, "5"))
                    oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "7"))
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "9"))
                    oUserTable.UserFields.Fields.Item("U_Z_NoofMonths").Value = (oApplication.Utilities.getEdittextvalue(aForm, "11"))
                    oUserTable.UserFields.Fields.Item("U_Z_AnnualRent").Value = (oApplication.Utilities.getEdittextvalue(aForm, "13"))
                    oUserTable.UserFields.Fields.Item("U_Z_Month").Value = Month(dtFrom)
                    oUserTable.UserFields.Fields.Item("U_Z_Year").Value = Year(dtFrom)
                    oUserTable.UserFields.Fields.Item("U_Z_Amount").Value = dblAnnualRent / dblNoofMonths
                    If oUserTable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
                dtFrom = DateAdd(DateInterval.Month, 1, dtFrom)
            End While
        End If
        Return True

    End Function


    Private Function AddtoUDT_Table(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim dtFrom, dtTo As Date
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode As String
        Dim dblAnnualRent, dblNoofMonths As Double
        dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "7"))
        dtTo = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "9"))
        Dim otest As SAPbobsCOM.Recordset
        dblAnnualRent = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "13"))
        dblNoofMonths = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "11"))
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
         If 1 = 1 Then ' otest.RecordCount <= 0 Then
            oUserTable = oApplication.Company.UserTables.Item("Z_CONINS")
            oGrid = aForm.Items.Item("15").Specific
            ' oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where 1=2")
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                If oUserTable.GetByKey(strCode) Then
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_ConID").Value = CInt(oApplication.Utilities.getEdittextvalue(aForm, "3"))
                    oUserTable.UserFields.Fields.Item("U_Z_ConNo").Value = (oApplication.Utilities.getEdittextvalue(aForm, "5"))
                    oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "7"))
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "9"))
                    oUserTable.UserFields.Fields.Item("U_Z_NoofMonths").Value = (oApplication.Utilities.getEdittextvalue(aForm, "11"))
                    oUserTable.UserFields.Fields.Item("U_Z_AnnualRent").Value = (oApplication.Utilities.getEdittextvalue(aForm, "13"))
                    oUserTable.UserFields.Fields.Item("U_Z_Month").Value = oGrid.DataTable.GetValue(2, intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Year").Value = oGrid.DataTable.GetValue(3, intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Amount").Value = oGrid.DataTable.GetValue("Amount", intRow)
                    If oUserTable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next
        End If
        Return True

    End Function
#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_ItemGroup
                    If pVal.BeforeAction = False Then
                        '  LoadForm()
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

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
            If pVal.FormTypeEx = frm_Installments Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "14"
                                        If oApplication.SBO_Application.MessageBox("Do you want to save the details ?", , "Yes", "No") = 2 Then
                                            Exit Sub
                                        End If
                                        If AddtoUDT_Table(oForm) = True Then
                                            oForm.Close()
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
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
