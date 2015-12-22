Public Class clsTermTransaction
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox, oCombobox1, oCombobox2 As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, strQuery As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim ORecSet As SAPbobsCOM.Recordset
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_TermTransaction, frm_TermTransaction)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.DataBrowser.BrowseBy = "4"
            AddChooseFromList(oForm)
            databind(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False

            oCFLCreationParams.ObjectType = "Z_CONTRACT"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)

            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Status"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "AGR"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        oEditText = aForm.Items.Item("4").Specific
        oEditText.ChooseFromListUID = "CFL2"
        oEditText.ChooseFromListAlias = "U_Z_CntNo"
    End Sub
    Private Sub PopulateDetails(ByVal aForm As SAPbouiCOM.Form, ByVal ContNo As String)
        Try
            ORecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "select * from [@Z_CONTRACT] where U_Z_CNTNO='" & ContNo & "'"
            ORecSet.DoQuery(strQuery)
            If ORecSet.RecordCount > 0 Then
                oApplication.Utilities.setEdittextvalue(aForm, "41", ORecSet.Fields.Item("DocNum").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "5", ORecSet.Fields.Item("U_Z_ConNo").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "7", ORecSet.Fields.Item("U_Z_ContDate").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "9", ORecSet.Fields.Item("U_Z_UnitCode").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "10", ORecSet.Fields.Item("U_Z_Desc").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "14", ORecSet.Fields.Item("U_Z_StartDate").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "16", ORecSet.Fields.Item("U_Z_EndDate").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "18", ORecSet.Fields.Item("U_Z_ChgMonth").Value)
                Try
                    oApplication.Utilities.setEdittextvalue(aForm, "20", ORecSet.Fields.Item("U_Z_TermDate").Value)
                Catch ex As Exception

                End Try

                oApplication.Utilities.setEdittextvalue(aForm, "22", ORecSet.Fields.Item("U_Z_Period").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "24", ORecSet.Fields.Item("U_Z_ChgAmt").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "26", ORecSet.Fields.Item("U_Z_TenCode").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "27", ORecSet.Fields.Item("U_Z_TenName").Value)
                oCombobox = aForm.Items.Item("29").Specific
                oCombobox.Select(ORecSet.Fields.Item("U_Z_IsCommission").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                oApplication.Utilities.setEdittextvalue(aForm, "31", ORecSet.Fields.Item("U_Z_Comm").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "35", ORecSet.Fields.Item("U_Z_Annualrent").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "37", ORecSet.Fields.Item("U_Z_Monthly").Value)
                oApplication.Utilities.setEdittextvalue(aForm, "39", ORecSet.Fields.Item("U_Z_Deposit").Value)
                oCombobox1 = aForm.Items.Item("33").Specific
                oCombobox1.Select(ORecSet.Fields.Item("U_Z_ProType").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            ORecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEdittextvalue(aForm, "20") = "" Then
                oApplication.Utilities.Message("Termination date is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "22") = "" Then
                oApplication.Utilities.Message("Termination Period is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "22") = "" Then
                oApplication.Utilities.Message("Termination charge is missing", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            strQuery = "Select * from [@Z_TCONTRACT] where U_Z_DocNo='" & oApplication.Utilities.getEdittextvalue(aForm, "41") & "' and U_Z_Status<>'R'"
            ORecSet.DoQuery(strQuery)
            If ORecSet.RecordCount > 0 Then
                oApplication.Utilities.Message("Contract Number already exists...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_TermTransaction Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "42" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim strcode As String
                                    strcode = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    Dim objct As New clsTenContracts
                                    objct.LoadForm1(strcode)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "4" And pVal.CharPressed <> 9 And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2 As String
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
                                        If pVal.ItemUID = "4" Then
                                            val = oDataTable.GetValue("U_Z_CNTNO", 0)
                                            PopulateDetails(oForm, val)
                                            oApplication.Utilities.setEdittextvalue(oForm, "41", oDataTable.GetValue("DocEntry", 0))
                                            oApplication.Utilities.setEdittextvalue(oForm, "4", val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_TermTransaction
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_TCONTRACT")
                If oDataSrc_Line.GetValue("U_Z_STATUS", 0).Trim <> "O" Then
                    oForm.Items.Item("1").Enabled = False
                Else
                    oForm.Items.Item("1").Enabled = True
                End If
            End If


            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Dim sQuery As String
                Dim stXML As String = BusinessObjectInfo.ObjectKey
                stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Termination ContractsParams><DocEntry>", "")
                stXML = stXML.Replace("</DocEntry></Termination ContractsParams>", "")
                Dim otest As SAPbobsCOM.Recordset
                otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If stXML <> "" Then
                    otest.DoQuery("select * from [@Z_TCONTRACT]  where DocEntry=" & stXML)
                    If otest.RecordCount > 0 Then
                        Dim intTempID As String = oApplication.Utilities.GetTemplateID(oForm, "TER")
                        If intTempID <> "0" Then
                            oApplication.Utilities.UpdateApprovalRequired("@Z_TCONTRACT", "DocEntry", otest.Fields.Item("DocEntry").Value, "Y", intTempID)
                            oApplication.Utilities.InitialMessage("Termination Contracts Request", otest.Fields.Item("DocEntry").Value, oApplication.Utilities.DocApproval(oForm, "TER"), intTempID, otest.Fields.Item("U_Z_TENNAME").Value, "TER")
                            sQuery = "Update ""@Z_CONTRACT"" Set U_Z_TerStatus = 'Y',U_Z_TerAppStatus='P' Where DocEntry=" & otest.Fields.Item("U_Z_DocNo").Value
                            otest.DoQuery(sQuery)
                        Else
                            oApplication.Utilities.UpdateApprovalRequired("@Z_TCONTRACT", "DocEntry", otest.Fields.Item("DocEntry").Value, "N", intTempID)
                            sQuery = "Update ""@Z_CONTRACT"" Set U_Z_TerStatus = 'Y',U_Z_Status='TER',U_Z_TerAppStatus='A' Where DocEntry=" & otest.Fields.Item("U_Z_DocNo").Value
                            otest.DoQuery(sQuery)
                            sQuery = "Update ""@Z_TCONTRACT"" Set U_Z_Status = 'A' Where DocEntry=" & stXML
                            otest.DoQuery(sQuery)
                        End If
                    End If

                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class

