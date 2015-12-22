Public Class clsInsurance
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
    Dim count, RowtoDelete, MatrixId As Integer
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Insurance) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Insurance, frm_Insurance)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataBrowser.BrowseBy = "14"

        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        AddChooseFromList(oForm)
        '  databind(oForm)
        ' oForm.Items.Item("38").DisplayDesc = True
        oForm.Freeze(False)
    End Sub

#Region "Add Choose From List"
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
            oCFL = oCFLs.Item("CFL_2")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)


            oCFL = oCFLs.Item("CFL_3")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "U_Z_ProFlg"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oCombo = aForm.Items.Item("23").Specific
            For intRow As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception

                End Try
            Next
            oCombo.ValidValues.Add("", "")
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("SELECT T0.[GroupNum], T0.[PymntGroup] FROM OCTG T0")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub FillCombo(ByVal aForm As SAPbouiCOM.Form)
        Dim strPropertyCode, strstring As String
        oCombo = aForm.Items.Item("34").Specific
        strPropertyCode = oApplication.Utilities.getEdittextvalue(aForm, "6")
        strstring = "Select DocEntry,U_Z_Desc from [@Z_PROPUNIT] where U_Z_PropCode=" & strPropertyCode & " order by DocEntry"
        aForm.Freeze(True)
        oApplication.Utilities.FillComboBox(oCombo, strstring)
        aForm.Freeze(False)
    End Sub

#Region "AddRow /Delete Row"

    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strDocEntry As String
        Return True
        If validation(aForm) = False Then
            Return False
        End If
        'AssignLineNo(aForm)
        Return True
    End Function
#End Region

#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject, strActivity, strActivity1 As String
        Dim oTemp As SAPbobsCOM.Recordset
        oCombobox = aform.Items.Item("4").Specific
        strProject = oCombobox.Selected.Value
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If strProject = "" Then
                oApplication.Utilities.Message("Project code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oTemp.DoQuery("Select * from [@Z_PROP1] where U_Z_PRJCODE='" & strProject & "'")
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Project code already exists...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If

        Dim dblBudget, dblLineBudget As Double
        Dim strHours As String
        Dim dblHours As Double

        dblBudget = 0
        dblLineBudget = 0

        oMatrix = aform.Items.Item("12").Specific
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        For intRow As Integer = 1 To oMatrix.RowCount
            strCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intRow)
            strActivity = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intRow)
            If strCode <> "" Then
                strHours = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", intRow)
                If strHours <> "" Then
                    dblHours = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_2", intRow))
                    If (dblHours > 0) Then
                        dblHours = dblHours * 8
                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", intRow, dblHours)
                    End If
                End If

                strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)
                If strcode1 = "" Then
                    oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If CInt(strcode1) <= 0 Then
                    oApplication.Utilities.Message("No of Hours is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If strActivity = "" Then
                    oApplication.Utilities.Message("Activity detail is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                For intLoop As Integer = intRow + 1 To oMatrix.RowCount
                    strcode1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intLoop)
                    strActivity1 = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", intLoop)
                    If strcode1 <> "" Then
                        If strCode.ToUpper = strcode1.ToUpper And strActivity.ToUpper = strActivity1.ToUpper Then
                            oApplication.Utilities.Message("Process and Activity details already exists : " & strCode & "-" & strActivity, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
                Dim strdays As String
                strdays = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", intRow)

                dblLineBudget = dblLineBudget + oApplication.Utilities.getDocumentQuantity(strdays)
            Else
                oMatrix.DeleteRow(intRow)
            End If
        Next

        oMatrix = aform.Items.Item("12").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Process details are missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        dblBudget = CDbl(oApplication.Utilities.getEdittextvalue(aform, "8"))
        If dblBudget <> dblLineBudget Then
            If oApplication.SBO_Application.MessageBox("Total man days does not match with Line man days. Do you want to save this document ? ", , "Continue", "Cancel") = 2 Then
                Return False
            Else


            End If
        End If

        Return True
    End Function
#End Region
    
#End Region

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Insurance
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oApplication.Utilities.setEdittextvalue(oForm, "14", oApplication.Utilities.getMaxCode("@Z_INSURANCE", "DocEntry"))
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
            If pVal.FormTypeEx = frm_Insurance Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "17" And pVal.CharPressed = 9 Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                    Dim strIns As String
                                    strIns = oApplication.Utilities.getEdittextvalue(oForm, "17")
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Select * from [@Z_PROPUNIT] where U_Z_ProItemCode='" & strIns & "'")
                                    If otest.RecordCount <= 0 Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "17", "")
                                        strIns = ""
                                    Else
                                        Exit Sub
                                    End If
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    clsChooseFromList.ItemUID = pVal.ItemUID
                                    clsChooseFromList.SourceFormUID = FormUID
                                    clsChooseFromList.SourceLabel = 0
                                    clsChooseFromList.CFLChoice = "ProUnit" 'oCombo.Selected.Value
                                    clsChooseFromList.choice = "Bin"
                                    'clsChooseFromList.Documentchoice = oApplication.Utilities.getEdittextvalue(oForm, "6") 'TenCode
                                    clsChooseFromList.ItemCode = "" 'oApplication.Utilities.getEdittextvalue(oForm, "6") 'Unit Code
                                    ' clsChooseFromList.BinDescrUID = "BinToBinHeader"
                                    clsChooseFromList.sourceColumID = ""
                                    oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID

                                End Select
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
                                        If pVal.ItemUID = "12" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If

                                        'If pVal.ItemUID = "17" Then
                                        '    val = oDataTable.GetValue("ItemCode", 0)
                                        '    oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        'End If
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
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#End Region
End Class
