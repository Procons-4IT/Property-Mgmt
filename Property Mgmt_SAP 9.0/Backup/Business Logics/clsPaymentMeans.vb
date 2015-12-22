Public Class clsPaymentmeans
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
    Private Sub LoadForm(ByVal aForm As SAPbouiCOM.Form)
        Dim strCardCode, strCardtype As String
        If (aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
            oCombo = aForm.Items.Item("40").Specific
            If oCombo.Selected.Value = "C" Then
                strCardCode = oApplication.Utilities.getEdittextvalue(aForm, "5")
                oForm = oApplication.Utilities.LoadForm(xml_ItemGroup, frm_Itemgroup)
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                oForm.Freeze(True)
                databind(oForm, strCardCode)
                oForm.Freeze(False)
            End If
        End If
    End Sub
#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form, ByVal acode As String)
        Try

            aForm.Freeze(True)
            oGrid = aForm.Items.Item("6").Specific
            oGrid.DataTable.ExecuteQuery("Select * from [@DABT_OITB] where U_CardCode='" & acode & "' order by Convert(Numeric,Code)")
            oApplication.Utilities.setEdittextvalue(aForm, "4", acode)
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

        oGrid = aForm.Items.Item("9").Specific
    End Sub


#End Region





#End Region

#End Region

#Region "Events"
#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_PaymentMeans
                    If pVal.BeforeAction = False Then
                        '  LoadForm()
                    Else
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        frmSourcePaymentform = oForm
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        If oForm.TypeEx = frm_BPMaster Then
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                oForm.Items.Item("btnDisplay").Visible = False
                            Else
                                oCombo = oForm.Items.Item("40").Specific
                                If oCombo.Selected.Value = "C" Then
                                    oForm.Items.Item("btnDisplay").Visible = True
                                Else
                                    oForm.Items.Item("btnDisplay").Visible = False

                                End If

                            End If
                        End If
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
            If pVal.FormTypeEx = frm_PaymentMeans Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "28" And pVal.ColUID = "U_Z_CONTID" And pVal.CharPressed <> 9 Then
                                    ' BubbleEvent = False
                                    '  Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '  oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If frmSourcePaymentform.TypeEx = frm_IncomingPayment Then
                                    Dim oItems As SAPbouiCOM.Item
                                    oApplication.Utilities.AddControls(oForm, "st", "8", SAPbouiCOM.BoFormItemTypes.it_STATIC, "RIGHT", , , , "Tenant Code", 60)
                                    oApplication.Utilities.AddControls(oForm, "ed", "st", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", , , , , 60)
                                    oForm.DataSources.UserDataSources.Add("TenCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                                    oEditText = oForm.Items.Item("ed").Specific
                                    oEditText.DataBind.SetBound(True, "", "TenCode")
                                    oApplication.Utilities.setEdittextvalue(oForm, "ed", oApplication.Utilities.getEdittextvalue(frmSourcePaymentform, "5"))
                                    oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Items.Item("ed").Enabled = False
                                    oItems = oForm.Items.Item("st")
                                    oItems.LinkTo = "ed"

                                    oApplication.Utilities.AddControls(oForm, "stFrom", "50", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 2, 3, , "From Date")
                                    oApplication.Utilities.AddControls(oForm, "edFrom", "30", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 2, 3, )
                                    oItems = oForm.Items.Item("stFrom")
                                    oItems.LinkTo = "edFrom"

                                    oEditText = oForm.Items.Item("edFrom").Specific
                                    oEditText.DataBind.SetBound(True, "ORCT", "U_Z_FromDate")

                                    oApplication.Utilities.AddControls(oForm, "stTo", "stFrom", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 2, 3, , "To Date")
                                    oApplication.Utilities.AddControls(oForm, "edTo", "edFrom", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 2, 3, )
                                    oItems = oForm.Items.Item("stTo")
                                    oItems.LinkTo = "edTo"
                                    oEditText = oForm.Items.Item("edTo").Specific
                                    oEditText.DataBind.SetBound(True, "ORCT", "U_Z_ToDate")




                                End If
                              Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "28" And pVal.ColUID = "U_Z_CONTNUMBER" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim strIns As String
                                    strIns = oApplication.Utilities.getEdittextvalue(oForm, "ed")
                                    Dim otest As SAPbobsCOM.Recordset
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Select * from [@Z_CONTRACT] where U_Z_CONNO='" & oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_CONTNUMBER", pVal.Row) & "' and U_Z_TenCode='" & strIns & "'")
                                    If otest.RecordCount <= 0 Then
                                        oApplication.Utilities.SetMatrixValues(oMatrix, pVal.ColUID, pVal.Row, "")
                                        strIns = ""
                                    Else
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_CONTID", pVal.Row, otest.Fields.Item("DocEntry").Value)

                                        Exit Sub
                                    End If
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsContractCFL
                                    clsContractCFL.ItemUID = pVal.ItemUID
                                    clsContractCFL.SourceFormUID = FormUID
                                    clsContractCFL.SourceLabel = pVal.Row
                                    clsContractCFL.CFLChoice = "Contract" 'oCombo.Selected.Value
                                    clsContractCFL.choice = "Contract"
                                    clsContractCFL.Documentchoice = "" 'oApplication.Utilities.getEdittextvalue(oForm, "9") 'TenCode
                                    clsContractCFL.ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "ed") 'Unit Code
                                    ' clsChooseFromList.BinDescrUID = "BinToBinHeader"
                                    clsContractCFL.sourceColumID = pVal.ColUID
                                    oApplication.Utilities.LoadForm("\CFL1.xml", frm_ChoosefromList1)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "28" And pVal.ColUID = "U_Z_CONTNUMBER" And pVal.CharPressed = 9 Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim strIns As String
                                    strIns = oApplication.Utilities.getEdittextvalue(oForm, "ed")
                                    Dim otest As SAPbobsCOM.Recordset
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Select * from [@Z_CONTRACT] where U_Z_CONNO='" & oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_CONTNUMBER", pVal.Row) & "' and U_Z_TenCode='" & strIns & "'")
                                    If otest.RecordCount <= 0 Then
                                        oApplication.Utilities.SetMatrixValues(oMatrix, pVal.ColUID, pVal.Row, "")
                                        strIns = ""
                                    Else
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "U_Z_CONTID", pVal.Row, otest.Fields.Item("DocEntry").Value)

                                        Exit Sub
                                    End If
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsContractCFL
                                    clsContractCFL.ItemUID = pVal.ItemUID
                                    clsContractCFL.SourceFormUID = FormUID
                                    clsContractCFL.SourceLabel = pVal.Row
                                    clsContractCFL.CFLChoice = "Contract" 'oCombo.Selected.Value
                                    clsContractCFL.choice = "Contract"
                                    clsContractCFL.Documentchoice = "" 'oApplication.Utilities.getEdittextvalue(oForm, "9") 'TenCode
                                    clsContractCFL.ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "ed") 'Unit Code
                                    ' clsChooseFromList.BinDescrUID = "BinToBinHeader"
                                    clsContractCFL.sourceColumID = pVal.ColUID
                                    oApplication.Utilities.LoadForm("\CFL1.xml", frm_ChoosefromList1)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If
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
