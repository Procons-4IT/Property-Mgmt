Public Class clsIncomingPayment
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
    Private oItems, oItems1 As SAPbouiCOM.Item
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

    End Sub

    Private Sub addcontrol(ByVal aform As SAPbouiCOM.Form)
        oApplication.Utilities.AddControls(aform, "stDocEntry", "53", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "Contract ID")
        oApplication.Utilities.AddControls(aform, "edDocEntry", "52", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, "stDocEntry")
        oApplication.Utilities.AddControls(aform, "edDocLink", "edDocEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON, "LEFT", 0, 0, "stDocEntry")
        oItems = aform.Items.Item("stDocEntry")
        oItems.LinkTo = "edDocLink"
        oItems1 = aform.Items.Item("edDocLink")
        oItems1.LinkTo = "edDocEntry"
        oEditText = aform.Items.Item("edDocEntry").Specific
        oEditText.DataBind.SetBound(True, "ORCT", "U_Z_CONTID")


        oApplication.Utilities.AddControls(aform, "stCntNo", "stDocEntry", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "Contract Number")
        oApplication.Utilities.AddControls(aform, "edCntNo", "edDocEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, "stCntNo")
        oItems = aform.Items.Item("stCntNo")
        oItems.LinkTo = "edCntNo"

        oEditText = aform.Items.Item("edCntNo").Specific
        oEditText.DataBind.SetBound(True, "ORCT", "U_Z_CONTNUMBER")

        oApplication.Utilities.AddControls(aform, "stCnt1No", "stCntNo", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "Contract Document Number")
        oApplication.Utilities.AddControls(aform, "edCnt1No", "edCntNo", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, "stCnt1No")
        oItems = aform.Items.Item("stCnt1No")
        oItems.LinkTo = "edCnt1No"
        oEditText = aform.Items.Item("edCnt1No").Specific
        oEditText.DataBind.SetBound(True, "ORCT", "U_Z_CNTNUMBER")
        oApplication.Utilities.AddControls(aform, "stseqNo", "stCnt1No", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 0, 0, , "Contract Seq. Number")
        oApplication.Utilities.AddControls(aform, "edseqNo", "edCnt1No", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 0, 0, "stseqNo")
        oItems = aform.Items.Item("stseqNo")
        oItems.LinkTo = "edseqNo"
        oEditText = aform.Items.Item("edseqNo").Specific
        oEditText.DataBind.SetBound(True, "ORCT", "U_SEQ")

     
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
            If pVal.FormTypeEx = frm_Incomingpayment Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "edDocEntry" Or pVal.ItemUID = "edCnt1No" Or pVal.ItemUID = "edCntNo" Or pVal.ItemUID = "edseqNo" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '  oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If oForm.TypeEx = frm_Incomingpayment Then
                                    addcontrol(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "edDocLink" Then
                                    Dim oObj As New clsTenContracts
                                    oObj.LoadForm_Contract_View_Payment(oApplication.Utilities.getEdittextvalue(oForm, "edCnt1No"))
                                    Exit Sub

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Exit Sub
                                End If
                                If pVal.ItemUID = "edDocEntry" Or pVal.ItemUID = "edCntNo" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Dim strIns As String
                                    strIns = oApplication.Utilities.getEdittextvalue(oForm, "5")
                                    Dim otest As SAPbobsCOM.Recordset
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsContractCFL2
                                    clsContractCFL2.ItemUID = pVal.ItemUID
                                    clsContractCFL2.SourceFormUID = FormUID
                                    clsContractCFL2.SourceLabel = pVal.Row
                                    clsContractCFL2.CFLChoice = "Contract" 'oCombo.Selected.Value
                                    clsContractCFL2.choice = "Contract"
                                    clsContractCFL2.Documentchoice = "" 'oApplication.Utilities.getEdittextvalue(oForm, "9") 'TenCode
                                    clsContractCFL2.ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "5") 'Unit Code
                                    ' clsChooseFromList.BinDescrUID = "BinToBinHeader"
                                    clsContractCFL2.sourceColumID = 0 ' pVal.ColUID
                                    oApplication.Utilities.LoadForm("\CFL2.xml", frm_ChoosefromList1)
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
