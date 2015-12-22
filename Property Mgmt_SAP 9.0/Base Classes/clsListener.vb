Public Class clsListener
    Inherits Object
    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter

#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property
#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SetFilter(Filters)
    End Sub
    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Menu Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent

        If BusinessObjectInfo.BeforeAction = False Then
            Select Case BusinessObjectInfo.FormTypeEx
                'Case frm_InvoicePayment, frm_Delivery, frm_Invoice, frm_Return, frm_ARCreditMemo, frm_GRPO, frm_APInvoice, frm_GoodsReturn, frm_APCreditMemo
                '    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_TermTransaction
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsTermTransaction
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case frm_ApproveTemp
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsApproveTemp
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_PropertyData
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsPropertyData
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_PropertyUnitDetails
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsPropertyUnitDetails
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_Evaluation
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsPropertyEvalution
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_Reservation
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsReservation
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_Contracts
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsContracts
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_TenContracts
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsTenContracts
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select
        Else
            Select Case BusinessObjectInfo.FormTypeEx
                'Case frm_InvoicePayment, frm_Delivery, frm_Invoice, frm_Return, frm_ARCreditMemo, frm_GRPO, frm_APInvoice, frm_GoodsReturn, frm_APCreditMemo
                '    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case frm_PropertyData
                    If Not _Collection.ContainsKey(_FormUID) Then
                        oItemObject = New clsPropertyData
                        oItemObject.FrmUID = _FormUID
                        _Collection.Add(_FormUID, oItemObject)
                    End If
                    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case frm_PropertyUnitDetails
                    '    If Not _Collection.ContainsKey(_FormUID) Then
                    '        oItemObject = New clsPropertyUnitDetails
                    '        oItemObject.FrmUID = _FormUID
                    '        _Collection.Add(_FormUID, oItemObject)
                    '    End If
                    '    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case frm_Reservation
                    '    If Not _Collection.ContainsKey(_FormUID) Then
                    '        oItemObject = New clsReservation
                    '        oItemObject.FrmUID = _FormUID
                    '        _Collection.Add(_FormUID, oItemObject)
                    '    End If
                    '    _Collection.Item(_FormUID).FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select
        End If
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_ContractApproval
                        oMenuObject = New clsContractApproval
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_TermApproval
                        oMenuObject = New clsTerminationApproval
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_TermTransaction
                        oMenuObject = New clsTermTransaction
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_ApproveTemp
                        oMenuObject = New clsApproveTemp
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ProFac
                        oMenuObject = New clsProFacMaster
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ProUFac
                        oMenuObject = New clsProUFacMaster
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Report
                        oMenuObject = New clsReportWizard
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Propertydata
                        oMenuObject = New clsPropertyData
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_PaymentMeans
                        oMenuObject = New clsPaymentmeans
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Renewal
                        oMenuObject = New clsRenewal
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_postingWizard
                        oMenuObject = New clsPostingWizard
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Contracts, "InvList", "Agreement"
                        oMenuObject = New clsContracts
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_TenContracts, "InvList1", "Agreement1", "Renewal", "CheckList", "TerApp", "ConApp"
                        oMenuObject = New clsTenContracts
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Insurance
                        oMenuObject = New clsInsurance
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Evalution
                        oMenuObject = New clsPropertyEvalution
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Earning
                        oMenuObject = New clsEarning
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Search
                        oMenuObject = New clsSearch
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_PropertyType
                        oMenuObject = New clspropertyType
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_PropertyUnitSetup
                        oMenuObject = New clsPropertyUnitDetails
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_PropertyUnitType
                        oMenuObject = New clsPropertyUnitType
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Reservation
                        oMenuObject = New clsReservation
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_PriceList
                        oMenuObject = New clsPriceList
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_BillGeneration
                        oMenuObject = New clsBillGeneration
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_Location
                        oMenuObject = New clsLocation
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_PaymentMeans, mnu_ADD, mnu_FIND, mnu_FIRST, mnu_Remove, mnu_NEXT, mnu_PREVIOUS, mnu_LAST, mnu_ADD_ROW, mnu_DELETE_ROW
                        Dim aForm As SAPbouiCOM.Form
                        aForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If _Collection.ContainsKey(aForm.UniqueID) Then
                            oItemObject = _Collection.Item(aForm.UniqueID)
                            _Collection.Item(aForm.UniqueID).menuevent(pVal, BubbleEvent)
                        End If
                    Case "Reserve", "Contract", "Billing", "1287"
                        oMenuObject = New clsPropertyUnitDetails
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                End Select
            Else
                Select Case pVal.MenuUID
                    Case mnu_DELETE_ROW, mnu_Remove, "1287", mnu_ADD_ROW
                        Dim aForm As SAPbouiCOM.Form
                        aForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If _Collection.ContainsKey(aForm.UniqueID) Then
                            oItemObject = _Collection.Item(aForm.UniqueID)
                            _Collection.Item(aForm.UniqueID).menuevent(pVal, BubbleEvent)
                        End If
                    Case mnu_PaymentMeans
                        oMenuObject = New clsPaymentmeans
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                End Select
            End If


        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub
#End Region

#Region "Item Event"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormTypeEx
                    Case frm_ItemMaster
                        If pVal.Before_Action = False Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            If oForm.TypeEx = frm_ItemMaster Then
                                Try
                                    oForm.Items.Item("chkPro").Enabled = False
                                Catch ex As Exception

                                End Try

                            End If
                        End If
                End Select
            End If
            If pVal.FormTypeEx = "frm_Check" And pVal.ItemUID = "6" And pVal.Before_Action = False Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    Dim oForm As SAPbouiCOM.Form
                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                    oApplication.Utilities.LoadCheckDetails(oForm)
                End If
            End If
            If pVal.FormTypeEx = "frm_Check" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                Dim oForm As SAPbouiCOM.Form
                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                Dim oCFL As SAPbouiCOM.ChooseFromList
                Dim sCHFL_ID As String

                Dim intChoice As Integer
                Dim codebar, val2, val As String
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
                        If pVal.ItemUID = "3" Or pVal.ItemUID = "5" Then
                            val2 = oDataTable.GetValue("U_Z_CNTNO", 0)
                            val = oDataTable.GetValue("DocEntry", 0)
                            If pVal.ItemUID = "3" Then
                                oApplication.Utilities.setEdittextvalue(oForm, "8", val)
                            Else
                                oApplication.Utilities.setEdittextvalue(oForm, "9", val)
                            End If
                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val2)
                        End If
                        oForm.Freeze(False)
                    End If
                Catch ex As Exception
                    oForm.Freeze(False)
                End Try
            End If

            If pVal.Before_Action = False And pVal.FormTypeEx = frm_BPMaster And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                Dim oForm As SAPbouiCOM.Form
                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                If oForm.TypeEx = frm_BPMaster Then

                    Try
                        Dim oItem As SAPbouiCOM.Item
                        Dim ofolder As SAPbouiCOM.Folder
                        Dim oEditText As SAPbouiCOM.EditText
                        oApplication.Utilities.AddControls(oForm, "stNation", "358", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 1, 1, , "Nationality")
                        oApplication.Utilities.AddControls(oForm, "edNation", "362", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 1, 1)
                        oEditText = oForm.Items.Item("edNation").Specific
                        oEditText.DataBind.SetBound(True, "OCRD", "U_Nationality")

                        oItem = oForm.Items.Item("stNation")
                        oItem.LinkTo = "edNation"

                        oApplication.Utilities.AddControls(oForm, "stOccp", "343", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 1, 1, , "Occupation")
                        oApplication.Utilities.AddControls(oForm, "edOccp", "345", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 1, 1)

                        oEditText = oForm.Items.Item("edOccp").Specific
                        oEditText.DataBind.SetBound(True, "OCRD", "U_Occupation")

                        oItem = oForm.Items.Item("stOccp")
                        oItem.LinkTo = "edOccp"



                        oApplication.Utilities.AddControls(oForm, "stMat", "stNation", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , 1, 1, "Marital Status")
                        oApplication.Utilities.AddControls(oForm, "edMat", "edNation", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 1, 1)
                        oEditText = oForm.Items.Item("edMat").Specific
                        oEditText.DataBind.SetBound(True, "OCRD", "U_MaritalStatus")

                        oItem = oForm.Items.Item("stMat")
                        oItem.LinkTo = "edMat"

                        oApplication.Utilities.AddControls(oForm, "stRegdate", "stOccp", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , 1, 1, "Registration Date")
                        oApplication.Utilities.AddControls(oForm, "edRegdate", "edOccp", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 1, 1)

                        oEditText = oForm.Items.Item("edRegdate").Specific
                        oEditText.DataBind.SetBound(True, "OCRD", "U_RegDate")

                        oItem = oForm.Items.Item("stRegdate")
                        oItem.LinkTo = "edRegdate"
                    Catch ex As Exception

                    End Try

                End If

            End If


            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
                Select Case pVal.FormType
                    Case "0"
                        Dim oform As SAPbouiCOM.Form
                        oform = oApplication.SBO_Application.Forms.Item(FormUID)
                        If oform.TypeEx = "0" Then
                            Try
                                Dim ostatic As SAPbouiCOM.StaticText
                                ostatic = oform.Items.Item("4").Specific
                                'MsgBox(ostatic.Caption)
                                ' you want to save the changes?
                                ' Do you want to save the changes?
                                If ostatic.Caption.Contains("save") And ostatic.Caption.Contains("changes") = True Then
                                    oform.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                End Select
            ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = True Then
                Select Case pVal.FormType
                    Case "0"
                        Dim oform As SAPbouiCOM.Form
                        oform = oApplication.SBO_Application.Forms.Item(FormUID)
                        If oform.TypeEx = frm_PropertyData Then
                            Try
                                'Dim obj As New clsPropertyData
                                'If obj.AddtoUDT(oform) = False Then
                                '    BubbleEvent = False
                                '    Exit Sub
                                'End If
                                'Dim ostatic As SAPbouiCOM.StaticText
                                'ostatic = oform.Items.Item("4").Specific
                                ''MsgBox(ostatic.Caption)
                                '' you want to save the changes?
                                '' Do you want to save the changes?
                                'If ostatic.Caption.Contains("save") And ostatic.Caption.Contains("changes") = True Then
                                '    oform.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                'End If
                            Catch ex As Exception
                            End Try
                        End If
                End Select
            End If


            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
                Select Case pVal.FormTypeEx
                    Case frm_ItemMaster
                        If pVal.Before_Action = False Then
                            Dim oForm As SAPbouiCOM.Form
                            Dim oCheckbox As SAPbouiCOM.CheckBox
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            If oForm.TypeEx = frm_ItemMaster Then
                                oApplication.Utilities.AddControls(oForm, "chkPro", "42", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, "DOWN", , , "42", "Property Item")
                                oCheckbox = oForm.Items.Item("chkPro").Specific
                                oCheckbox.DataBind.SetBound(True, "OITM", "U_Z_PROFLG")
                                oForm.Items.Item("chkPro").Enabled = False
                            End If
                        End If
                    Case frm_AppHisDetails
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsAppHisDetails
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_TermTransaction
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsTermTransaction
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ContractApproval
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsContractApproval
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_TermApproval
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsTerminationApproval
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ApproveTemp
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsApproveTemp
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ProFac
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProFacMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ProUFac
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsProUFacMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Installments
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInstallment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_IncomingPayment
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsIncomingPayment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_Report
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsReportWizard
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_DocumentView
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocumentsView
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_RenewalHistory
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsRenewalHistory
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_PaymentMeans
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPaymentmeans
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_renewal
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsRenewal
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_PostingWizard
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPostingWizard
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Evaluation
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPropertyEvalution
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ChoosefromList
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsChooseFromList
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ChoosefromList1
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsContractCFL
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ChoosefromList2
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsContractCFL2
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_UnitReport
                        oItemObject = New clsPropertyUnitReport
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Search
                        oItemObject = New clsSearch
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Location
                        oItemObject = New clsLocation
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_PropertyData
                        oItemObject = New clsPropertyData
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Earning
                        oItemObject = New clsEarning
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)

                    Case frm_PropertyType

                        oItemObject = New clspropertyType
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)

                    Case frm_PriceList
                        oItemObject = New clsPriceList
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Contracts
                        oItemObject = New clsContracts
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_TenContracts
                        oItemObject = New clsTenContracts
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Insurance
                        oItemObject = New clsInsurance
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)

                    Case frm_PropertyUnitType
                        oItemObject = New clsPropertyUnitType
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_PropertyUnitDetails
                        oItemObject = New clsPropertyUnitDetails
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_Reservation
                        oItemObject = New clsReservation
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                    Case frm_BillGeneration
                        oItemObject = New clsBillGeneration
                        oItemObject.FrmUID = FormUID
                        _Collection.Add(FormUID, oItemObject)
                End Select
            End If
            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Application Event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                    oApplication.Utilities.Message("Property Mgmt addon disconnected successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub
#End Region

#Region "Close Application"
    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Set Application"
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

    Private Sub _SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.RightClickEvent
        Try


            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            'If eventInfo.FormUID = "RightClk" Then
            If oForm.TypeEx = frm_PropertyUnitDetails Then
                oMenuObject = New clsPropertyUnitDetails
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            End If
            If oForm.TypeEx = frm_Contracts Then
                oItemObject = _Collection.Item(eventInfo.FormUID)
                _Collection.Item(eventInfo.FormUID).RightClickEvent(eventInfo, BubbleEvent)
            End If

            If oForm.TypeEx = frm_TenContracts Then
                oItemObject = _Collection.Item(eventInfo.FormUID)
                _Collection.Item(eventInfo.FormUID).RightClickEvent(eventInfo, BubbleEvent)
            End If
           
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
End Class

Public Class WindowWrapper
    Implements System.Windows.Forms.IWin32Window
    Private _hwnd As IntPtr

    Public Sub New(ByVal handle As IntPtr)
        _hwnd = handle
    End Sub

    Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property

End Class
