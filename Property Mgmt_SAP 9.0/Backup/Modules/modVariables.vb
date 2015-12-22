Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public sSearchList As String

    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public blnMasterExport As Boolean = False
    Public blnFEExport As Boolean = False
    Public blnDocumentItem As Boolean = False
    Public frm_SPBP, frmSourceForm, aSourceForm, frm_ScaleDiscountForm, frm_FilterTableForm, rm_SourceForm As SAPbouiCOM.Form
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public strItemSelectionQuery As String = ""
    Public frmSourcePaymentform As SAPbouiCOM.Form

    Public intSelectedMatrixrow As Integer = 0
    Public strSPBPCode As String = ""
    Public strSPItemCode As String = ""
    Public ErrorLogFile As String

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum

    Public Const xml_Installments As String = "frm_Installment.xml"
    Public Const frm_Installments As String = "frm_Installment"

    Public Const mnu_Report As String = "Z_Mnu_Report"
    Public Const xml_Report As String = "frm_Report.xml"
    Public Const frm_Report As String = "frm_Report"


    Public Const mnu_Renewal As String = "Z_Mnu_Renewal"
    Public Const xml_renewal As String = "xml_renewal.xml"
    Public Const frm_renewal As String = "frm_Renewal"

    Public Const frm_PostingWizard As String = "frm_Posting"
    Public Const xml_Postingwizard As String = "frm_Postingwizard.xml"
    Public Const mnu_postingWizard As String = "Z_Mnu_Posting"

    Public Const frm_PropertyUnitType As String = "frm_UnitType"
    Public Const frm_PropertyUnitDetails As String = "frm_UnitDetails"
    Public Const frm_Reservation As String = "frm_Reserve"
    Public Const frm_PropertyData As String = "frm_ProData"
    Public Const frm_Contracts As String = "frm_Contracts"
    Public Const frm_TenContracts As String = "frm_TenContracts"
    Public Const frm_Insurance As String = "frm_Insur"
    Public Const frm_PriceList As String = "frm_PriceList"
    Public Const frm_DocumentView As String = "frm_Documents"
    Public Const frm_RenewalHistory As String = "frm_RenewalHistory"

    Public Const frm_Itemgroup As String = "frm_ItemGroup"
    Public Const frm_Earning As String = "frm_Earning"
    Public Const frm_PropertyType As String = "frm_PropertyType"
    Public Const frm_BillGeneration As String = "frm_Bill"
    Public Const frm_UnitReport As String = "frm_UnitReport"
    Public Const frm_Location As String = "frm_Location"
    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_Search As String = "frm_Search"
    Public Const frm_ChoosefromList As String = "frm_CFL"
    Public Const frm_ChoosefromList1 As String = "frm_CFL1"
    Public Const frm_ChoosefromList2 As String = "frm_CFL2"
    Public Const frm_Evaluation As String = "frm_Evalution"

    Public Const frm_SalesQuotation As Integer = 149

    Public Const frm_SalesOrder As Integer = 139
    Public Const frm_Delivery As Integer = 140
    Public Const frm_Return As Integer = 180
    Public Const frm_Downpaymentrequest As Integer = 65308
    Public Const frm_DownpaymentInvoice As Integer = 65300
    Public Const frm_InvoicePayment As Integer = 60090
    Public Const frm_reverseInvoice As Integer = 60091



    Public Const frm_ARInvoice As Integer = 133
    Public Const frm_ARCreditNote As Integer = 179

    Public Const mnu_ItemGroup As String = "Z_Itms01"
  
    Public Const frm_ItemMaster As String = "150"
    Public Const frm_BPMaster As String = "134"

    Public Const mnu_PaymentMeans As String = "5892"
    Public Const frm_IncomingPayment As String = "170"
    Public Const frm_PaymentMeans As String = "146"

    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_Remove As String = "1283"
    Public Const mnu_Cancel As String = "1284"

    Public Const mnu_PropertyUnitType As String = "Z_Menu_103"
    Public Const mnu_Propertydata As String = "Z_Menu_104"
    Public Const mnu_PropertyUnitSetup As String = "Z_Menu_105"
    Public Const mnu_Reservation As String = "Z_Menu_106"
    Public Const mnu_Contracts As String = "Z_Menu_107"
    Public Const mnu_TenContracts As String = "Z_Menu_117"
    Public Const mnu_Insurance As String = "Z_Menu_108"
    ' Public Const frm_ChoosefromList As String = "d"
    Public Const mnu_PriceList As String = "Z_Menu_110"
    Public Const mnu_Search As String = "Z_Menu_109"
    Public Const mnu_Earning As String = "Z_Menu_111"
    Public Const mnu_PropertyType As String = "Z_Menu_112"
    Public Const mnu_BillGeneration As String = "Z_Menu_113"
    Public Const mnu_Location As String = "Z_Menu_123"
    Public Const XML_CFL As String = "CFL.xml"
    Public Const mnu_Evalution As String = "mnu_Evalution"


    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_ItemGroup As String = "xml_ItemGroup.xml"
    Public Const xml_PropertyUnitType As String = "xml_PropertyUnitType.xml"
    Public Const xml_PropertyData As String = "xml_PropertyData.xml"
    Public Const xml_PropertyUnitSetup As String = "xml_PropertyUnitDetails.xml"
    Public Const xml_Reservation As String = "xml_Reservation.xml"
    Public Const xml_Contracts As String = "xml_Contracts.xml"
    Public Const xml_TenContracts As String = "xml_Ten_Contracts.xml"
    Public Const xml_Insurance As String = "xml_Insurance.xml"
    Public Const xml_PriceList As String = "xml_PriceList.xml"
    Public Const xml_Earning As String = "frm_Earning.xml"
    Public Const xml_PropertyType As String = "frm_PropertyType.xml"
    Public Const xml_BillGeneration As String = "frm_BillGeneration.xml"
    Public Const xml_UnitReport As String = "frm_Unitreport.xml"
    Public Const xml_Location As String = "frm_Location.xml"
    Public Const xml_Search As String = "xml_Search.xml"
    Public Const xml_Docuemntsview As String = "frm_Documents.xml"
    Public Const xml_RenewalHistory As String = "frm_RenewalHistory.xml"

    Public Const xml_Evalution As String = "frm_Evalution.xml"

End Module
