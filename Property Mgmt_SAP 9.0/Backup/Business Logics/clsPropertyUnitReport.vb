Public Class clsPropertyUnitReport
    Inherits clsBase

    Public Shared ItemUID As String
    Public Shared ItemCode As String
    Public Shared Choice As String

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
    Public Sub LoadForm(ByVal strItemCode As String, ByVal strChoice As String)
        oForm = oApplication.Utilities.LoadForm(xml_UnitReport, frm_UnitReport)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        ItemCode = strItemCode
        Choice = strChoice
        databind(oForm)
        Select Case strChoice
            Case "Billing"
                oForm.Title = "Property Unit : " & strItemCode & " - Billing details"
            Case "Contract"
                oForm.Title = "Property Unit : " & strItemCode & " - Contract details"
            Case "Reserve"
                oForm.Title = "Property Unit : " & strItemCode & " - Reservation details"
        End Select
    End Sub
    Private Sub databind(ByVal oform As SAPbouiCOM.Form)
        oGrid = oform.Items.Item("1").Specific
        Dim strsql As String
        If Choice = "Reserve" Then
            strsql = "select DocEntry 'Reservation ID',CreateDate 'Reservation Date',U_Z_PropCode 'Property Code',U_Z_PropDesc 'Description'"
            strsql = strsql & ",U_Z_CARDCODE 'Customer Code' ,U_Z_CARDNAME 'Customer Name',U_Z_Address 'Address',U_Z_AgentName 'AgentName',U_Z_EMPName 'EmployeeName',"
            strsql = strsql & " U_Z_StartDate 'Start Date',U_Z_EndDate 'EndDate',U_Z_DOWNAMOUNT 'Down Payment Amount'  from [@Z_RESER]"
            strsql = strsql & " where U_Z_UNITCODE='" & ItemCode & "'"

            'oGrid.DataTable.ExecuteQuery("Select * from [@Z_RESER] where U_Z_UNITCODE='" & ItemCode & "'")
        End If
        If Choice = "Contract" Then
            strsql = "select DocEntry 'Contract ID', U_Z_ConNo 'Contract Number',CreateDate 'Contract Date',case U_Z_Type when 'O' then 'Owner' else 'Tenant' end 'Contract Type',U_Z_STATUS 'Status'"
            strsql = strsql & ",U_Z_TENCODE 'Tenent Code' ,U_Z_TENNAME 'Tenent Name',U_Z_OFFADDRESS ' Office Address',U_Z_Annualrent 'Annual Rent',"
            strsql = strsql & " U_Z_Deposit 'Deposit',U_Z_ChgMonth 'Number of Month',"
            strsql = strsql & " U_Z_StartDate 'Start Date',U_Z_EndDate 'EndDate'  from [@Z_CONTRACT]"
            strsql = strsql & " where U_Z_UNITCODE='" & ItemCode & "'"
            'oGrid.DataTable.ExecuteQuery("Select * from [@Z_CONTRACT] where U_Z_UNITCODE='" & ItemCode & "'")
        End If
        If Choice = "Billing" Then
            strsql = "Select  U_Year 'Year',DATENAME(m, str(U_Month) + '/1/2011') 'Month',U_ContractId 'Contract ID',U_UNITCode 'Unit ID',U_CARDCODE 'Tenent Code',U_MONTHRENT"
            strsql = strsql & " 'Monthly rental',U_Expenses 'Monthly Expenses',U_TOTAL 'Total Amount',case  U_Invoiced when 'Y' then 'Yes' else 'No' end 'Invoiced' from [@Z_OBILL]"
            strsql = strsql & " where U_UNITCODE='" & ItemCode & "'"
            ' oGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where U_UNITCODE='" & ItemCode & "'")
        End If
        oGrid.DataTable.ExecuteQuery(strsql)

    End Sub
#End Region

    Private Sub FormatGrid(ByVal aChoice As String, ByVal aGrid As SAPbouiCOM.Grid)



    End Sub

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
            If pVal.FormTypeEx = frm_UnitReport Then
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
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID

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
