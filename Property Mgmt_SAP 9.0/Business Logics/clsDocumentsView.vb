Public Class clsDocumentsView
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Public Sub LoadForm(ByVal acode As String, ByVal aChoice As String)
        oForm = oApplication.Utilities.LoadForm(xml_Docuemntsview, frm_DocumentView)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.DataSources.UserDataSources.Add("DocN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("DocN1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        'oEditText = oForm.Items.Item("10").Specific
        'oEditText.DataBind.SetBound(True, "", "DocN")
        'oEditText = oForm.Items.Item("12").Specific
        'oEditText.DataBind.SetBound(True, "", "DocN1")
        'oEditText.ChooseFromListUID = "Z_ORDR"
        'oEditText.ChooseFromListAlias = "DocEntry"
        Try
            oForm.DataSources.DataTables.Add("oMatrixDT")
            oForm.DataSources.DataTables.Item("oMatrixDT").Clear()
        Catch ex As Exception
            oForm.DataSources.DataTables.Item("oMatrixDT").Clear()
        End Try
        oGrid = oForm.Items.Item("1").Specific
        oMatrix = oForm.Items.Item("11").Specific
        Dim oColumn As SAPbouiCOM.Column
        oColumn = oMatrix.Columns.Item(0)
        Dim oTest As SAPbobsCOM.Recordset
        Dim strContractNumber, strString, strSql, strTenCode, cardName As String
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select * from [@Z_CONTRACT] where DocEntry=" & acode)
        strContractNumber = oTest.Fields.Item("U_Z_Conno").Value
        strContractNumber = oTest.Fields.Item("U_Z_CntNo").Value
        strTenCode = oTest.Fields.Item("U_Z_TenCode").Value
        cardName = oTest.Fields.Item("U_Z_TEnName").Value
        Dim dblDepositedamount, dblInvoicedAmount, dblOwnerAmount As Double
        Dim strsql1, strsql2, strSql3, strsql4, strCreditNote As String
        If aChoice = "Booking" Then
            '   oApplication.Utilities.setEdittextvalue(oForm, "4", acode)
            '  oApplication.Utilities.setEdittextvalue(oForm, "7", acode)
            'strString = "Select U_Z_ContID, U_Z_ContNumber 'ContractNumber','DP'  'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from OINV where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " Union All Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate,sum(t1.LineTotal) 'Doctotal' from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "') group by U_Z_ContID,U_Z_ContNumber ,T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate "
            'strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"

            'strString = strString & " union all Select T1. U_Z_ContID ,T1.U_Z_ContNumber 'ContractNumber','RC' 'TransType',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocDate,T1.CheckSum from ORCT T0  INNER JOIN RCT1 T1 ON T0.DocEntry = T1.DocNum where (isnull(T1.U_Z_ContID,'')=" & acode & " or T1.U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','DE' 'TransType',DeposId 'DocEntry',DeposNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',DeposDate 'DocDate',LocTotal 'DocTotal' from ODPS  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','JE' 'TransType',TransId 'DocEntry',BatchNum 'DocNum',ShortName 'CardCode','" & cardName & "' 'CardName',DueDate 'DocDate',(Debit-Credit) 'DocTotal' from JDT1  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,Transref 'ContractNumber','CFP' 'TransType',CheckKey 'DocEntry',CheckNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',PmntDate 'DocDate',CheckSum 'DocTotal' from OCHO  where ( Transref='" & strContractNumber & "')"


            'strString = "Select U_Z_ContID, U_Z_ContNumber 'ContractNumber','DP'  'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where (isnull(U_Z_ContID,'')=" & acode & " )"
            '' strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from OINV where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " Union All Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate,sum(t1.LineTotal) 'Doctotal' from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where  (isnull(U_Z_ContID,'')=" & acode & ") group by U_Z_ContID,U_Z_ContNumber ,T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate "
            'strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "') and CheckSum<=0"
            'strString = strString & " union all Select T1. U_Z_ContID ,T1.U_Z_ContNumber 'ContractNumber','RC' 'TransType',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocDate,T1.CheckSum from ORCT T0  INNER JOIN RCT1 T1 ON T0.DocEntry = T1.DocNum where (isnull(T1.U_Z_ContID,'')=" & acode & ")"
            '' strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','DE' 'TransType',DeposId 'DocEntry',DeposNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',DeposDate 'DocDate',LocTotal 'DocTotal' from ODPS  where (U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " Union All SELECT T1.[U_Z_CONTID], T1.[U_Z_CONTNUMBER]'ContractNumber','DE' 'TransType', T2.[DeposNum] 'DocEntry', T2.[DeposNum] 'DocNum', T0.[CardCode],'" & cardName & "' 'CardName', T2.[DeposDate], T1.[CheckSum] 'DocTotal' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId where (isnull(T1.U_Z_ContID,'')=" & acode & ")"
            'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','JE' 'TransType',TransId 'DocEntry',BatchNum 'DocNum',ShortName 'CardCode','" & cardName & "' 'CardName',DueDate 'DocDate',(Debit-Credit) 'DocTotal' from JDT1  where (U_Z_ContNumber='" & strContractNumber & "')"
            'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,Transref 'ContractNumber','CFP' 'TransType',CheckKey 'DocEntry',CheckNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',PmntDate 'DocDate',CheckSum 'DocTotal' from OCHO  where ( Transref='" & strContractNumber & "')"

            strString = "Select U_Z_ContID, U_Z_CntNumber 'ContractNumber','DP'  'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where (isnull(U_Z_ContID,'')=" & acode & " )"
            ' strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from OINV where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            strString = strString & " Union All Select U_Z_ContID,U_Z_CntNumber 'ContractNumber','IN' 'TransType',T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate,sum(t1.LineTotal) 'Doctotal' from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where  (isnull(U_Z_ContID,'')=" & acode & ") group by U_Z_ContID,U_Z_CntNumber ,T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate "
            strString = strString & " union all Select U_Z_ContID,U_Z_CntNumber 'ContractNumber','CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_CntNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "') and CheckSum<=0"
            strString = strString & " union all Select T1. U_Z_ContID ,T1.U_Z_CntNumber 'ContractNumber','RC' 'TransType',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocDate,T1.CheckSum from ORCT T0  INNER JOIN RCT1 T1 ON T0.DocEntry = T1.DocNum where (isnull(T1.U_Z_ContID,'')=" & acode & ")"
            ' strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','DE' 'TransType',DeposId 'DocEntry',DeposNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',DeposDate 'DocDate',LocTotal 'DocTotal' from ODPS  where (U_Z_ContNumber='" & strContractNumber & "')"
            strString = strString & " Union All SELECT T1.[U_Z_CONTID], T1.[U_Z_CntNumber]'ContractNumber','DE' 'TransType', T2.[DeposNum] 'DocEntry', T2.[DeposNum] 'DocNum', T0.[CardCode],'" & cardName & "' 'CardName', T2.[DeposDate], T1.[CheckSum] 'DocTotal' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId where (isnull(T1.U_Z_ContID,'')=" & acode & ")"
            strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_CntNumber 'ContractNumber','JE' 'TransType',TransId 'DocEntry',BatchNum 'DocNum',ShortName 'CardCode','" & cardName & "' 'CardName',DueDate 'DocDate',(Debit-Credit) 'DocTotal' from JDT1  where (U_Z_ContNumber='" & strContractNumber & "')"
            strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,Transref 'ContractNumber','CFP' 'TransType',CheckKey 'DocEntry',CheckNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',PmntDate 'DocDate',CheckSum 'DocTotal' from OCHO  where ( Transref='" & strContractNumber & "')"


            strsql1 = "Select  sum(t1.LineTotal) 'Doctotal',0 from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(U_Z_ContID,'')=" & acode & " ) and isnull(U_Z_InvType,'T')='T'"
            'strSql3 = "Select  sum(t1.LineTotal) 'Doctotal',0 from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "') and isnull(U_Z_InvType,'T')='O'"
            strCreditNote = "Select  sum(t1.LineTotal) 'Doctotal',0 from ORIN T0 inner Join RIN1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(U_Z_ContID,'')=" & acode & " ) and isnull(U_Z_InvType,'T')='T'"

            strSql3 = " Select 0,sum(CheckSum) 'DocTotal1' from OCHO  where (Transref='" & strContractNumber & "')"
            strsql2 = "SELECT 0, sum(T1.[CheckSum]) 'DocTotal1' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId where (isnull(T1.U_Z_ContID,'')=" & acode & ")"

            ' strsql2 = " Select 0 'x',sum(LocTotal) 'DocTotal1' from ODPS  T0 inner Join [@Z_CONTRACT] T4 on T4.U_Z_CONNO =T0.U_Z_CONTNUMBER  where (isnull(T4.U_Z_ContID,'')=" & acode & ")"
            strsql2 = " SELECT 0 'x', Sum(T1.[CheckSum]) 'DocTotal1' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId inner Join [@Z_CONTRACT] T4 on T4.DocEntry   =T1.U_Z_CONTID  where (isnull(T4.DocEntry,'')=" & acode & ") and (isnull(T1.U_Z_ContID,'')<>'')"
            strsql2 = "Select Sum(X.x),SUm(X.DocTotal1) from (" & strsql2 & ") X"

            'strsql2 = " Select 0,sum(LocTotal) 'DocTotal1' from ODPS  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            strsql4 = " Select 0,sum(cashsum+creditsum+trsfrsum)  'DocTotal1' from ORCT  where checksum<=0 and (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
            'SELECT T0.DeposId,T0.[DeposType], T0.[DeposNum], T0.[DeposDate],* FROM ODPS T0
        ElseIf aChoice = "Activity" Then
            '   strString = "SELECT T0.[ClgCode] 'Activity Number', 'ACT' 'TransType', isnull(T0.U_BookRef,'') 'Booking Number', T0.[CardCode] 'Customer Code', T0.[Recontact] 'Activity Date',T0.[SeStartDat] 'Start Date', T0.[BeginTime] 'Start Time', T0.[endDate] 'End Date', T0.[ENDTime] 'End Time', case DurType when 'D' then (T0.[Duration]*24) when 'M' then  T0.Duration/60 else T0.Duration end 'Duration in Hours', T0.[Details] 'Remarks' FROM OCLG T0 where isnull(T0.U_BookRef,'')=" & acode
            strString = "SELECT T0.[ClgCode] 'Activity Number', 'ACT' 'TransType', isnull(T0.U_BookRef,'') 'Booking Number', T0.[CardCode] 'Customer Code', T0.[Recontact] 'Activity Date',T0.[SeStartDat] 'Start Date', T0.[BeginTime] 'Start Time', T0.[endDate] 'End Date', T0.[ENDTime] 'End Time', T0.U_Z_Duration  'Duration in  Hours', T0.[Details] 'Remarks' FROM OCLG T0 where isnull(T0.U_BookRef,'')=" & acode
        Else
            oApplication.Utilities.setEdittextvalue(oForm, "4", acode)
            oApplication.Utilities.setEdittextvalue(oForm, "7", acode)
            strString = "Select U_Z_ContID,'DP' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where isnull(CardCode,'')='" & acode & "' and isnull(U_Z_ContID,'')<>''"
            strString = strString & " union all Select U_Z_ContID,'IN' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from OINV where isnull(CardCode,'')='" & acode & "' and isnull(U_Z_ContID,'')<>''"
            strString = strString & " union all Select U_Z_ContID,'CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where isnull(U_BookRef,'')=" & acode & "' and isnull(U_Z_ContID,'')<>''"
            strString = strString & " union all Select CounterRef 'U_Z_ContID' ,'RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where isnull(CardCode,'')='" & acode & "' and isnull(CounterRef,'')<>''"
        End If

        If aChoice <> "Activity" Then
            strString = "Select X.U_Z_ContID,X.ContractNumber,X.TransType,X.DocEntry,X.DocNum,X.CardCode,X.CardName,X.DocDate,X.DocTotal from (" & strString & ") x order by X.TransType,X.U_Z_ContID,X.CardCode,X.DocDate"
            oGrid.DataTable.ExecuteQuery(strString)
            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery(strString)
            oColumn = oMatrix.Columns.Item("V_10")
            oColumn.DataBind.Bind("oMatrixDT", "U_Z_ContID")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oColumn = oMatrix.Columns.Item("V_1")
            oColumn.DataBind.Bind("oMatrixDT", "ContractNumber")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oColumn = oMatrix.Columns.Item("V_2")
            oColumn.DataBind.Bind("oMatrixDT", "TransType")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oColumn = oMatrix.Columns.Item("V_3")
            oColumn.DataBind.Bind("oMatrixDT", "DocEntry")
            '  oColumn.Type = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oColumn = oMatrix.Columns.Item("V_4")
            oColumn.DataBind.Bind("oMatrixDT", "DocNum")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oColumn = oMatrix.Columns.Item("V_5")
            oColumn.DataBind.Bind("oMatrixDT", "CardCode")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

            oColumn = oMatrix.Columns.Item("V_6")
            oColumn.DataBind.Bind("oMatrixDT", "CardName")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oColumn = oMatrix.Columns.Item("V_7")
            oColumn.DataBind.Bind("oMatrixDT", "DocDate")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oColumn = oMatrix.Columns.Item("V_8")
            oColumn.DataBind.Bind("oMatrixDT", "DocTotal")
            oColumn.TitleObject.Sortable = True
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

            'Try
            '    o.CollapseLevel = 2
            'Catch ex As Exception
            '    oMatrix.CollapseLevel = 2
            '    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try
            Dim oTes As SAPbobsCOM.Recordset
            oTes = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTes.DoQuery(strsql2)
            dblDepositedamount = oTes.Fields.Item(1).Value

            oTes.DoQuery(strsql4)
            dblDepositedamount = dblDepositedamount + oTes.Fields.Item(1).Value
            oTes.DoQuery(strsql1)
            dblInvoicedAmount = oTes.Fields.Item(0).Value
            oTes.DoQuery(strSql3)
            dblOwnerAmount = oTes.Fields.Item(1).Value
            Dim dblCreditNote As Double
            oTes.DoQuery(strCreditNote)
            dblCreditNote = oTes.Fields.Item(0).Value

            oApplication.Utilities.setEdittextvalue(oForm, "26", dblCreditNote)
            oApplication.Utilities.setEdittextvalue(oForm, "edDep", dblDepositedamount)
            oApplication.Utilities.setEdittextvalue(oForm, "edInv", dblInvoicedAmount)
            dblInvoicedAmount = dblInvoicedAmount - dblCreditNote
            oApplication.Utilities.setEdittextvalue(oForm, "edBal", dblInvoicedAmount - dblDepositedamount)
            'oApplication.Utilities.setEdittextvalue(oForm, "22", dblOwnerAmount)
            oForm.PaneLevel = 0
        Else
            oGrid.DataTable.ExecuteQuery(strString)
        End If

        If aChoice <> "Activity" Then
            oGrid.Columns.Item(2).TitleObject.Caption = "TransType"
            oGrid.Columns.Item(2).TitleObject.Sortable = True
            oGrid.Columns.Item(2).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(0).TitleObject.Caption = "Contract No"
            oGrid.Columns.Item(3).TitleObject.Caption = "Document Entry"
            oGrid.Columns.Item(3).TitleObject.Sortable = True
            oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(4).TitleObject.Caption = "Document Number"
            oGrid.Columns.Item(4).TitleObject.Sortable = True
            oGrid.Columns.Item(4).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(5).TitleObject.Caption = "Customer Code"
            oGrid.Columns.Item(5).TitleObject.Sortable = True
            oGrid.Columns.Item(5).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(6).TitleObject.Caption = "Customer Name"
            oGrid.Columns.Item(6).TitleObject.Sortable = True
            oGrid.Columns.Item(6).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(7).TitleObject.Caption = "Document Date"
            oGrid.Columns.Item(7).TitleObject.Sortable = True
            oGrid.Columns.Item(7).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oGrid.Columns.Item(8).TitleObject.Caption = "Document Total"
            oGrid.Columns.Item(8).TitleObject.Sortable = True
            oGrid.Columns.Item(8).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
            oEditTextColumn = oGrid.Columns.Item(3)
            oEditTextColumn.LinkedObjectType = 203
            Try
                oGrid.CollapseLevel = 2
            Catch ex As Exception
                oGrid.CollapseLevel = 2
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            oForm.Items.Item("3").Visible = True
            oForm.Items.Item("4").Visible = True
            oForm.Items.Item("5").Visible = True
            oForm.Items.Item("6").Visible = True
            oForm.Items.Item("7").Visible = True
            oForm.Items.Item("13").Visible = True
        Else
            oEditTextColumn = oGrid.Columns.Item(0)
            oGrid.Columns.Item("TransType").Visible = False
            oEditTextColumn.LinkedObjectType = 33
            oEditTextColumn = oGrid.Columns.Item("Duration in  Hours")
            oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            oForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("3").Visible = False
            oForm.Items.Item("4").Visible = False
            oForm.Items.Item("5").Visible = False
            oForm.Items.Item("6").Visible = False
            oForm.Items.Item("7").Visible = False
            oForm.Items.Item("13").Visible = False
            oForm.Title = "Related Activities"
        End If
        oGrid.AutoResizeColumns()
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
            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region
#Region "DataBind"
    Private Sub DataBind(ByVal aform As SAPbouiCOM.Form)
        oForm = aform
        oForm.Freeze(True)
        oGrid = oForm.Items.Item("1").Specific
        Dim oTest As SAPbobsCOM.Recordset
        Dim achoice, acode, strString, strSql, strfromBP, strToBP, strFromBooking, strToBooking As String
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strfromBP = oApplication.Utilities.getEdittextvalue(aform, "4")
        Dim stCondition, stCondition1 As String
        strToBP = oApplication.Utilities.getEdittextvalue(aform, "7")
        If strfromBP <> "" Then
            strSql = "X.CardCode >='" & strfromBP & "'"
            stCondition = " CardCode >='" & strfromBP & "'"
            stCondition1 = "T4.U_Z_TenCode >='" & strfromBP & "'"
        Else
            strSql = "1=1"
            stCondition = " 1=1"
            stCondition1 = " 1=1"
        End If
        If strToBP <> "" Then
            strSql = strSql & " and X.CardCode <='" & strToBP & "'"
            stCondition1 = stCondition1 & " and T4.U_Z_TenCode <='" & strToBP & "'"
            stCondition = stCondition & " and CardCode <='" & strToBP & "'"
        Else
            strSql = strSql & " and 2=2"
            stCondition = stCondition & " and 2=2"
            stCondition1 = stCondition1 & " and 2=2"
        End If

      
        'strString = "Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','DP' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where  isnull(U_Z_ContID,'')<>''"
        'strString = strString & " Union All Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate,sum(t1.LineTotal) 'Doctotal' from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry    where isnull(U_Z_ContID,'')<>'' or isnull(U_Z_ContNumber,'')<>'' group by U_Z_ContID,U_Z_ContNumber ,T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate "
        'strString = strString & " Union All  Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where isnull(U_Z_ContID,'')<>''"
        'strString = strString & " union all Select U_Z_ContID 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where (isnull(U_Z_ContID,'')<>'')"
        ''  strString = strString & " union all Select U_Z_ContID 'U_Z_ContID' , 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','DE' 'TransType',DeposId 'DocEntry',DeposNum 'DocNum',CardCode 'CardCode','" & cardName & "' 'CardName',DeposDate 'DocDate',LocTotal 'DocTotal' from ODPS  where (isnull(U_Z_ContID,'')<>''"
        ' strString = strString & " union all Select U_Z_ContID 'U_Z_ContID' , 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','JE' 'TransType',TransId 'DocEntry',BatchNum 'DocNum',ShortName 'CardCode','" & cardName & "' 'CardName',DueDate 'DocDate',(Debit-Credit) 'DocTotal' from JDT1  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
        ' strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,Transref 'ContractNumber','CFP' 'TransType',CheckKey 'DocEntry',CheckNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',PmntDate 'DocDate',CheckSum 'DocTotal' from OCHO  where ( Transref='" & strContractNumber & "')"



        ' strString = strString & " union all Select U_Z_ContID 'U_Z_ContID',U_Z_ContNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where  isnull(U_Z_ContID,'')<>''"

        'strString = "Select X.U_Z_ContID,X.TransType,X.DocEntry,X.DocNum,X.CardCode,X.CardName,X.DocDate,X.DocTotal from (" & strString & ") x  where " & strSql & "  order by X.TransType,X.U_Z_ContID,X.CardCode,X.DocDate"


        'strString = "Select U_Z_ContID, U_Z_ContNumber 'ContractNumber','DP'  'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where (isnull(U_Z_ContID,'')=" & acode & " )"
        '' strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from OINV where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
        'strString = strString & " Union All Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate,sum(t1.LineTotal) 'Doctotal' from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where  (isnull(U_Z_ContID,'')=" & acode & ") group by U_Z_ContID,U_Z_ContNumber ,T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate "
        'strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where  (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
        ''  strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
        'strString = strString & " union all Select T1. U_Z_ContID ,T1.U_Z_ContNumber 'ContractNumber','RC' 'TransType',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocDate,T1.CheckSum from ORCT T0  INNER JOIN RCT1 T1 ON T0.DocEntry = T1.DocNum where (isnull(T1.U_Z_ContID,'')=" & acode & ")"
        '' strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','DE' 'TransType',DeposId 'DocEntry',DeposNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',DeposDate 'DocDate',LocTotal 'DocTotal' from ODPS  where (U_Z_ContNumber='" & strContractNumber & "')"
        'strString = strString & " Union All SELECT T1.[U_Z_CONTID], T1.[U_Z_CONTNUMBER]'ContractNumber','DE' 'TransType', T2.[DeposNum] 'DocEntry', T2.[DeposNum] 'DocNum', T0.[CardCode],'" & cardName & "' 'CardName', T2.[DeposDate], T1.[CheckSum] 'DocTotal' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId where (isnull(T1.U_Z_ContID,'')=" & acode & ")"
        'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','JE' 'TransType',TransId 'DocEntry',BatchNum 'DocNum',ShortName 'CardCode','" & cardName & "' 'CardName',DueDate 'DocDate',(Debit-Credit) 'DocTotal' from JDT1  where (U_Z_ContNumber='" & strContractNumber & "')"
        'strString = strString & " union all Select " & acode & " 'U_Z_ContID' ,Transref 'ContractNumber','CFP' 'TransType',CheckKey 'DocEntry',CheckNum 'DocNum','" & strTenCode & "' 'CardCode','" & cardName & "' 'CardName',PmntDate 'DocDate',CheckSum 'DocTotal' from OCHO  where ( Transref='" & strContractNumber & "')"




        'strString = "Select U_Z_ContID, U_Z_ContNumber 'ContractNumber','DP'  'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>'')"
        'strString = strString & " Union All Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','IN' 'TransType',T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate,sum(t1.LineTotal) 'Doctotal' from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where  ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>'')) group by U_Z_ContID,U_Z_ContNumber ,T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate "
        'strString = strString & " union all Select U_Z_ContID,U_Z_ContNumber 'ContractNumber','CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where  ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>''))"
        'strString = strString & " union all Select U_Z_ContID 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>'')) and CheckSum<=0"
        'strString = strString & " union all Select T1. U_Z_ContID ,T1.U_Z_ContNumber 'ContractNumber','RC' 'TransType',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocDate,T1.CheckSum 'Doctotal' from ORCT T0  INNER JOIN RCT1 T1 ON T0.DocEntry = T1.DocNum where (isnull(T1.U_Z_ContID,'')<>''  and T1.CheckSum > 0)"

        'strString = strString & " Union All SELECT T1.[U_Z_CONTID], T1.[U_Z_CONTNUMBER]'ContractNumber','DE' 'TransType', T2.[DeposId] 'DocEntry', T2.[DeposNum] 'DocNum', T0.[CardCode],' ' 'CardName', T2.[DeposDate] 'DocDate', T1.[CheckSum] 'DocTotal' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId where (isnull(T1.U_Z_ContID,'')<>'')"
        ' strString = strString & " union all Select T3.DocEntry 'U_Z_CONTID'   ,U_Z_ContNumber 'ContractNumber','DE' 'TransType',DeposId 'DocEntry',DeposNum 'DocNum',T3.U_Z_TENCODE 'CardCode',T3.U_Z_TENNAME  'CardName',DeposDate 'DocDate',LocTotal 'DocTotal' from ODPS   T0 inner Join [@Z_CONTRACT] T3 on T3.U_Z_CONNO =T0.U_Z_CONTNUMBER  where ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>''))"
        'strString = strString & " union all Select U_Z_ContID 'U_Z_ContID' ,U_Z_ContNumber 'ContractNumber','JE' 'TransType',TransId 'DocEntry',BatchNum 'DocNum',ShortName 'CardCode', T5.CardName 'CardName',DueDate 'DocDate',(Debit-Credit) 'DocTotal' from JDT1 T1 inner Join OCRD T5 on T5.CardCode=T1.ShortName  where ((isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>''))"
        'strString = strString & " union all  Select T5.DocEntry  'U_Z_ContID' ,Transref 'ContractNumber','CFP' 'TransType',CheckKey 'DocEntry',CheckNum 'DocNum',T5.U_Z_TENCODE 'CardCode',T5.U_Z_TENNAME  'CardName',PmntDate 'DocDate',CheckSum 'DocTotal' from OCHO T4 inner Join [@Z_CONTRACT]  T5 on T5.U_Z_CONNO=T4.TransRef "


        strString = "Select U_Z_ContID, U_Z_CntNumber 'ContractNumber','DP'  'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ODPI where (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>'')"
        strString = strString & " Union All Select U_Z_ContID,U_Z_CntNumber 'ContractNumber','IN' 'TransType',T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate,sum(t1.LineTotal) 'Doctotal' from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where  ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>'')) group by U_Z_ContID,U_Z_CntNumber ,T0.DocEntry,T0.DocNum,CardCode,CardName,T0.DocDate "
        strString = strString & " union all Select U_Z_ContID,U_Z_CntNumber 'ContractNumber','CR' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORIN where  ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>''))"
        strString = strString & " union all Select U_Z_ContID 'U_Z_ContID' ,U_Z_CntNumber 'ContractNumber','RC' 'TransType',DocEntry,DocNum,CardCode,CardName,DocDate,Doctotal from ORCT where ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>'')) and CheckSum<=0"
        strString = strString & " union all Select T1. U_Z_ContID ,T1.U_Z_CntNumber 'ContractNumber','RC' 'TransType',T0.DocEntry,T0.DocNum,T0.CardCode,T0.CardName,T0.DocDate,T1.CheckSum 'Doctotal' from ORCT T0  INNER JOIN RCT1 T1 ON T0.DocEntry = T1.DocNum where (isnull(T1.U_Z_ContID,'')<>''  and T1.CheckSum > 0)"

        strString = strString & " Union All SELECT T1.[U_Z_CONTID], T1.[U_Z_CONTNUMBER]'ContractNumber','DE' 'TransType', T2.[DeposId] 'DocEntry', T2.[DeposNum] 'DocNum', T0.[CardCode],' ' 'CardName', T2.[DeposDate] 'DocDate', T1.[CheckSum] 'DocTotal' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId where (isnull(T1.U_Z_ContID,'')<>'')"
        ' strString = strString & " union all Select T3.DocEntry 'U_Z_CONTID'   ,U_Z_ContNumber 'ContractNumber','DE' 'TransType',DeposId 'DocEntry',DeposNum 'DocNum',T3.U_Z_TENCODE 'CardCode',T3.U_Z_TENNAME  'CardName',DeposDate 'DocDate',LocTotal 'DocTotal' from ODPS   T0 inner Join [@Z_CONTRACT] T3 on T3.U_Z_CONNO =T0.U_Z_CONTNUMBER  where ( (isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>''))"
        strString = strString & " union all Select U_Z_ContID 'U_Z_ContID' ,U_Z_CntNumber 'ContractNumber','JE' 'TransType',TransId 'DocEntry',BatchNum 'DocNum',ShortName 'CardCode', T5.CardName 'CardName',DueDate 'DocDate',(Debit-Credit) 'DocTotal' from JDT1 T1 inner Join OCRD T5 on T5.CardCode=T1.ShortName  where ((isnull(U_Z_ContID,'')<>''  or U_Z_ContNumber<>''))"
        strString = strString & " union all  Select T5.DocEntry  'U_Z_ContID' ,Transref 'ContractNumber','CFP' 'TransType',CheckKey 'DocEntry',CheckNum 'DocNum',T5.U_Z_TENCODE 'CardCode',T5.U_Z_TENNAME  'CardName',PmntDate 'DocDate',CheckSum 'DocTotal' from OCHO T4 inner Join [@Z_CONTRACT]  T5 on T5.U_Z_CONNO=T4.TransRef "


        ' strString = "Select X.CardCode,X.Cardname, X.ContractNumber,X.TransType,X.DocEntry,X.DocNum,X.DocDate,X.DocTotal from (" & strString & ") x  where " & strSql & "  order by X.TransType,X.CardCode,X.DocEntry,X.DocDate"
        strString = "Select X.U_Z_ContID,X.ContractNumber,X.TransType,X.DocEntry,X.DocNum,X.CardCode,X.CardName,X.DocDate,X.DocTotal from (" & strString & ") x  where " & strSql & "  order by X.TransType,X.CardCode,X.DocEntry,X.DocDate"

        Dim strsql1, strSQl2, strSQL3, strsql4, strCreditNote As String
        strsql1 = "Select  sum(t1.LineTotal) 'Doctotal',0 from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(U_Z_ContID,'')<>'' or U_Z_ContNumber<>'') and isnull(U_Z_InvType,'T')='T'" & " and " & stCondition
        'strSql3 = "Select  sum(t1.LineTotal) 'Doctotal',0 from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "') and isnull(U_Z_InvType,'T')='O'"
        strCreditNote = "Select  sum(t1.LineTotal) 'Doctotal',0 from ORIN T0 inner Join RIN1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(U_Z_ContID,'')<>'' or U_Z_ContNumber<>'') and isnull(U_Z_InvType,'T')='T'" & " and " & stCondition

        strSQL3 = " Select 0 'x',sum(CheckSum) 'DocTotal1' from OCHO T3 inner Join [@Z_CONTRACT]  T4 on T4.U_Z_CONNO=T3.TransRef where " & stCondition1
        strSQl2 = " Select 0 'x',sum(LocTotal) 'DocTotal1' from ODPS  T0 inner Join [@Z_CONTRACT] T4 on T4.U_Z_CONNO =T0.U_Z_CONTNUMBER  where  (isnull(U_Z_ContID,'')<>'' or U_Z_ContNumber<>'') and " & stCondition1
        strSQl2 = strSQl2 & " Union All SELECT 0, Sum(T1.[CheckSum]) 'DocTotal1' FROM OCHH T0  INNER JOIN RCT1 T1 ON T0.CheckKey = T1.CheckAbs INNER JOIN ODPS T2 ON T0.DpstAbs = T2.DeposId inner Join [@Z_CONTRACT] T4 on T4.DocEntry   =T1.U_Z_CONTID   where " & stCondition1 & " and (isnull(T1.U_Z_ContID,'')<>'')"
        strSQl2 = "Select Sum(X.x),SUm(X.DocTotal1) from (" & strSQl2 & ") X"
        strsql4 = " Select 0,sum(cashsum+creditsum+trsfrsum)  'DocTotal1' from ORCT  where checksum<=0 and (isnull(U_Z_ContID,'')<>'' or U_Z_ContNumber<>'' ) And " & stCondition

        Dim oColumn As SAPbouiCOM.Column
        oMatrix = aform.Items.Item("11").Specific
        oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery(strString)
        oColumn = oMatrix.Columns.Item("V_10")
        oColumn.DataBind.Bind("oMatrixDT", "U_Z_ContID")
        oColumn.TitleObject.Sortable = True

        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        oColumn = oMatrix.Columns.Item("V_1")
        oColumn.DataBind.Bind("oMatrixDT", "ContractNumber")
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oColumn = oMatrix.Columns.Item("V_2")
        oColumn.DataBind.Bind("oMatrixDT", "TransType")
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oColumn = oMatrix.Columns.Item("V_3")
        oColumn.DataBind.Bind("oMatrixDT", "DocEntry")
        '  oColumn.Type = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oColumn = oMatrix.Columns.Item("V_4")
        oColumn.DataBind.Bind("oMatrixDT", "DocNum")
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oColumn = oMatrix.Columns.Item("V_5")
        oColumn.DataBind.Bind("oMatrixDT", "CardCode")
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        oColumn = oMatrix.Columns.Item("V_6")
        oColumn.DataBind.Bind("oMatrixDT", "CardName")
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oColumn = oMatrix.Columns.Item("V_7")
        oColumn.DataBind.Bind("oMatrixDT", "DocDate")
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oColumn = oMatrix.Columns.Item("V_8")
        oColumn.DataBind.Bind("oMatrixDT", "DocTotal")
        oColumn.TitleObject.Sortable = True
        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oMatrix.LoadFromDataSource()


        oGrid.DataTable.ExecuteQuery(strString)
        oGrid.Columns.Item(3).TitleObject.Caption = "TransType"
        oGrid.Columns.Item(3).TitleObject.Sortable = True
        oGrid.Columns.Item(3).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oGrid.Columns.Item(2).TitleObject.Caption = "Contract No"
        oGrid.Columns.Item(4).TitleObject.Caption = "Document Entry"
        oGrid.Columns.Item(4).TitleObject.Sortable = True
        oGrid.Columns.Item(4).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oGrid.Columns.Item(5).TitleObject.Caption = "Document Number"
        oGrid.Columns.Item(5).TitleObject.Sortable = True
        oGrid.Columns.Item(5).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oGrid.Columns.Item(0).TitleObject.Caption = "Customer Code"
        oGrid.Columns.Item(1).TitleObject.Caption = "Customer Name"
        oGrid.Columns.Item(6).TitleObject.Caption = "Document Date"
        oGrid.Columns.Item(6).TitleObject.Sortable = True
        oGrid.Columns.Item(6).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oGrid.Columns.Item(7).TitleObject.Caption = "Document Total"
        oGrid.Columns.Item(7).TitleObject.Sortable = True
        oGrid.Columns.Item(7).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
        oEditTextColumn = oGrid.Columns.Item("DocEntry")
        oEditTextColumn.LinkedObjectType = 203
        'Try
        '    oGrid.CollapseLevel = 3
        'Catch ex As Exception
        '    oGrid.CollapseLevel = 1
        '    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'End Try
        oGrid.AutoResizeColumns()
        Dim dblDepositedamount, dblInvoicedAmount, dblOwnerAmount As Double
        ''Dim strsql1, strsql2, strSql3 As String

        ''strsql1 = "Select  sum(t1.LineTotal) 'Doctotal',0 from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(Card,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "') and isnull(U_Z_InvType,'T')='T'"
        ' ''strSql3 = "Select  sum(t1.LineTotal) 'Doctotal',0 from OINV T0 inner Join INV1 T1 on T1.DocEntry=T0.DocEntry  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "') and isnull(U_Z_InvType,'T')='O'"
        ' ''strSql3 = " Select 0,sum(CheckSum) 'DocTotal1' from OCHO  where (Transref='" & strContractNumber & "')"
        ' '' strsql2 = " Select 0,sum(LocTotal) 'DocTotal1' from ODPS  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"
        ''strsql2 = " Select 0,sum(DocTotal) 'DocTotal1' from ORCT  where (isnull(U_Z_ContID,'')=" & acode & " or U_Z_ContNumber='" & strContractNumber & "')"

        'oApplication.Utilities.setEdittextvalue(aform, "edInv", "0")
        'oApplication.Utilities.setEdittextvalue(aform, "edDep", "0")
        'oApplication.Utilities.setEdittextvalue(aform, "edBal", "0")

        Dim oTes As SAPbobsCOM.Recordset
        oTes = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTes.DoQuery(strSQl2)
        dblDepositedamount = oTes.Fields.Item(1).Value

        oTes.DoQuery(strsql4)
        dblDepositedamount = dblDepositedamount + oTes.Fields.Item(1).Value
        oTes.DoQuery(strsql1)
        dblInvoicedAmount = oTes.Fields.Item(0).Value
        oTes.DoQuery(strSQL3)
        dblOwnerAmount = oTes.Fields.Item(1).Value
        Dim dblcreditnote As Double
        oTes.DoQuery(strCreditNote)
        dblcreditnote = oTes.Fields.Item(0).Value

        oApplication.Utilities.setEdittextvalue(oForm, "edDep", dblDepositedamount)
        oApplication.Utilities.setEdittextvalue(oForm, "edInv", dblInvoicedAmount)
        oApplication.Utilities.setEdittextvalue(oForm, "26", dblcreditnote)
        oApplication.Utilities.setEdittextvalue(oForm, "edBal", dblInvoicedAmount - dblcreditnote - dblDepositedamount)
        'oApplication.Utilities.setEdittextvalue(oForm, "22", dblOwnerAmount)
        oForm.Freeze(False)
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_DocumentView Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then


                                    oGrid = oForm.Items.Item("1").Specific
                                    Dim achoice As String
                                    ' MsgBox(oGrid.GetDataTableRowIndex(pVal.Row))
                                    achoice = oGrid.DataTable.GetValue("TransType", oGrid.GetDataTableRowIndex(pVal.Row))

                                    oEditTextColumn = oGrid.Columns.Item("DocEntry")
                                    oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_None
                                    Select Case achoice
                                        Case "DP"
                                            oEditTextColumn.LinkedObjectType = 203
                                        Case "IN"
                                            oEditTextColumn.LinkedObjectType = 13
                                        Case "RC"
                                            oEditTextColumn.LinkedObjectType = 24
                                        Case "DE"
                                            oEditTextColumn.LinkedObjectType = 25
                                        Case "DL"
                                            oEditTextColumn.LinkedObjectType = 15
                                        Case "CFP"
                                            oEditTextColumn.LinkedObjectType = 57
                                        Case "ACT"
                                            oEditTextColumn.LinkedObjectType = 33
                                        Case "JE"
                                            oEditTextColumn.LinkedObjectType = 30
                                        Case "CR"
                                            oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                                    End Select
                                End If

                                If pVal.ItemUID = "11" Then


                                    oMatrix = oForm.Items.Item("11").Specific
                                    Dim achoice As String
                                    ' MsgBox(oGrid.GetDataTableRowIndex(pVal.Row))
                                    achoice = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", pVal.Row) ' oGrid.DataTable.GetValue("TransType", oGrid.GetDataTableRowIndex(pVal.Row))
                                    Dim oColumn As SAPbouiCOM.Column
                                    Dim oLinkedObject As SAPbouiCOM.LinkedButton
                                    oColumn = oMatrix.Columns.Item("V_3")
                                    oLinkedObject = oColumn.ExtendedObject
                                    oLinkedObject.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_None

                                    '  oEditTextColumn = oGrid.Columns.Item("DocEntry")
                                    ' oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_None
                                    Select Case achoice
                                        Case "DP"
                                            oLinkedObject.LinkedObjectType = 203
                                        Case "IN"
                                            oLinkedObject.LinkedObjectType = 13
                                        Case "RC"
                                            oLinkedObject.LinkedObjectType = 24
                                        Case "DE"
                                            oLinkedObject.LinkedObjectType = 25
                                        Case "DL"
                                            oLinkedObject.LinkedObjectType = 15
                                        Case "CFP"
                                            oLinkedObject.LinkedObjectType = 57
                                        Case "ACT"
                                            oLinkedObject.LinkedObjectType = 33
                                        Case "JE"
                                            oLinkedObject.LinkedObjectType = 30
                                        Case "CR"
                                            oLinkedObject.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                                    End Select
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" Then
                                    DataBind(oForm)
                                End If
                                If pVal.ItemUID = "24" Then
                                    Dim obj As New clsPrint
                                    obj.PrintDocuments(oForm)

                                End If
                              

                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oGrid = oForm.Items.Item("1").Specific
                                Dim achoice As String
                                'achoice = oGrid.DataTable.GetValue("TransType", oGrid.GetDataTableRowIndex(pVal.Row))
                                'oEditTextColumn = oGrid.Columns.Item("DocEntry")
                                'Select Case achoice
                                '    Case "DP"
                                '        oEditTextColumn.LinkedObjectType = 203
                                '    Case "IN"
                                '        oEditTextColumn.LinkedObjectType = 13
                                '    Case "RC"
                                '        oEditTextColumn.LinkedObjectType = 24
                                '    Case "DE"
                                '        oEditTextColumn.LinkedObjectType = 25
                                '    Case "DL"
                                '        oEditTextColumn.LinkedObjectType = 15
                                '    Case "CFP"
                                '        oEditTextColumn.LinkedObjectType = 57
                                '    Case "ACT"
                                '        oEditTextColumn.LinkedObjectType = 33
                                '    Case "JE"
                                '        oEditTextColumn.LinkedObjectType = 30

                                '    Case "CR"
                                '        oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                                'End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim sCHFL_ID As String

                                Dim intChoice As Integer
                                Dim codebar, val2 As String
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
                                        If pVal.ItemUID = "18" Then
                                        End If
                                        If pVal.ItemUID = "4" Or pVal.ItemUID = "7" Then
                                            val2 = oDataTable.GetValue("CardCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val2)
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
