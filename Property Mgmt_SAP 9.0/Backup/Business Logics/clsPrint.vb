Imports System
Imports System.Collections
Imports System.ComponentModel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports System.Collections.Generic

Public Class clsPrint
    'Private rptaccountreport As New AcctStatement
    Dim cryRpt As New ReportDocument
    Private ds As New dtProperty      '(dataset)
    Private oDRow As DataRow
#Region "Add Crystal Report"

    Private Sub addCrystal(ByVal ds1 As DataSet, ByVal aChoice As String, ByVal aCode As String)
        Dim strFilename, strCompanyName, stfilepath As String
        Dim blnCrystal As Boolean = False
        Dim strReportFileName As String
        If aChoice = "Contract" Then
            strReportFileName = "Contract.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Contract_Agreement"
        ElseIf aChoice = "Document" Then
            strReportFileName = "cryDocuments.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Financial_Transactions"
        ElseIf aChoice = "Contract-Default" Then
            strReportFileName = "ContractTen.rpt"
            ' strReportFileName = "Awqaf Contract.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Contract_Ten_Agreement"
            blnCrystal = False
        ElseIf aChoice = "ContractTen" Then
            strReportFileName = "ContractTen.rpt"
            ' strReportFileName = "Awqaf Contract.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Contract_Ten_Agreement"
            blnCrystal = True

        ElseIf aChoice = "Private Eng office" Then
            strReportFileName = "Private Eng office.rpt"
            ' strReportFileName = "Awqaf Contract.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Private Eng office"
            blnCrystal = True
        ElseIf aChoice = "Awqaf" Then
            ' strReportFileName = "ContractTen.rpt"
            strReportFileName = "Awqaf Contract.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Contract_Ten_Agreement"
            blnCrystal = True
        ElseIf aChoice = "Rent-General" Then
            ' strReportFileName = "ContractTen.rpt"
            strReportFileName = "rent-general format.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Contract_Ten_Agreement"
            blnCrystal = True
        ElseIf aChoice = "Al Bayan" Then
            strReportFileName = "ContractTen.rpt"
            strReportFileName = "al bayan.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Contract_Ten_Agreement"
            blnCrystal = True
        ElseIf aChoice = "Evaluation" Then
            strReportFileName = "Evaluation.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Evaluation"
            blnCrystal = False
        ElseIf aChoice = "Evaluation_Arabic" Then
            strReportFileName = "Evaluation_ar.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Evaluation"
            blnCrystal = True
        ElseIf aChoice = "Agreement" Then
            strReportFileName = "Agreement.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Rental_Agreement"
        ElseIf aChoice = "Receiable" Then
            strReportFileName = "Receivable.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Receivable"
        ElseIf aChoice = "Production" Then
            strReportFileName = "Production.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Production"
        ElseIf aChoice = "SalesAgent" Then
            strReportFileName = "SalesAgent.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\SalesAgent"
        Else
            strReportFileName = "AcctStatement.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\AccountStatement"
        End If
        strReportFileName = strReportFileName
        strFilename = strFilename & ".pdf"
        '  strFilename = strFilename & ".doc"
        stfilepath = System.Windows.Forms.Application.StartupPath & "\CrystalReports\" & strReportFileName
        If File.Exists(stfilepath) = False Then
            oApplication.Utilities.Message("Report does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        If File.Exists(strFilename) Then
            File.Delete(strFilename)
        End If
        ' If ds1.Tables.Item("AccountBalance").Rows.Count > 0 Then
        If 1 = 1 Then
            cryRpt.Load(System.Windows.Forms.Application.StartupPath & "\CrystalReports\" & strReportFileName)
            Try
                cryRpt.SetDataSource(ds1)
            Catch ex As Exception
            End Try




            If blnCrystal = True Then
                Dim mythread As New System.Threading.Thread(AddressOf openFileDialog)
                mythread.SetApartmentState(ApartmentState.STA)
                mythread.Start()
                mythread.Join()
                ds1.Clear()
            Else

                'Dim reportWord As New CrystalReport1() ' Report Name         
                'Dim strExportFile As String = "d:\TestWord.doc"
                'reportWord.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                'reportWord.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows
                'Dim objOptions As DiskFileDestinationOptions = New DiskFileDestinationOptions()
                'objOptions.DiskFileName = strExportFile
                'reportWord.ExportOptions.DestinationOptions = objOptions
                'reportWord.SetDataSource(myDS)
                'reportWord.Export()


                'Dim CrExportOptions As ExportOptions
                'Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                ''Dim CrFormatTypeOptions As New DiskFileDestinationOptions
                'CrDiskFileDestinationOptions.DiskFileName = strFilename
                'CrExportOptions = cryRpt.ExportOptions
                'With CrExportOptions
                '    .ExportDestinationType = ExportDestinationType.DiskFile
                '    .ExportFormatType = ExportFormatType.WordForWindows
                '    .DestinationOptions = CrDiskFileDestinationOptions
                '    '  .FormatOptions = CrFormatTypeOptions
                'End With
                'cryRpt.Export()
                'cryRpt.Close()

                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New _
                DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                CrDiskFileDestinationOptions.DiskFileName = strFilename
                CrExportOptions = cryRpt.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                cryRpt.Export()
                cryRpt.Close()
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                ' objUtility.ShowSuccessMessage("Report exported into PDF File")
            End If

        Else
            ' objUtility.ShowWarningMessage("No data found")
        End If

    End Sub

    Private Sub Print_Evalution(ByVal aChoice As String, ByVal aCode As String)
        Dim strFilename, strCompanyName, stfilepath As String
        Dim strReportFileName As String
        If aChoice = "Contract" Then
            strReportFileName = "Contract.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Contract_Agreement"
        ElseIf aChoice = "Evaluation" Then
            strReportFileName = "Evaluation.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Evaluation"
        ElseIf aChoice = "Agreement" Then
            strReportFileName = "Agreement.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Rental_Agreement"
        ElseIf aChoice = "Receiable" Then
            strReportFileName = "Receivable.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Receivable"
        ElseIf aChoice = "Production" Then
            strReportFileName = "Production.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\Production"
        ElseIf aChoice = "SalesAgent" Then
            strReportFileName = "SalesAgent.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\SalesAgent"
        Else
            strReportFileName = "AcctStatement.rpt"
            strFilename = System.Windows.Forms.Application.StartupPath & "\AccountStatement"
        End If
        strReportFileName = strReportFileName
        strFilename = strFilename & ".pdf"
        stfilepath = System.Windows.Forms.Application.StartupPath & "\CrystalReports\" & strReportFileName
        If File.Exists(stfilepath) = False Then
            oApplication.Utilities.Message("Report does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        If File.Exists(strFilename) Then
            File.Delete(strFilename)
        End If
        ' If ds1.Tables.Item("AccountBalance").Rows.Count > 0 Then
        If 1 = 1 Then
            cryRpt.Load(System.Windows.Forms.Application.StartupPath & "\CrystalReports\" & strReportFileName)
            ' cryRpt.SetDataSource(ds1)
            If "T" = "W" Then
                Dim mythread As New System.Threading.Thread(AddressOf openFileDialog)
                mythread.SetApartmentState(ApartmentState.STA)
                mythread.Start()
                mythread.Join()
                '  ds1.Clear()
            Else
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New _
                DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
                CrDiskFileDestinationOptions.DiskFileName = strFilename
                CrExportOptions = cryRpt.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                cryRpt.Export()
                cryRpt.Close()
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                ' objUtility.ShowSuccessMessage("Report exported into PDF File")
            End If

        Else
            ' objUtility.ShowWarningMessage("No data found")
        End If

    End Sub

    Private Sub openFileDialog()
        Dim objPL As New frmReportViewer
        objPL.iniViewer = AddressOf objPL.GenerateReport
        objPL.rptViewer.ReportSource = cryRpt
        objPL.rptViewer.Refresh()
        objPL.WindowState = FormWindowState.Maximized
        objPL.ShowDialog()
        System.Threading.Thread.CurrentThread.Abort()
    End Sub

    Public Sub PrintContract(ByVal aOrderNo As Integer, ByVal aChoice As String)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select isnull(U_Z_Type,'O')'Type',* from [@Z_CONTRACT] where DocEntry=" & aOrderNo)
        ds.Clear()
        If oRec.RecordCount <= 0 Then
            oApplication.Utilities.Message("Contract  details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            oDRow = ds.Tables("Contract").NewRow()
            oDRow.Item("ID") = oRec.Fields.Item("DocEntry").Value
            oDRow.Item("ContNumber") = oRec.Fields.Item("U_Z_ConNo").Value ' oRec.Fields.Item("DocEntry").Value
            oDRow.Item("UnitCode") = oRec.Fields.Item("U_Z_UnitCode").Value
            oDRow.Item("StartDate") = oRec.Fields.Item("U_Z_StartDate").Value
            oDRow.Item("EndDate") = oRec.Fields.Item("U_Z_EndDate").Value
            oDRow.Item("Status") = oRec.Fields.Item("U_Z_Status").Value
            oDRow.Item("TenCode") = oRec.Fields.Item("U_Z_TenCode").Value
            '  oDRow.Item("PickupLocation") = oRec.Fields.Item("U_Z_ToLOC").Value
            oDRow.Item("TenName") = oRec.Fields.Item("U_Z_TenName").Value
            oDRow.Item("OffAddress") = oRec.Fields.Item("U_Z_OffAddress").Value
            ' oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_ToLOC").Value)
            oDRow.Item("AIJAddress") = oRec.Fields.Item("U_Z_AlJAddress").Value
            oDRow.Item("AnnualRent") = oRec.Fields.Item("U_Z_AnnualRent").Value
            oDRow.Item("AcctCode") = oRec.Fields.Item("U_Z_AcctCode").Value
            oDRow.Item("Deposit") = oRec.Fields.Item("U_Z_Deposit").Value
            ' oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_FromLOC").Value)
            oTemp.DoQuery("SELECT T0.[GroupNum], T0.[PymntGroup] FROM OCTG T0 where GroupNum=" & oRec.Fields.Item("U_Z_PayTrms").Value)
            oDRow.Item("PayTrms") = oTemp.Fields.Item(1).Value
            If oRec.Fields.Item("U_Z_Insurance").Value = "Y" Then
                oDRow.Item("Insurance") = "Yes"
            Else
                oDRow.Item("Insurance") = "No"
            End If
            'oDRow.Item("Insurance") = oRec.Fields.Item("U_Z_Insurance").Value
            oDRow.Item("PolicyNumber") = oRec.Fields.Item("U_Z_PolicyNumber").Value
            oDRow.Item("ChgMonth") = oRec.Fields.Item("U_Z_ChgMonth").Value
            oDRow.Item("ChgAmt") = oRec.Fields.Item("U_Z_ChgAmt").Value
            oDRow.Item("Period") = oRec.Fields.Item("U_Z_Period").Value
            oDRow.Item("Rules") = oRec.Fields.Item("U_Z_Rules").Value
            oDRow.Item("Type") = oRec.Fields.Item("U_Z_Type").Value
            oDRow.Item("OwnerCode") = oRec.Fields.Item("U_Z_OwnerCode").Value
            oTemp.DoQuery("select isnull(CardName,'') 'CardName' from OCRD where cardCode='" & oRec.Fields.Item("U_Z_OwnerCode").Value & "'")
            oDRow.Item("OwnerName") = oTemp.Fields.Item("CardName").Value
            ds.Tables("Contract").Rows.Add(oDRow)
            addCrystal(ds, aChoice, aOrderNo)
        End If
        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)

    End Sub

    Public Sub PrintContract_Tenant(ByVal aOrderNo As Integer, ByVal aChoice As String)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select isnull(U_Z_Type,'O')'Type',* from [@Z_CONTRACT] where DocEntry=" & aOrderNo)
        ds.Clear()
        If oRec.RecordCount <= 0 Then
            oApplication.Utilities.Message("Contract  details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            oDRow = ds.Tables("Contract").NewRow()
            oDRow.Item("ID") = oRec.Fields.Item("DocEntry").Value
            oDRow.Item("ContNumber") = oRec.Fields.Item("U_Z_ConNo").Value ' oRec.Fields.Item("DocEntry").Value
            oDRow.Item("UnitCode") = oRec.Fields.Item("U_Z_UnitCode").Value
            oDRow.Item("StartDate") = oRec.Fields.Item("U_Z_StartDate").Value
            oDRow.Item("EndDate") = oRec.Fields.Item("U_Z_EndDate").Value
            oDRow.Item("Status") = oRec.Fields.Item("U_Z_Status").Value
            oDRow.Item("TenCode") = oRec.Fields.Item("U_Z_TenCode").Value
            '  oDRow.Item("PickupLocation") = oRec.Fields.Item("U_Z_ToLOC").Value
            oDRow.Item("TenName") = oRec.Fields.Item("U_Z_TenName").Value
            oDRow.Item("OffAddress") = oRec.Fields.Item("U_Z_OffAddress").Value
            ' oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_ToLOC").Value)
            oDRow.Item("AIJAddress") = oRec.Fields.Item("U_Z_AlJAddress").Value
            oDRow.Item("AnnualRent") = oRec.Fields.Item("U_Z_AnnualRent").Value
            Dim dblAnnualrent As Double
            dblAnnualrent = oRec.Fields.Item("U_Z_AnnualRent").Value
            oDRow.Item("AnnualRentWord") = oApplication.Utilities.SFormatNumber(dblAnnualrent)
            If dblAnnualrent > 0 Then
                dblAnnualrent = dblAnnualrent / 12
            Else
                dblAnnualrent = 0
            End If


            oDRow.Item("MonthlyRentWord") = oApplication.Utilities.SFormatNumber(dblAnnualrent)
            oDRow.Item("AcctCode") = oRec.Fields.Item("U_Z_AcctCode").Value
            oDRow.Item("Deposit") = oRec.Fields.Item("U_Z_Deposit").Value
            ' oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_FromLOC").Value)
            Try
                oTemp.DoQuery("SELECT T0.[GroupNum], T0.[PymntGroup] FROM OCTG T0 where GroupNum=" & oRec.Fields.Item("U_Z_PayTrms").Value)
                oDRow.Item("PayTrms") = oTemp.Fields.Item(1).Value
            Catch ex As Exception
                oDRow.Item("PayTrms") = "-1"
            End Try
            
            If oRec.Fields.Item("U_Z_Insurance").Value = "Y" Then
                oDRow.Item("Insurance") = "Yes"
            Else
                oDRow.Item("Insurance") = "No"
            End If
            'oDRow.Item("Insurance") = oRec.Fields.Item("U_Z_Insurance").Value
            oDRow.Item("PolicyNumber") = oRec.Fields.Item("U_Z_PolicyNumber").Value
            oDRow.Item("ChgMonth") = oRec.Fields.Item("U_Z_ChgMonth").Value
            oDRow.Item("ChgAmt") = oRec.Fields.Item("U_Z_ChgAmt").Value
            oDRow.Item("Period") = oRec.Fields.Item("U_Z_Period").Value
            oDRow.Item("Rules") = oRec.Fields.Item("U_Z_Rules").Value
            oDRow.Item("Type") = oRec.Fields.Item("U_Z_Type").Value
            oDRow.Item("OwnerCode") = oRec.Fields.Item("U_Z_OwnerCode").Value
            oDRow.Item("ContDate") = oRec.Fields.Item("U_Z_ContDate").Value
            oTemp.DoQuery("select isnull(CardName,'') 'CardName' from OCRD where cardCode='" & oRec.Fields.Item("U_Z_OwnerCode").Value & "'")
            oDRow.Item("OwnerName") = oTemp.Fields.Item("CardName").Value

            Dim st As String
            Dim st1 As String
            '"W,I,C,R", "Windows,Inside Office,Corridor,Receiption Area"
            st1 = "Case T0.U_Z_FSTYPE when 'W' then 'Windows' when 'I' then 'Inside Office' when 'C' then 'Corridor' else 'Receiption Area' end 'U_Z_FSTYPE'"
            st = "SELECT T2.U_Z_ConNo,T1.[U_Z_UNITTYPE] 'UnitType' ,T0.U_Z_UnitNo 'UnitNo',T0.[U_Z_ELECTRICITY] 'Electricity'," & st1 & ",T0.U_Z_OWNERCODE,T0.U_Z_OWNERNAME ,T0.U_Z_Water,T0.U_Z_UnitArea,T0.U_Z_UnitStreet,"
            st = st & " T3.PHONE1, T3.PHONE2, T3.ADDID, T3.Picture, T3.Fax, T3.CNTCtPRSN, T3.CELLULAR, T3.U_Nationality, "
            st = st & "isnull(T4.FirstName,'')'FirstName',isnull(T4.MiddleName ,'') 'MiddleName',isnull(T4.LastName,'')'LastName'  , T0.[U_Z_CODE],  T1.[U_Z_STATUS], T0.[U_Z_CODE], T0.[U_Z_DESC], T0.[U_Z_PROPCODE], T2.U_Z_UnitCode,T0.[U_Z_PROITEMCODE],T0.[U_Z_PROPDESC],T0.[U_Z_OWNERFORNAME] 'OwnerForName',T0.[U_Z_OWNERREFNO] 'RefNo' FROM [dbo].[@Z_PROPUNIT]   T0  inner Join  [dbo].[@Z_UNITTYPE]  T1 on T1.DocEntry=T0.U_Z_UnitType inner Join [dbo].[@Z_CONTRACT]  T2 on T2.U_Z_UnitCode=T0.[U_Z_PROITEMCODE] left outer Join OCRD T3 on T3.CardCode=T2.U_Z_TENCODE "
            st = st & " left outer join OCPR T4 on T4.Name =T3.CntctPrsn  where T2.U_Z_ConNo='" & oRec.Fields.Item("U_Z_Conno").Value & "'"
            oTemp.DoQuery(st)
            If oTemp.RecordCount > 0 Then
                oDRow.Item("OwnerForName") = oTemp.Fields.Item("OwnerForName").Value
                oDRow.Item("OwnerRefNo") = oTemp.Fields.Item("RefNo").Value
                oDRow.Item("UnitType") = oTemp.Fields.Item("UnitType").Value
                oDRow.Item("UnitNo") = oTemp.Fields.Item("UnitNo").Value
                oDRow.Item("Electricity") = oTemp.Fields.Item("Electricity").Value
                oDRow.Item("FSTYPE") = oTemp.Fields.Item("U_Z_FSTYPE").Value
                oDRow.Item("Water") = oTemp.Fields.Item("U_Z_Water").Value
                oDRow.Item("UnitArea") = oTemp.Fields.Item("U_Z_UnitArea").Value
                oDRow.Item("UnitStreet") = oTemp.Fields.Item("U_Z_UnitStreet").Value

                oDRow.Item("Phone1") = oTemp.Fields.Item("Phone1").Value
                oDRow.Item("Phone2") = oTemp.Fields.Item("Phone2").Value
                oDRow.Item("AddID") = oTemp.Fields.Item("AddID").Value
                oDRow.Item("Picture") = oTemp.Fields.Item("Picture").Value
                oDRow.Item("Fax") = oTemp.Fields.Item("Fax").Value
                oDRow.Item("CntctPrsn") = oTemp.Fields.Item("CntctPrsn").Value
                oDRow.Item("CELLULAR") = oTemp.Fields.Item("CELLULAR").Value
                oDRow.Item("Nationality") = oTemp.Fields.Item("U_Nationality").Value
                oDRow.Item("FirstName") = oTemp.Fields.Item("FirstName").Value
                oDRow.Item("MiddleName") = oTemp.Fields.Item("MiddleName").Value
                oDRow.Item("LastName") = oTemp.Fields.Item("LastName").Value
                ' oDRow.Item("ZipCode") = oTemp.Fields.Item("CELLULAR").Value
            End If
            ds.Tables("Contract").Rows.Add(oDRow)
            addCrystal(ds, aChoice, aOrderNo)
        End If
        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)

    End Sub

    Public Sub PrintEvalution(ByVal aOrderNo As String)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aOrderNo = "" Then
            oApplication.Utilities.Message("Evaluation details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oRec.DoQuery("Select * from [@Z_AQEVAL] where DocEntry=" & CInt(aOrderNo))
        ds.Clear()
        If oRec.RecordCount <= 0 Then
            oApplication.Utilities.Message("Evaluation details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            oDRow = ds.Tables("Evaluation").NewRow()
            oDRow.Item("ID") = aOrderNo 'oRec.Fields.Item("DocEntry").Value
            oDRow.Item("Z_ProCode") = oRec.Fields.Item("U_Z_ProCode").Value
            oDRow.Item("Z_ProName") = oRec.Fields.Item("U_Z_ProName").Value
            oDRow.Item("Z_EVL_ID") = oRec.Fields.Item("U_Z_EVL_ID").Value
            oDRow.Item("Z_EV_DATE") = oRec.Fields.Item("U_Z_EV_DATE").Value
            oDRow.Item("Z_OWNERCODE") = oRec.Fields.Item("U_Z_OWNERCODE").Value
            oDRow.Item("Z_OWNER") = oRec.Fields.Item("U_Z_OWNER").Value
            oDRow.Item("Z_OWNRTEL") = oRec.Fields.Item("U_Z_OWNRTEL").Value
            oDRow.Item("Z_OWNRGSM") = oRec.Fields.Item("U_Z_OWNRGSM").Value
            '  oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_ToLOC").Value)
            oDRow.Item("Z_OWNRPOBOX") = oRec.Fields.Item("U_Z_OWNRPOBOX").Value
            oDRow.Item("Z_OWNRADD") = oRec.Fields.Item("U_Z_OWNRADD").Value
            oDRow.Item("Z_AQTYP") = oRec.Fields.Item("U_Z_AQTYP").Value
            oDRow.Item("Z_AQBOND") = oRec.Fields.Item("U_Z_AQBOND").Value

            oDRow.Item("Z_AQCITY") = oRec.Fields.Item("U_Z_AQCITY").Value
            oDRow.Item("Z_AQZONE") = oRec.Fields.Item("U_Z_AQZONE").Value
            oDRow.Item("Z_AQSTREET") = oRec.Fields.Item("U_Z_AQSTREET").Value
            oDRow.Item("Z_AQLSNC") = oRec.Fields.Item("U_Z_AQLSNC").Value
            oDRow.Item("Z_AQLSNC_DATE") = oRec.Fields.Item("U_Z_AQLSNC_DATE").Value
            oDRow.Item("Z_AQNEARST") = oRec.Fields.Item("U_Z_AQNEARST").Value
            oDRow.Item("Z_AQAGE") = oRec.Fields.Item("U_Z_AQAGE").Value
            oDRow.Item("Z_AQBUILD") = oRec.Fields.Item("U_Z_AQBUILD").Value
            oDRow.Item("Z_AQQUALITY") = oRec.Fields.Item("U_Z_AQQUALITY").Value




            oDRow.Item("Z_USRENTNAME") = oRec.Fields.Item("U_Z_USRENTNAME").Value
            oDRow.Item("Z_USRMNTHRNT") = oRec.Fields.Item("U_Z_USRMNTHRNT").Value
            oDRow.Item("Z_RENTTYP") = oRec.Fields.Item("U_Z_RENTTYP").Value
            oDRow.Item("Z_MRKTLAND") = oRec.Fields.Item("U_Z_MRKTLAND").Value
            oDRow.Item("Z_MRKTZFOOT") = oRec.Fields.Item("U_Z_MRKTZFOOT").Value
            oDRow.Item("Z_BUILTCOST") = oRec.Fields.Item("U_Z_BUILTCOST").Value
            oDRow.Item("Z_MOREVALS") = oRec.Fields.Item("U_Z_MOREVALS").Value
            oDRow.Item("Z_AQAREA") = oRec.Fields.Item("U_Z_AQAREA").Value
            oDRow.Item("Z_AQAREAMT") = oRec.Fields.Item("U_Z_AQAREAMT").Value



            oDRow.Item("Z_BLDAREAFT") = oRec.Fields.Item("U_Z_BLDAREAFT").Value
            oDRow.Item("Z_FORCPER") = oRec.Fields.Item("U_Z_FORCPER").Value
            oDRow.Item("Z_MRKTBLDMT") = oRec.Fields.Item("U_Z_MRKTBLDMT").Value
            oDRow.Item("Z_BUILDAREA") = oRec.Fields.Item("U_Z_BUILDAREA").Value
            oDRow.Item("Z_REQ") = oRec.Fields.Item("U_Z_REQ").Value
            oDRow.Item("Z_RQTEL") = oRec.Fields.Item("U_Z_RQTEL").Value
            oDRow.Item("Z_RQGSM") = oRec.Fields.Item("U_Z_RQGSM").Value
            oDRow.Item("Z_RQADD") = oRec.Fields.Item("U_Z_RQADD").Value
            oDRow.Item("Z_RQPOBOX") = oRec.Fields.Item("U_Z_RQPOBOX").Value



            oDRow.Item("Z_REQEMAIL") = oRec.Fields.Item("U_Z_REQEMAIL").Value
            oDRow.Item("Z_REQFAX") = oRec.Fields.Item("U_Z_REQFAX").Value
            oDRow.Item("Z_AQDESC") = oRec.Fields.Item("U_Z_AQDESC").Value
            oDRow.Item("Z_AQNOTES") = oRec.Fields.Item("U_Z_AQNOTES").Value
            oDRow.Item("Z_AQFACILITY") = oRec.Fields.Item("U_Z_AQFACILITY").Value
            oDRow.Item("Z_MOREFAC") = oRec.Fields.Item("U_Z_MOREFAC").Value
            oDRow.Item("Z_USEINFO") = oRec.Fields.Item("U_Z_USEINFO").Value
            oDRow.Item("Z_RENTERNAME") = oRec.Fields.Item("U_Z_RENTERNAME").Value
            oDRow.Item("Z_CHKRNAME") = oRec.Fields.Item("U_Z_CHKRNAME").Value



            oDRow.Item("Z_USERNAME") = oRec.Fields.Item("U_Z_USERNAME").Value
            oDRow.Item("Z_RCHKR") = oRec.Fields.Item("U_Z_RCHKR").Value
            oDRow.Item("Z_EVCHRG") = oRec.Fields.Item("U_Z_EVCHRG").Value
            oDRow.Item("Z_INVNO") = oRec.Fields.Item("U_Z_INVNO").Value
            oDRow.Item("Z_DBRCPER") = oRec.Fields.Item("U_Z_DBRCPER").Value
            oDRow.Item("Z_AC_ID") = oRec.Fields.Item("U_Z_AC_ID").Value
            oDRow.Item("Z_NOTES") = oRec.Fields.Item("U_Z_NOTES").Value
            oDRow.Item("Z_AQZNOTES") = oRec.Fields.Item("U_Z_AQZNOTES").Value
            ' oDRow.Item("Z_CHKRNAME") = oTemp.Fields.Item("CardName").Value

            ds.Tables("Evaluation").Rows.Add(oDRow)
            addCrystal(ds, "Evaluation", aOrderNo)
            '  Print_Evalution("Evaluation", aOrderNo)
        End If
        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
    End Sub


    Public Sub PrintDocuments(ByVal aForm As SAPbouiCOM.Form)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, aOrderNo, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Dim omatrix As SAPbouiCOM.Matrix
        omatrix = aForm.Items.Item("11").Specific
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If omatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("No Records found", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        ds.Clear()
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 1 To omatrix.RowCount
            oDRow = ds.Tables("Document").NewRow()
            oDRow.Item("CardCode") = oApplication.Utilities.getMatrixValues(omatrix, "V_5", intRow)
            oTest.DoQuery("Select * from OCRD where CardCode='" & oApplication.Utilities.getMatrixValues(omatrix, "V_5", intRow) & "'")

            oDRow.Item("CardName") = oTest.Fields.Item("CardName").Value ' oApplication.Utilities.getMatrixValues(omatrix, "V_6", intRow)
            oDRow.Item("ContractNumber") = oApplication.Utilities.getMatrixValues(omatrix, "V_1", intRow)
            oDRow.Item("ContractEntry") = oApplication.Utilities.getMatrixValues(omatrix, "V_10", intRow)
            oDRow.Item("TransType") = oApplication.Utilities.getMatrixValues(omatrix, "V_2", intRow)
            oDRow.Item("TransID") = oApplication.Utilities.getMatrixValues(omatrix, "V_3", intRow)
            oDRow.Item("TransNum") = oApplication.Utilities.getMatrixValues(omatrix, "V_4", intRow)
            oDRow.Item("DocDate") = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getMatrixValues(omatrix, "V_7", intRow))
            oDRow.Item("DocTotal") = oApplication.Utilities.getMatrixValues(omatrix, "V_8", intRow)
            ds.Tables("Document").Rows.Add(oDRow)
            '  Print_Evalution("Evaluation", aOrderNo)
        Next
        addCrystal(ds, "Document", 1)
        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
    End Sub


    Public Sub PrintEvalution_Arabic(ByVal aOrderNo As String)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aOrderNo = "" Then
            oApplication.Utilities.Message("Evaluation details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oRec.DoQuery("Select * from [@Z_AQEVAL] where DocEntry=" & CInt(aOrderNo))
        ds.Clear()
        If oRec.RecordCount <= 0 Then
            oApplication.Utilities.Message("Evaluation details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            oDRow = ds.Tables("Evaluation").NewRow()
            oDRow.Item("ID") = aOrderNo 'oRec.Fields.Item("DocEntry").Value
            oDRow.Item("Z_ProCode") = oRec.Fields.Item("U_Z_ProCode").Value
            oDRow.Item("Z_ProName") = oRec.Fields.Item("U_Z_ProName").Value
            oDRow.Item("Z_EVL_ID") = oRec.Fields.Item("U_Z_EVL_ID").Value
            oDRow.Item("Z_EV_DATE") = oRec.Fields.Item("U_Z_EV_DATE").Value
            oDRow.Item("Z_OWNERCODE") = oRec.Fields.Item("U_Z_OWNERCODE").Value
            oDRow.Item("Z_OWNER") = oRec.Fields.Item("U_Z_OWNER").Value
            oDRow.Item("Z_OWNRTEL") = oRec.Fields.Item("U_Z_OWNRTEL").Value
            oDRow.Item("Z_OWNRGSM") = oRec.Fields.Item("U_Z_OWNRGSM").Value
            '  oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_ToLOC").Value)
            oDRow.Item("Z_OWNRPOBOX") = oRec.Fields.Item("U_Z_OWNRPOBOX").Value
            oDRow.Item("Z_OWNRADD") = oRec.Fields.Item("U_Z_OWNRADD").Value
            oDRow.Item("Z_AQTYP") = oRec.Fields.Item("U_Z_AQTYP").Value
            oDRow.Item("Z_AQBOND") = oRec.Fields.Item("U_Z_AQBOND").Value

            oDRow.Item("Z_AQCITY") = oRec.Fields.Item("U_Z_AQCITY").Value
            oDRow.Item("Z_AQZONE") = oRec.Fields.Item("U_Z_AQZONE").Value
            oDRow.Item("Z_AQSTREET") = oRec.Fields.Item("U_Z_AQSTREET").Value
            oDRow.Item("Z_AQLSNC") = oRec.Fields.Item("U_Z_AQLSNC").Value
            oDRow.Item("Z_AQLSNC_DATE") = oRec.Fields.Item("U_Z_AQLSNC_DATE").Value
            oDRow.Item("Z_AQNEARST") = oRec.Fields.Item("U_Z_AQNEARST").Value
            oDRow.Item("Z_AQAGE") = oRec.Fields.Item("U_Z_AQAGE").Value
            oDRow.Item("Z_AQBUILD") = oRec.Fields.Item("U_Z_AQBUILD").Value
            oDRow.Item("Z_AQQUALITY") = oRec.Fields.Item("U_Z_AQQUALITY").Value




            oDRow.Item("Z_USRENTNAME") = oRec.Fields.Item("U_Z_USRENTNAME").Value
            oDRow.Item("Z_USRMNTHRNT") = oRec.Fields.Item("U_Z_USRMNTHRNT").Value
            oDRow.Item("Z_RENTTYP") = oRec.Fields.Item("U_Z_RENTTYP").Value
            oDRow.Item("Z_MRKTLAND") = oRec.Fields.Item("U_Z_MRKTLAND").Value
            oDRow.Item("Z_MRKTZFOOT") = oRec.Fields.Item("U_Z_MRKTZFOOT").Value
            oDRow.Item("Z_BUILTCOST") = oRec.Fields.Item("U_Z_BUILTCOST").Value
            oDRow.Item("Z_MOREVALS") = oRec.Fields.Item("U_Z_MOREVALS").Value
            oDRow.Item("Z_AQAREA") = oRec.Fields.Item("U_Z_AQAREA").Value
            oDRow.Item("Z_AQAREAMT") = oRec.Fields.Item("U_Z_AQAREAMT").Value



            oDRow.Item("Z_BLDAREAFT") = oRec.Fields.Item("U_Z_BLDAREAFT").Value
            oDRow.Item("Z_FORCPER") = oRec.Fields.Item("U_Z_FORCPER").Value
            oDRow.Item("Z_MRKTBLDMT") = oRec.Fields.Item("U_Z_MRKTBLDMT").Value
            oDRow.Item("Z_BUILDAREA") = oRec.Fields.Item("U_Z_BUILDAREA").Value
            oDRow.Item("Z_REQ") = oRec.Fields.Item("U_Z_REQ").Value
            oDRow.Item("Z_RQTEL") = oRec.Fields.Item("U_Z_RQTEL").Value
            oDRow.Item("Z_RQGSM") = oRec.Fields.Item("U_Z_RQGSM").Value
            oDRow.Item("Z_RQADD") = oRec.Fields.Item("U_Z_RQADD").Value
            oDRow.Item("Z_RQPOBOX") = oRec.Fields.Item("U_Z_RQPOBOX").Value



            oDRow.Item("Z_REQEMAIL") = oRec.Fields.Item("U_Z_REQEMAIL").Value
            oDRow.Item("Z_REQFAX") = oRec.Fields.Item("U_Z_REQFAX").Value
            oDRow.Item("Z_AQDESC") = oRec.Fields.Item("U_Z_AQDESC").Value
            oDRow.Item("Z_AQNOTES") = oRec.Fields.Item("U_Z_AQNOTES").Value
            oDRow.Item("Z_AQFACILITY") = oRec.Fields.Item("U_Z_AQFACILITY").Value
            oDRow.Item("Z_MOREFAC") = oRec.Fields.Item("U_Z_MOREFAC").Value
            oDRow.Item("Z_USEINFO") = oRec.Fields.Item("U_Z_USEINFO").Value
            oDRow.Item("Z_RENTERNAME") = oRec.Fields.Item("U_Z_RENTERNAME").Value
            oDRow.Item("Z_CHKRNAME") = oRec.Fields.Item("U_Z_CHKRNAME").Value



            oDRow.Item("Z_USERNAME") = oRec.Fields.Item("U_Z_USERNAME").Value
            oDRow.Item("Z_RCHKR") = oRec.Fields.Item("U_Z_RCHKR").Value
            oDRow.Item("Z_EVCHRG") = oRec.Fields.Item("U_Z_EVCHRG").Value
            oDRow.Item("Z_INVNO") = oRec.Fields.Item("U_Z_INVNO").Value
            oDRow.Item("Z_DBRCPER") = oRec.Fields.Item("U_Z_DBRCPER").Value
            oDRow.Item("Z_AC_ID") = oRec.Fields.Item("U_Z_AC_ID").Value
            oDRow.Item("Z_NOTES") = oRec.Fields.Item("U_Z_NOTES").Value
            oDRow.Item("Z_AQZNOTES") = oRec.Fields.Item("U_Z_AQZNOTES").Value
            ' oDRow.Item("Z_CHKRNAME") = oTemp.Fields.Item("CardName").Value

            ds.Tables("Evaluation").Rows.Add(oDRow)
            addCrystal(ds, "Evaluation_Arabic", aOrderNo)
            '  Print_Evalution("Evaluation", aOrderNo)
        End If
        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
    End Sub


    Private Function BuildQuery(ByVal aDate As String, ByVal bDate As String) As String
        Dim StrString As String
        'StrString = "   select x.U_Model,x.U_RegNo,avg(x.Quantity) 'Number of Days',avg(x.LineTotal) 'Total Revenue',"
        'StrString = StrString & " isnull(sum(X.Cash), 0) 'Paid by Cash',isnull(Sum(X.Check1),0) 'Paid by Cheque',isnull(Sum(X.Bank),0) 'Paid by Bank',isnull(Sum(x.Credit),0) 'paid by Credit Card',"
        'StrString = StrString & " isnull(sum(X.Cash),0)+isnull(Sum(X.Check1),0) +isnull(Sum(X.Bank),0)+ isnull(Sum(x.Credit),0) 'Total Collection'"
        'StrString = StrString & " , avg(x.LineTotal)-(isnull(sum(X.Cash),0)+isnull(Sum(X.Check1),0) +isnull(Sum(X.Bank),0)+ isnull(Sum(x.Credit),0)) 'Total Receivable'"
        'StrString = StrString & " from "
        'StrString = StrString & "( select   T0.SlpCode,T1.U_RegNo,T1.U_Model, T0.DocDate,T0.Quantity, T0.LineTotal, T2.SumApplied,"
        'StrString = StrString & "  case  when (LineTotal - isnull(T2.SumApplied,0) ) < = 0 then LineTotal else  isnull(T2.SumApplied,0) end 'TotalPayment',"
        'StrString = StrString & " T2.CashSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CashSum)<=0 then Linetotal  else isnull(T2.CashSum ,0)end else case when (isnull(T2.SumApplied,0)-T2.CashSum)<=0 then  isnull(T2.SumApplied,0) else T2.CashSum end end 'Cash',"
        'StrString = StrString & " T2.CheckSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CheckSum)<=0 then Linetotal  else isnull(T2.CheckSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.CheckSum)<=0 then  isnull(T2.SumApplied,0) else T2.CheckSum end end 'Check1',"
        'StrString = StrString & " T2.TrsfrSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.TrsfrSum)<=0 then Linetotal  else isnull(T2.TrsfrSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.TrsfrSum)<=0 then  isnull(T2.SumApplied,0) else T2.TrsfrSum end end 'Bank',"
        'StrString = StrString & " T2.CreditSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CreditSum)<=0 then Linetotal  else isnull(T2.CreditSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.CreditSum)<=0 then  isnull(T2.SumApplied,0) else T2.CreditSum end end 'Credit'"
        'StrString = StrString & " From    OITM T1 Left outer  Join "
        'StrString = StrString & "("
        'StrString = StrString & "  Select  T4.SlpCode,T4.CardCode,T4.DocDate, T4.DocEntry, T0.ItemCode, T0.Quantity, T0.LineTotal"
        'StrString = StrString & " From    INV1 T0 Inner Join OINV T4 on T4.DocEntry=T0.DocEntry  And " & aDate '(T0.docdate) >= '2013-02-20' and (T0.DocDate) <='2013-02-21'"
        'StrString = StrString & ") T0 on T0.ItemCode=T1.ItemCode "
        'StrString = StrString & "  Left outer  Join ("
        'StrString = StrString & " Select  T2.DocEntry, T2.SumApplied, T3.CashSum, T3.CheckSum, T3.TrsfrSum, T3.CreditSum"
        'StrString = StrString & " From    RCT2 T2 Inner Join ORCT T3 on T3.DocEntry=T2.DocNum  and T3.Canceled='N' And T2.InvType=13 "
        'StrString = StrString & " ) T2 On T2.DocEntry=T0.DocEntry "
        'StrString = StrString & " where T1.QryGroup1='Y'  "
        'StrString = StrString & " ) as  X  group by X.U_Model,X.U_RegNo "

        StrString = " select X.U_Model,X.U_RegNo,isnull(sum(x.Quantity),0) 'Days',isnull(sum(x.LineTotal),0) 'Total Value',sum(isnull(T2.Applied,0)) 'Applied SUm' ,sum(isnull(T2.Cash,0)) 'Cash',sum(isnull(T2.Cheqeue,0)) 'Check',sum(isnull(T2.Bank,0)) 'bank',sum(isnull(T2.Credit1,0))  'Credit Card' from "
        StrString = StrString & " ( select   T0.SlpCode, T0.DocEntry, T0.DocDate,T1.U_RegNo,T1.U_Model, T0.Quantity, T0.LineTotal  From    OITM T1 Left outer  Join ( Select  T4.SlpCode, T4.DocEntry, T4.DocDate,T0.ItemCode, T0.Quantity, T0.LineTotal From    INV1 T0 Inner Join OINV T4 on T4.DocEntry=T0.DocEntry And " & aDate ' (T0.docdate) >= '2013-02-20' and (T0.DocDate) <='2013-02-21' "
        StrString = StrString & " ) T0 on T0.ItemCode=T1.ItemCode  where T1.QryGroup1='Y'  ) as X left outer join "
        StrString = StrString & " ("
        StrString = StrString & " Select  T2.DocEntry, sum(T2.SumApplied) 'Applied', sum(T3.CashSum) 'Cash', sum(T3.CheckSum) 'Cheqeue', sum(T3.TrsfrSum) 'Bank',sum(T3.CreditSum)'Credit1'"
        StrString = StrString & " From    RCT2 T2 Inner Join ORCT T3 on T3.DocEntry=T2.DocNum  and T3.Canceled='N' And T2.InvType=13  and " & bDate & " group by T2.DocEntry"
        StrString = StrString & ") T2 On  T2.DocEntry=X.DocEntry group by x.U_Model,X.U_RegNo											"


        'With RollUp
        Return StrString

    End Function

    Private Function BuildQuery_SalesAgent(ByVal aDate As String, ByVal bDate As String) As String
        Dim StrString As String
        'StrString = "   select T7.SlpName,x.CardName,X.DocNum,Sum(x.Quantity) 'Number of Days',Sum(x.LineTotal) 'Total Revenue',"
        'StrString = StrString & " isnull(sum(X.Cash), 0) 'Paid by Cash',isnull(Sum(X.Check1),0) 'Paid by Cheque',isnull(Sum(X.Bank),0) 'Paid by Bank',isnull(Sum(x.Credit),0) 'paid by Credit Card',"
        'StrString = StrString & " isnull(sum(X.Cash),0)+isnull(Sum(X.Check1),0) +isnull(Sum(X.Bank),0)+ isnull(Sum(x.Credit),0) 'Total Collection'"
        'StrString = StrString & " , avg(x.LineTotal)-(isnull(sum(X.Cash),0)+isnull(Sum(X.Check1),0) +isnull(Sum(X.Bank),0)+ isnull(Sum(x.Credit),0)) 'Total Receivable'"
        'StrString = StrString & " from "
        'StrString = StrString & "( select   T0.SlpCode,T0.CardName,T0.DocNum,T1.U_RegNo,T1.U_Model, T0.DocDate,T0.Quantity, T0.LineTotal, T2.SumApplied,"
        'StrString = StrString & "  case  when (LineTotal - isnull(T2.SumApplied,0) ) < = 0 then LineTotal else  isnull(T2.SumApplied,0) end 'TotalPayment',"
        'StrString = StrString & " T2.CashSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CashSum)<=0 then Linetotal  else isnull(T2.CashSum ,0)end else case when (isnull(T2.SumApplied,0)-T2.CashSum)<=0 then  isnull(T2.SumApplied,0) else T2.CashSum end end 'Cash',"
        'StrString = StrString & " T2.CheckSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CheckSum)<=0 then Linetotal  else isnull(T2.CheckSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.CheckSum)<=0 then  isnull(T2.SumApplied,0) else T2.CheckSum end end 'Check1',"
        'StrString = StrString & " T2.TrsfrSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.TrsfrSum)<=0 then Linetotal  else isnull(T2.TrsfrSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.TrsfrSum)<=0 then  isnull(T2.SumApplied,0) else T2.TrsfrSum end end 'Bank',"
        'StrString = StrString & " T2.CreditSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CreditSum)<=0 then Linetotal  else isnull(T2.CreditSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.CreditSum)<=0 then  isnull(T2.SumApplied,0) else T2.CreditSum end end 'Credit'"
        'StrString = StrString & " From    OITM T1 Left outer  Join "
        'StrString = StrString & "("
        'StrString = StrString & "  Select  T4.SlpCode,T4.CardCode,T4.DocNum,T4.DocDate, T4.DocEntry, T0.ItemCode, T0.Quantity, T0.LineTotal"
        'StrString = StrString & " From    INV1 T0 Inner Join OINV T4 on T4.DocEntry=T0.DocEntry  And " & aDate '(T0.docdate) >= '2013-02-20' and (T0.DocDate) <='2013-02-21'"
        'StrString = StrString & ") T0 on T0.ItemCode=T1.ItemCode "
        'StrString = StrString & "  Left outer  Join ("
        'StrString = StrString & " Select  T2.DocEntry, T2.SumApplied, T3.CashSum, T3.CheckSum, T3.TrsfrSum, T3.CreditSum"
        'StrString = StrString & " From    RCT2 T2 Inner Join ORCT T3 on T3.DocEntry=T2.DocNum  and T3.Canceled='N' And T2.InvType=13 "
        'StrString = StrString & " ) T2 On T2.DocEntry=T0.DocEntry "
        'StrString = StrString & " where T1.QryGroup1='Y'  "
        'StrString = StrString & " ) as  X  inner join OSLP T7 on T7.SlpCode=X.SlpCode group by T7.SlpName,x.CardName,X.DocNum"
        ''With RollUp


        StrString = "select isnull(T6.Slpname,''),x.DocEntry,isnull(sum(x.Quantity),0) 'Days',isnull(sum(x.LineTotal),0) 'Total Value',sum(isnull(T2.Applied,0)) 'Applied SUm' ,"
        StrString = StrString & "    sum(isnull(T2.Cash, 0)) 'Cash',sum(isnull(T2.Cheqeue,0)) 'Check',sum(isnull(T2.Bank,0)) 'bank',sum(isnull(T2.Credit1,0))  'Credit Card',X.CardName 'CardName' from "
        StrString = StrString & " (   select   T0.SlpCode, T0.DocEntry,T0.CardName, T0.DocDate,T1.U_RegNo,T1.U_Model, T0.Quantity, T0.LineTotal   From    OITM T1 Left outer  Join "
        StrString = StrString & " (  Select  T4.SlpCode, T4.DocEntry,T4.CardName, T4.DocDate,T0.ItemCode, T0.Quantity, T0.LineTotal From    INV1 T0 Inner Join OINV T4 on T4.DocEntry=T0.DocEntry  and " & aDate 'And (T0.docdate) >= '2013-02-20' and (T0.DocDate) <='2013-02-21'"
        StrString = StrString & " ) T0 on T0.ItemCode=T1.ItemCode  where T1.QryGroup1='Y'  ) as X"
        StrString = StrString & " left outer join "
        StrString = StrString & " (  Select  T2.DocEntry, sum(T2.SumApplied) 'Applied', sum(T3.CashSum) 'Cash', sum(T3.CheckSum) 'Cheqeue', sum(T3.TrsfrSum) 'Bank',sum(T3.CreditSum)'Credit1'"
        StrString = StrString & " From    RCT2 T2 Inner Join ORCT T3 on T3.DocEntry=T2.DocNum  and T3.Canceled='N' And T2.InvType=13 and " & bDate & "  group by T2.DocEntry ) T2 On  T2.DocEntry=X.DocEntry "
        StrString = StrString & " left outer join OSLP T6 on T6.SlpCode=X.SlpCode group by T6.Slpname,X.CardName,x.DocEntry order by T6.Slpname"

        Return StrString

    End Function

    Private Function BuildQuery_Production(ByVal aDate As String, ByVal bDate As String) As String
        Dim StrString As String
        'StrString = "   select T7.SlpName,x.U_Model,Sum(x.Quantity/x.Quantity) 'Number of Days',Sum(x.LineTotal/x.quantity) 'Total Revenue',"
        'StrString = StrString & " isnull(sum(X.Cash), 0) 'Paid by Cash',isnull(Sum(X.Check1),0) 'Paid by Cheque',isnull(Sum(X.Bank),0) 'Paid by Bank',isnull(Sum(x.Credit),0) 'paid by Credit Card',"
        'StrString = StrString & " isnull(sum(X.Cash),0)+isnull(Sum(X.Check1),0) +isnull(Sum(X.Bank),0)+ isnull(Sum(x.Credit),0) 'Total Collection'"
        'StrString = StrString & " , Sum(x.LineTotal/x.quantity) -(isnull(sum(X.Cash),0)+isnull(Sum(X.Check1),0) +isnull(Sum(X.Bank),0)+ isnull(Sum(x.Credit),0)) 'Total Receivable'"
        'StrString = StrString & " from "
        'StrString = StrString & "( select   T0.SlpCode,T1.U_RegNo,T1.U_Model, T0.DocDate,T0.Quantity, T0.LineTotal, T2.SumApplied,"
        'StrString = StrString & "  case  when (LineTotal - isnull(T2.SumApplied,0) ) < = 0 then LineTotal else  isnull(T2.SumApplied,0) end 'TotalPayment',"
        'StrString = StrString & " T2.CashSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CashSum)<=0 then Linetotal  else isnull(T2.CashSum ,0)end else case when (isnull(T2.SumApplied,0)-T2.CashSum)<=0 then  isnull(T2.SumApplied,0) else T2.CashSum end end 'Cash',"
        'StrString = StrString & " T2.CheckSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CheckSum)<=0 then Linetotal  else isnull(T2.CheckSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.CheckSum)<=0 then  isnull(T2.SumApplied,0) else T2.CheckSum end end 'Check1',"
        'StrString = StrString & " T2.TrsfrSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.TrsfrSum)<=0 then Linetotal  else isnull(T2.TrsfrSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.TrsfrSum)<=0 then  isnull(T2.SumApplied,0) else T2.TrsfrSum end end 'Bank',"
        'StrString = StrString & " T2.CreditSum, case  when (LineTotal -  isnull(T2.SumApplied,0)) < = 0 then case when (LineTotal -T2.CreditSum)<=0 then Linetotal  else isnull(T2.CreditSum,0) end else case when  (isnull(T2.SumApplied,0)-T2.CreditSum)<=0 then  isnull(T2.SumApplied,0) else T2.CreditSum end end 'Credit'"
        'StrString = StrString & " From    OITM T1 Left outer  Join "
        'StrString = StrString & "("
        'StrString = StrString & "  Select  T4.SlpCode,T4.CardCode,T4.DocDate, T4.DocEntry, T0.ItemCode, T0.Quantity, T0.LineTotal"
        'StrString = StrString & " From    INV1 T0 Inner Join OINV T4 on T4.DocEntry=T0.DocEntry  And " & aDate '(T0.docdate) >= '2013-02-20' and (T0.DocDate) <='2013-02-21'"
        'StrString = StrString & ") T0 on T0.ItemCode=T1.ItemCode "
        'StrString = StrString & "  Left outer  Join ("
        'StrString = StrString & " Select  T2.DocEntry, T2.SumApplied, T3.CashSum, T3.CheckSum, T3.TrsfrSum, T3.CreditSum"
        'StrString = StrString & " From    RCT2 T2 Inner Join ORCT T3 on T3.DocEntry=T2.DocNum  and T3.Canceled='N' And T2.InvType=13 "
        'StrString = StrString & " ) T2 On T2.DocEntry=T0.DocEntry "
        'StrString = StrString & " where T1.QryGroup1='Y'  "
        'StrString = StrString & " ) as  X  inner join OSLP T7 on T7.SlpCode=X.SlpCode group by T7.SlpName,x.U_Model"

        StrString = " select isnull(T6.Slpname,''),X.U_Model,isnull(sum(x.Quantity),0) 'Days',isnull(sum(x.LineTotal),0) 'Total Value',sum(isnull(T2.Applied,0)) 'Applied SUm' ,"
        StrString = StrString & "sum(isnull(T2.Cash,0)) 'Cash',sum(isnull(T2.Cheqeue,0)) 'Check',sum(isnull(T2.Bank,0)) 'bank',sum(isnull(T2.Credit1,0))  'Credit Card' from "
        StrString = StrString & "(  select   T0.SlpCode, T0.DocEntry,T0.CardName, T0.DocDate,T1.U_RegNo,T1.U_Model, T0.Quantity, T0.LineTotal  From    OITM T1 Left outer  Join "
        StrString = StrString & "(  Select  T4.SlpCode, T4.DocEntry,T4.CardName, T4.DocDate,T0.ItemCode, T0.Quantity, T0.LineTotal  From    INV1 T0 Inner Join OINV T4 on T4.DocEntry=T0.DocEntry And " & aDate
        StrString = StrString & ") T0 on T0.ItemCode=T1.ItemCode  where T1.QryGroup1='Y'  ) as X"
        StrString = StrString & " left outer join "
        StrString = StrString & "(  Select  T2.DocEntry, sum(T2.SumApplied) 'Applied', sum(T3.CashSum) 'Cash', sum(T3.CheckSum) 'Cheqeue', sum(T3.TrsfrSum) 'Bank',sum(T3.CreditSum)'Credit1'"
        StrString = StrString & "From    RCT2 T2 Inner Join ORCT T3 on T3.DocEntry=T2.DocNum  and T3.Canceled='N' And T2.InvType=13 and " & bDate & "  group by T2.DocEntry ) T2 On  T2.DocEntry=X.DocEntry "
        StrString = StrString & "left outer join OSLP T6 on T6.SlpCode=X.SlpCode group by T6.Slpname,X.U_Model order by T6.Slpname"
        'With RollUp
        Return StrString

    End Function

    Public Sub PrintReciableReport(ByVal adate As Date, ByVal ToDate As Date, ByVal rptType As String)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ' oRec.DoQuery("Select * from [@Z_ORDR] where DocEntry=" & aOrderNo)
        Dim strStrin,strStrin1 As String
        strStrin = " T0.DocDate >='" & adate.ToString("yyyy-MM-dd") & "' and T0.DocDate <='" & ToDate.ToString("yyyy-MM-dd") & "'"
        strStrin1 = " T3.DocDate >='" & adate.ToString("yyyy-MM-dd") & "' and T3.DocDate <='" & ToDate.ToString("yyyy-MM-dd") & "'"
       
        Dim strMonth As String
        strMonth = MonthName(Month(adate))
        If rptType = "R" Then
            strStrin = BuildQuery(strStrin, strStrin1)
        ElseIf rptType = "S" Then
            strStrin = BuildQuery_SalesAgent(strStrin, strStrin1)
        ElseIf (rptType = "P") Then
            strStrin = BuildQuery_Production(strStrin, strStrin1)
        End If

        oRec.DoQuery(strStrin)
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        If oRec.RecordCount <= 0 Then
            oApplication.Utilities.Message("Booking details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            Dim dblTotalValue, dblinvoicevalue, DblAPplied, dblCash, dblcheck, dblBank, dblCard As Double
            For intRow As Integer = 0 To oRec.RecordCount - 1
                oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oDRow = ds.Tables("Receiable").NewRow()
                oDRow.Item("Month") = strMonth
                oDRow.Item("Year") = Year(adate)
                If rptType <> "S" Then
                    oDRow.Item("Customer") = ""
                Else
                    oDRow.Item("Customer") = oRec.Fields.Item("CardName").Value
                End If

                oDRow.Item("Modal") = oRec.Fields.Item(0).Value
                oDRow.Item("RegNo") = oRec.Fields.Item(1).Value
                oDRow.Item("Days") = oRec.Fields.Item(2).Value
                oDRow.Item("TotalValue") = oRec.Fields.Item(3).Value
                dblTotalValue = oRec.Fields.Item(3).Value
                dblinvoicevalue = dblTotalValue
                dblCash = oRec.Fields.Item(4).Value
                dblcheck = oRec.Fields.Item(5).Value
                dblBank = oRec.Fields.Item(6).Value
                dblCard = oRec.Fields.Item(7).Value

                If dblTotalValue <= dblCash Then
                    dblCash = dblTotalValue
                    dblTotalValue = 0
                Else
                    dblCash = dblCash
                    dblTotalValue = dblTotalValue - dblCash
                End If
                If dblTotalValue <= dblcheck Then
                    dblcheck = dblTotalValue
                    dblTotalValue = 0
                Else
                    dblcheck = dblcheck
                    dblTotalValue = dblTotalValue - dblcheck
                End If
                If dblTotalValue <= dblBank Then
                    dblBank = dblTotalValue
                    dblTotalValue = 0
                Else
                    dblBank = dblBank
                    dblTotalValue = dblTotalValue - dblBank
                End If
                If dblTotalValue <= dblCard Then
                    dblCard = dblTotalValue
                    dblTotalValue = 0
                Else
                    dblCard = dblCard
                    dblTotalValue = dblTotalValue - dblCard
                End If
                oDRow.Item("Cash") = dblCash
                oDRow.Item("Cheque") = dblcheck
                oDRow.Item("Bank") = dblBank
                oDRow.Item("Card") = dblCard


                'oDRow.Item("Cash") = oRec.Fields.Item(4).Value
                'oDRow.Item("Cheque") = oRec.Fields.Item(5).Value
                'oDRow.Item("Bank") = oRec.Fields.Item(6).Value
                'oDRow.Item("Card") = oRec.Fields.Item(7).Value

                oDRow.Item("TotalCollection") = dblCash + dblcheck + dblBank + dblCard ' oRec.Fields.Item(8).Value
                oDRow.Item("Balance") = dblinvoicevalue - (dblCash + dblcheck + dblBank + dblCard) ' oRec.Fields.Item(9).Value
                ds.Tables("Receiable").Rows.Add(oDRow)
                oRec.MoveNext()
            Next
            'addCrystal(ds, "Agreement", aOrderNo)
            If rptType = "R" Then
                addCrystal(ds, "Receiable", "")
            ElseIf rptType = "S" Then
                addCrystal(ds, "SalesAgent", "")
            ElseIf (rptType = "P") Then
                addCrystal(ds, "Production", "")
            End If
        End If
        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)

    End Sub
    Public Sub PrintRentalAgreement(ByVal aOrderNo As Integer)
        Dim oRec, oRecTemp, oRecBP, oBalanceRs, oTemp As SAPbobsCOM.Recordset
        Dim strfrom, dtPosting, dtdue, dttax, strto, strBranch, strSlpCode, strSlpName, strSMNo, strFromBP, strToBP, straging, strCardcode, strCardname, strBlock, strCity, strBilltoDef, strZipcode, strAddress, strCounty, strPhone1, strFax, strCntctprsn, strTerrtory, strNotes As String
        Dim dtFrom, dtTo, dtAging As Date
        Dim intReportChoice As Integer
        Dim dblRef1, dblCredit, dblDebit, dblCumulative, dblOpenBalance As Double
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from [@Z_ORDR] where DocEntry=" & aOrderNo)
        If oRec.RecordCount <= 0 Then
            oApplication.Utilities.Message("Booking details does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        Else
            oDRow = ds.Tables("Agreement").NewRow()
            oDRow.Item("BookId") = oRec.Fields.Item("DocEntry").Value
            oDRow.Item("BookDate") = oRec.Fields.Item("U_Z_BookingDate").Value
            oDRow.Item("CardCode") = oRec.Fields.Item("U_Z_CardCode").Value
            oDRow.Item("CardName") = oRec.Fields.Item("U_Z_CardName").Value
            oDRow.Item("FrgnName") = oRec.Fields.Item("U_Z_FrgnName").Value

            ' oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_FromLOC").Value)
            oDRow.Item("PickupLocation") = oRec.Fields.Item("U_Z_FromLOC").Value
            ' oDRow.Item("PickupLocation") = oRec.Fields.Item("U_Z_FromLOC").Value

            oDRow.Item("PickupDate") = oRec.Fields.Item("U_Z_InDate").Value
            Dim intTime As Integer
            intTime = oRec.Fields.Item("U_Z_InTime").Value
            oDRow.Item("PickupTime") = intTime.ToString("00:00") ' oRec.Fields.Item("U_Z_InTime").Value
            ' oTemp.DoQuery("select isnull(U_Z_LocName,'') from [@Z_OLOC] where DocEntry =" & oRec.Fields.Item("U_Z_ToLOC").Value)
            oDRow.Item("DropLocation") = oRec.Fields.Item("U_Z_ToLOC").Value
            '  oDRow.Item("DropLocation") = oTemp.Fields.Item(0).Value
            oDRow.Item("DropDate") = oRec.Fields.Item("U_Z_OutDate").Value
            intTime = oRec.Fields.Item("U_Z_OutTime").Value
            oDRow.Item("DropTime") = intTime.ToString("00:00") ' oRec.Fields.Item("U_Z_OutTime").Value
            oDRow.Item("Currency") = oApplication.Utilities.GetLocalCurrency()
            oDRow.Item("ContactNo") = oRec.Fields.Item("U_Z_Mobile").Value
            oDRow.Item("ItemCode") = oRec.Fields.Item("U_Z_ItemCode").Value
            oDRow.Item("ItemName") = oRec.Fields.Item("U_Z_ItemName").Value
            oDRow.Item("Quantity") = oRec.Fields.Item("U_Z_Qty").Value
            oDRow.Item("DailyRate") = oRec.Fields.Item("U_Z_Amount").Value
            oDRow.Item("Total") = oRec.Fields.Item("U_Z_Total").Value
            oDRow.Item("NoofDays") = oRec.Fields.Item("U_Z_NoofDays").Value
            oDRow.Item("SecurityDeposit") = oRec.Fields.Item("U_Z_Deposit").Value
            oDRow.Item("DownPayment") = oRec.Fields.Item("U_Z_DPAmount").Value

            'New Fields
            oDRow.Item("Email") = oRec.Fields.Item("U_Z_Email").Value
            oDRow.Item("Driver") = oRec.Fields.Item("U_Z_DriverName").Value
            oDRow.Item("DLNO") = oRec.Fields.Item("U_Z_DLNo").Value
            oDRow.Item("DLPlace") = oRec.Fields.Item("U_Z_DLIssuePlace").Value
            oDRow.Item("DLDate") = oRec.Fields.Item("U_Z_DLIssueDate").Value
            oDRow.Item("DLExp") = oRec.Fields.Item("U_Z_DLExpiry").Value
            oDRow.Item("PSNO") = oRec.Fields.Item("U_Z_PSNo").Value
            oDRow.Item("PSPlace") = oRec.Fields.Item("U_Z_PSIssuePlace").Value
            oDRow.Item("PSDate") = oRec.Fields.Item("U_Z_PSIssueDate").Value
            oDRow.Item("PSExp") = oRec.Fields.Item("U_Z_PSExpirty").Value
            oDRow.Item("DLBirth") = oRec.Fields.Item("U_Z_BirthPlace").Value
            oDRow.Item("DLDOB") = oRec.Fields.Item("U_Z_DOB").Value

            oDRow.Item("Driver1") = oRec.Fields.Item("U_Z_DriverName1").Value
            oDRow.Item("DLNo1") = oRec.Fields.Item("U_Z_DLNo1").Value
            oDRow.Item("DLPlace1") = oRec.Fields.Item("U_Z_DLIssuePlace1").Value
            oDRow.Item("DLDate1") = oRec.Fields.Item("U_Z_DLIssueDate1").Value
            oDRow.Item("DLExp1") = oRec.Fields.Item("U_Z_DLExpiry1").Value
            oDRow.Item("PSNO1") = oRec.Fields.Item("U_Z_PSNo1").Value
            oDRow.Item("PSPlace1") = oRec.Fields.Item("U_Z_PSIssuePlace1").Value
            oDRow.Item("PSDate1") = oRec.Fields.Item("U_Z_PSIssueDate1").Value
            If Year(oRec.Fields.Item("U_Z_PSExpirty1").Value) = 1899 Then
                oDRow.Item("PSExp1") = ""
            Else
                oDRow.Item("PSExp1") = oRec.Fields.Item("U_Z_PSExpirty1").Value
            End If
            oDRow.Item("DLBirth1") = oRec.Fields.Item("U_Z_BirthPlace1").Value
            oDRow.Item("DLDOB1") = oRec.Fields.Item("U_Z_DOB1").Value

            oDRow.Item("UserSign") = oRec.Fields.Item("UserSign").Value
            oDRow.Item("Model") = oRec.Fields.Item("U_Z_Model").Value
            oDRow.Item("RegNo") = oRec.Fields.Item("U_Z_RegNo").Value
            oDRow.Item("Salik") = oRec.Fields.Item("U_Z_Salik").Value
            oDRow.Item("CardType") = oRec.Fields.Item("U_Z_CardType").Value
            oDRow.Item("CardNo") = oRec.Fields.Item("U_Z_CardNo").Value
            oDRow.Item("CardExp") = oRec.Fields.Item("U_Z_CardExpDate").Value
            oDRow.Item("NameofCard") = oRec.Fields.Item("U_Z_NameofCard").Value
            oDRow.Item("BillTo") = oRec.Fields.Item("U_Z_BillTo").Value
            oDRow.Item("ChkOutDt") = oRec.Fields.Item("U_Z_ChkOutDt").Value
            oDRow.Item("ChkOutMile") = oRec.Fields.Item("U_Z_ChkOutMil").Value
            ' MsgBox(oRec.Fields.Item("U_Z_ChkOutMil").Value)
            oDRow.Item("ChkOutFuel1") = oRec.Fields.Item("U_Z_ChkOutFuel").Value
            oDRow.Item("ChkOutFuel2") = "" ' oRec.Fields.Item("U_Z_ChkOutFuel2").Value

            oDRow.Item("ChkInDt") = oRec.Fields.Item("U_Z_ChkInDt").Value
            oDRow.Item("ChkInMile") = oRec.Fields.Item("U_Z_ChkInMil").Value
            '   MsgBox(oRec.Fields.Item("U_Z_ChkInFuel1").Value)
            oDRow.Item("ChkInFuel1") = oRec.Fields.Item("U_Z_ChkInFuel3").Value
            ' oDRow.Item("ChkInFuel2") = oRec.Fields.Item("U_Z_ChkInFuel2").Value






            ds.Tables("Agreement").Rows.Add(oDRow)
            'addCrystal(ds, "Agreement", aOrderNo)


            oRec.DoQuery("Select * from [@Z_RDR1] where DocEntry=" & aOrderNo)
            For intRow As Integer = 0 To oRec.RecordCount - 1
                If oRec.Fields.Item("U_Z_EqpName").Value <> "" Then
                    oDRow = ds.Tables("CheckListIN").NewRow()
                    oDRow.Item("DocEntry") = oRec.Fields.Item("DocEntry").Value
                    oDRow.Item("Name") = oRec.Fields.Item("U_Z_EqpName").Value
                    oDRow.Item("DelStatus") = oRec.Fields.Item("U_Z_DelStatus").Value
                    oDRow.Item("RetStatus") = oRec.Fields.Item("U_Z_RetStatus").Value
                    ds.Tables("CheckListIN").Rows.Add(oDRow)
                Else
                    oDRow = ds.Tables("CheckListIN").NewRow()
                    oDRow.Item("DocEntry") = oRec.Fields.Item("DocEntry").Value
                    'oDRow.Item("Name") = oRec.Fields.Item("U_Z_EqpName").Value
                    'oDRow.Item("DelStatus") = oRec.Fields.Item("U_Z_DelStatus").Value
                    'oDRow.Item("RetStatus") = oRec.Fields.Item("U_Z_RetStatus").Value
                    ds.Tables("CheckListIN").Rows.Add(oDRow)
                End If
                oRec.MoveNext()
            Next
            
            addCrystal(ds, "Agreement", aOrderNo)
        End If
        oApplication.Utilities.Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)

    End Sub
#End Region

    
End Class
