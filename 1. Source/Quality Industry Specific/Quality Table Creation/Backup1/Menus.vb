Public Class Menus
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public oCompany As SAPbobsCOM.Company
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
#Region "Menus"

    Public Sub Intialize()
        Dim objSBOConnector As New SST.SBOConnector
        SBO_Application = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        oCompany = objSBOConnector.GetCompany(SBO_Application)
        createTables()
        CreateUDOS()
        SBO_Application.SetStatusBarMessage("QCTables Creation Add-On Connected Succesfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub
    Private Sub createTables()
        Dim objUDFEngine As New SST.UDFEngine(oCompany)
        SBO_Application.SetStatusBarMessage("Creating Tables Please Wait", SAPbouiCOM.BoMessageTime.bmt_Long, False)
        Try
            objUDFEngine.CreateTable("SST_QCREASON", "Reason", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_QCREASON", "desc", "Reason Desc", 40)
            objUDFEngine.AddAlphaField("@SST_QCREASON", "catcode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_QCREASON", "catdesc", "Category Name", 40)

            objUDFEngine.CreateTable("SST_PRDCAT", "Reason Category", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.CreateTable("SST_PARACAT", "Parameter Category", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.CreateTable("SST_MODE", "Mode of Transport", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.CreateTable("SST_STAGE", "Stages", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.CreateTable("SST_QCPARAMETER", "QC Parameter", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            objUDFEngine.AddAlphaField("@SST_QCPARAMETER", "paradesc", "Parameter Description", 40)
            objUDFEngine.AddAlphaField("@SST_QCPARAMETER", "catcode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_QCPARAMETER", "paramcat", "Category Name", 40)
            objUDFEngine.AddAlphaField("@SST_QCPARAMETER", "uomcode", "UOM Code", 20)
            objUDFEngine.AddAlphaField("@SST_QCPARAMETER", "uomdesc", "Uom Desc", 40)

            objUDFEngine.CreateTable("SST_QCUOM", "UOM Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            objUDFEngine.CreateTable("ss_cat", "Category Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            objUDFEngine.CreateTable("SST_SAMPLINGLEVEL", "Sampling Level Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            objUDFEngine.CreateTable("SST_ACCEPETLIMIT", "Acceptance Quality Limit ", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            objUDFEngine.CreateTable("SST_REF", "Reference", SAPbobsCOM.BoUTBTableType.bott_MasterData)

            objUDFEngine.CreateTable("SST_PLANHDR", "Sampling Plan-Prod", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "itemcode", "Item Code", 40)
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "itemname", "Item Name", 200)
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "stage", "Stage", 40)
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "stagedes", "Stage Description", 250)
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "Ref", "Reference", 50)
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "Active", "Active", 1)
            '*************
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "ccode", "Card Code", 20)
            objUDFEngine.AddAlphaField("@SST_PLANHDR", "cname", "Card Name", 200)
            '*************
            objUDFEngine.CreateTable("SST_PLANDTL", "Sampling Plan Detail-Prod", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddFloatField("@SST_PLANDTL", "fromqty", "From Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_PLANDTL", "toqty", "To Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "uomcode", "UOM Code", 20)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "uomdesc", "UOM Desc", 60)

            '*************
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "catcode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "paracode", "Parameter Code", 20)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "paradesc", "Parameter Description", 200)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "samplvl", "Sampling Level ", 20)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "percent", "Percentage", 20)
            objUDFEngine.AddNumericField("@SST_PLANDTL", "smpsize", "Sampling Size ", 3)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "accpqty", "Accepted Qty", 20)
            objUDFEngine.AddAlphaField("@SST_PLANDTL", "rejqty", "Rejected Qty", 20)


            '*********
            objUDFEngine.CreateTable("SST_SLMHDR", "Sampling Level Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_SLMHDR", "frmbth", "From Batch", 40)
            objUDFEngine.AddAlphaField("@SST_SLMHDR", "tobth", "To Batch", 40)
            objUDFEngine.AddAlphaField("@SST_SLMHDR", "Docnum", "Doc.Number", 40)
            objUDFEngine.AddAlphaField("@SST_SLMHDR", "Active", "Active", 1)

            objUDFEngine.CreateTable("SST_SLMDTL", "Sampling Level Detail", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@SST_SLMDTL", "lot", "Lot or Batch", 20)
            objUDFEngine.AddAlphaField("@SST_SLMDTL", "smpsize", "Sample Size", 60)

            objUDFEngine.CreateTable("SST_AQLHDR", "Acceptance Quality Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_AQLHDR", "smpsize", "Sample Size", 40)
            objUDFEngine.AddAlphaField("@SST_AQLHDR", "Active", "Active", 1)

            objUDFEngine.CreateTable("SST_AQLDTL", "Acceptance Quality Detail", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@SST_AQLDTL", "smpsize", "Sample Size", 60)
            objUDFEngine.AddAlphaField("@SST_AQLDTL", "Percent", "Percentage", 60)
            objUDFEngine.AddAlphaField("@SST_AQLDTL", "accepted", "Accepted", 60)
            objUDFEngine.AddAlphaField("@SST_AQLDTL", "rejected", "Rejected", 60)

            '********

            objUDFEngine.CreateTable("SST_QCSTANDHDR", "Item Parameter-Inward", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_QCSTANDHDR", "itemcode", "Item Code", 20)
            objUDFEngine.AddAlphaField("@SST_QCSTANDHDR", "itemname", "Item Name", 200)
            objUDFEngine.AddAlphaField("@SST_QCSTANDHDR", "Active", "Active", 1)

            objUDFEngine.CreateTable("SST_QCSTANDDTL", "Item Parameter-Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@SST_QCSTANDDTL", "catcode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_QCSTANDDTL", "paramcat", "Category Name", 40)
            objUDFEngine.AddAlphaField("@SST_QCSTANDDTL", "paracode", "Parameter code", 20)
            objUDFEngine.AddAlphaField("@SST_QCSTANDDTL", "paradesc", "Parameter Desc", 60)
            objUDFEngine.AddAlphaField("@SST_QCSTANDDTL", "uomcode", "Uom Code", 20)
            objUDFEngine.AddAlphaField("@SST_QCSTANDDTL", "uomdesc", "Uom desc", 40)
            objUDFEngine.AddFloatField("@SST_QCSTANDDTL", "value", "Parameter Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_QCSTANDDTL", "tollplus", "Tollerence Plus", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_QCSTANDDTL", "tollmins", "Tollerence Minus", SAPbobsCOM.BoFldSubTypes.st_Quantity)

            objUDFEngine.CreateTable("SST_PRDSTDHDR", "Item Parameter-Production", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_PRDSTDHDR", "itemcode", "Item Code", 40)
            objUDFEngine.AddAlphaField("@SST_PRDSTDHDR", "itemname", "Item Name", 250)
            objUDFEngine.AddAlphaField("@SST_PRDSTDHDR", "stage", "Stage", 40)
            objUDFEngine.AddAlphaField("@SST_PRDSTDHDR", "stagedes", "Stage Description", 250)
            objUDFEngine.AddAlphaField("@SST_PRDSTDHDR", "Active", "Active", 1)

            objUDFEngine.CreateTable("SST_PRDSTDDTL", "Item Parameter-Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@SST_PRDSTDDTL", "catcode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_PRDSTDDTL", "catname", "Category Name", 60)
            objUDFEngine.AddAlphaField("@SST_PRDSTDDTL", "paracode", "Parameter code", 20)
            objUDFEngine.AddAlphaField("@SST_PRDSTDDTL", "paradesc", "Parameter Desc", 60)
            objUDFEngine.AddAlphaField("@SST_PRDSTDDTL", "uomcode", "Uom Code", 20)
            objUDFEngine.AddAlphaField("@SST_PRDSTDDTL", "uomdesc", "Uom desc", 40)
            objUDFEngine.AddFloatField("@SST_PRDSTDDTL", "value", "Parameter Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_PRDSTDDTL", "tollplus", "Tollerence Plus", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_PRDSTDDTL", "tollmins", "Tollerence Minus", SAPbobsCOM.BoFldSubTypes.st_Quantity)

            objUDFEngine.CreateTable("SST_NPLANHDR", "Sampling Plan-Inward", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "itemcode", "Item Code", 20)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "itemname", "Item Name", 200)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "ItmGrp", "Item Group", 20)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "ItmGrpNM", "Item Group Name", 200)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "supcode", "Supplier code", 20)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "supname", "Supplier Name", 250)
            objUDFEngine.addField("@SST_NPLANHDR", "Opts", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "I,G", "ItemMaster,ItemGroup", "")
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "Supl", "Supplier Wise", 5)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "Ref", "Reference", 50)
            objUDFEngine.AddAlphaField("@SST_NPLANHDR", "Active", "Active", 1)


            objUDFEngine.CreateTable("SST_NPLANDTL", "Sampling Plan-Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            objUDFEngine.AddAlphaField("@SST_NPLANDTL", "catcode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_NPLANDTL", "paracode", "Parameter code", 20)
            objUDFEngine.AddAlphaField("@SST_NPLANDTL", "paradesc", "Parameter Des", 60)
            objUDFEngine.AddFloatField("@SST_NPLANDTL", "fromqty", "From Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_NPLANDTL", "toqty", "To Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("@SST_NPLANDTL", "uomcode", "UOM Code", 20)
            objUDFEngine.AddNumericField("@SST_NPLANDTL", "smpsize", "Sampling Size ", 3)
            objUDFEngine.AddNumericField("@SST_NPLANDTL", "AccNo", "Accept Number", 3)
            objUDFEngine.AddNumericField("@SST_NPLANDTL", "RejNo", "Reject Number", 3)
            objUDFEngine.AddAlphaField("@SST_NPLANDTL", "smplvl", "Sampling level", 20)
            objUDFEngine.AddAlphaField("@SST_NPLANDTL", "percen", "Percentage", 60)

            objUDFEngine.CreateTable("SST_NQCHDR", "Inward Inspection", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "insno", "Inspec.Number", 20)
            objUDFEngine.AddDateField("@SST_NQCHDR", "insdate", "Inspec.Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "GENo", "GRN No.", 20)
            objUDFEngine.AddDateField("@SST_NQCHDR", "GEDate", "GRN date.", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddFloatField("@SST_NQCHDR", "GEqty", "GRN Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "itemcode", "Item code", 20)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "itemname", "Item Name", 200)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "uom", "Item UOM", 5)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "supcode", "Supplier code", 20)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "supname", "Supplier Name", 250)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "Status", "Status", 3)
            objUDFEngine.AddAlphaField("@SST_NQCHDR", "Remarks", "Remarks", 250)
            objUDFEngine.addField("@SST_NQCHDR", "DocStat", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "D,O,C", "Draft,Open,Closed", "D")

            objUDFEngine.CreateTable("SST_NQCDTL", "Inward Inspection-Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            '*******
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "paradesc", "Parameter Desc", 250)
            objUDFEngine.AddNumericField("@SST_NQCDTL", "smpsize", "Sample No ", 5)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "acclvl", "Accepted Level", 20)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "rejlvl", "Rejected Level", 20)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "accqty", "Accepted Qty", 20)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "rejqty", "Rejected Qty", 20)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "obser", "Observations", 250)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "rmks", "Remarks", 250)
            '*******
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "CatCode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "pcode", "Parameter Code", 20)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "uomcode", "Uom code", 20)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "uom", "Uom Name", 40)
            objUDFEngine.AddFloatField("@SST_NQCDTL", "usl", "USL", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_NQCDTL", "lsl", "LSL", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_NQCDTL", "value", "Actual Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

            objUDFEngine.AddAlphaField("@SST_NQCDTL", "Status", "Status", 30)
            objUDFEngine.AddAlphaField("@SST_NQCDTL", "percen", "Percentage", 20)


            objUDFEngine.CreateTable("SST_CONSHDR", "Consolidation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("@SST_CONSHDR", "geno", "GE Number", 20)
            objUDFEngine.AddDateField("@SST_CONSHDR", "docdate", "Doc.Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("@SST_CONSHDR", "gedt", "GE.Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddNumericField("@SST_CONSHDR", "DocNum", "Doc.Number ", 10)
            objUDFEngine.AddAlphaField("@SST_CONSHDR", "InsNo", "Inspection Number", 20)
            objUDFEngine.AddAlphaField("@SST_CONSHDR", "Remarks", "Remarks", 250)
            objUDFEngine.addField("@SST_CONSHDR", "DocStat", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "D,O,C", "Draft,Open,Closed", "D")

            objUDFEngine.CreateTable("SST_CONSDTL", "Consolidation Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            objUDFEngine.AddAlphaField("@SST_CONSDTL", "itemcode", "Item code", 20)
            objUDFEngine.AddAlphaField("@SST_CONSDTL", "itemname", "Item Name", 250)
            objUDFEngine.AddAlphaField("@SST_CONSDTL", "reason", "Reason", 250)
            objUDFEngine.AddFloatField("@SST_CONSDTL", "recvdqty", "Received Quantity.Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_CONSDTL", "acptqty", "Accepted Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_CONSDTL", "rejctqty", "Rejected Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_CONSDTL", "rwkqty", "Rework Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_CONSDTL", "AccDev", "Accepted on Deviation", SAPbobsCOM.BoFldSubTypes.st_Quantity)

            objUDFEngine.CreateTable("SST_PRDCONHDR", "Prod.Cons Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("@SST_PRDCONHDR", "doccode", "Doc.Number", 20)
            objUDFEngine.AddAlphaField("@SST_PRDCONHDR", "prodno", "Production Number", 20)
            objUDFEngine.AddDateField("@SST_PRDCONHDR", "docdate", "Doc.Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddDateField("@SST_PRDCONHDR", "proddate", "Prod.Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.addField("@SST_PRDCONHDR", "DocStat", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "D,O,C", "Draft,Open,Closed", "D")

            objUDFEngine.CreateTable("SST_PRDCONDTL", "Prod.Cons Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            objUDFEngine.AddAlphaField("@SST_PRDCONDTL", "itemcode", "Item code", 40)
            objUDFEngine.AddAlphaField("@SST_PRDCONDTL", "itemname", "Item Name", 250)
            objUDFEngine.AddAlphaField("@SST_PRDCONDTL", "rcode", "Reason COde", 20)
            objUDFEngine.AddAlphaField("@SST_PRDCONDTL", "rsondesc", "Reason Name", 40)
            objUDFEngine.AddFloatField("@SST_PRDCONDTL", "recvdqty", "Received Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_PRDCONDTL", "acptqty", "Accepted Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_PRDCONDTL", "rejctqty", "Rejected Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_PRDCONDTL", "rwkqty", "Rework Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

            objUDFEngine.CreateTable("SST_PRDQCHDR", "Inspection Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "insno", "Inspec.Number", 20)
            objUDFEngine.AddDateField("@SST_PRDQCHDR", "insdate", "InspectionDate", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "prodno", "Production No.", 20)
            objUDFEngine.AddDateField("@SST_PRDQCHDR", "proddt", "Prod.Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "itemcode", "Item code", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "itemname", "Item Name", 200)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "uom", "Item", 5)
            objUDFEngine.AddFloatField("@SST_PRDQCHDR", "prdqty", "PROD Order Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "Stage", "Stage", 50)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "Type", "Type", 150)
            objUDFEngine.AddFloatField("@SST_PRDQCHDR", "Insqty", "Inspection Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddDateField("@SST_PRDQCHDR", "InsTime", "Inspection Time", SAPbobsCOM.BoFldSubTypes.st_Time)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "SrNo", "Serial No", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "rfd", "RFD No", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "bomrvno", "BOM Revision No", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "repsno", "Replacement Slno", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "ecn", "ECN", 20)
            objUDFEngine.AddDateField("@SST_PRDQCHDR", "AssTim", "Assembly Time", SAPbobsCOM.BoFldSubTypes.st_Time)
            objUDFEngine.AddDateField("@SST_PRDQCHDR", "rwTim", "Rework Time", SAPbobsCOM.BoFldSubTypes.st_Time)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "status", "Status", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "Rmrks", "Remarks", 200)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "Lot", "Lot Number", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "SubLvl", "Sub Level", 5)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "SItemCd", "Sub ItemCode", 20)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "SItmNme", "SItem Name", 200)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "PSlNo", "Parent SlNo", 50)
            objUDFEngine.AddAlphaField("@SST_PRDQCHDR", "SSlNo", "SubLevel SlNo", 50)
            objUDFEngine.addField("@SST_PRDQCHDR", "DocStat", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "D,O,C", "Draft,Open,Closed", "D")

            objUDFEngine.CreateTable("SST_PRDQCDTL", "Inspection Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "paradesc", "Parameter Desc", 40)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "smpsize", "Sample Size", 5)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "acclvl", "Accepted Level", 10)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "rejlvl", "Rejeted Level", 10)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "accqty", "Accepted Quantity", 10)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "rejqty", "Computed Quantity", 10)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "obser", "Observation", 200)
            objUDFEngine.AddAlphaField("@SST_PRDQCDTL", "rmks", "Remarks", 200)


            objUDFEngine.CreateTable("SST_OGAT", "Gate Entry", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("@SST_OGAT", "VCode", "Vendor Code", 20)
            objUDFEngine.AddAlphaField("@SST_OGAT", "VDesc", "Vendor Description", 100)
            objUDFEngine.AddAlphaField("@SST_OGAT", "DocNum", "DocNum", 8)
            objUDFEngine.addField("@SST_OGAT", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "O,C", "Open,Closed", "O")
            objUDFEngine.AddDateField("@SST_OGAT", "DocDate", "Doc Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@SST_OGAT", "PONum", "PO Number", 15)
            objUDFEngine.AddAlphaField("@SST_OGAT", "PODate", "PO Date", 30)
            objUDFEngine.AddAlphaField("@SST_OGAT", "WbNo", "Way Bill No", 50)
            objUDFEngine.AddAlphaField("@SST_OGAT", "Mode", "Mode of transport", 25)
            objUDFEngine.AddAlphaField("@SST_OGAT", "Recby", "Received by", 50)
            objUDFEngine.AddAlphaField("@SST_OGAT", "Vechno", "Vechicle No", 15)
            objUDFEngine.AddAlphaField("@SST_OGAT", "Rmrks", "Remarks", 250)
            objUDFEngine.AddAlphaField("@SST_OGAT", "RefNo", "Invoice Ref no", 25)
            objUDFEngine.AddAlphaField("@SST_OGAT", "PosStat", "Posting Status", 3)
            objUDFEngine.AddAlphaField("@SST_OGAT", "Docentry", "Docentry", 15)
            objUDFEngine.AddAlphaField("@SST_OGAT", "Series", "Series", 15)

            objUDFEngine.CreateTable("SST_GAT1", "Gate Entry-Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            objUDFEngine.AddAlphaField("@SST_GAT1", "ItmCode", "Item Code", 20)
            objUDFEngine.AddAlphaField("@SST_GAT1", "ItmDesc", "Way Bill No", 100)
            objUDFEngine.AddAlphaField("@SST_GAT1", "UOM", "UOM", 10)
            objUDFEngine.AddFloatField("@SST_GAT1", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.addField("@SST_GAT1", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "A,R", "Accepted,Rejected", "")
            objUDFEngine.AddAlphaField("@SST_GAT1", "GrnNo", "GRN No", 25)
            objUDFEngine.AddAlphaField("@SST_GAT1", "BrefNo", "BrefNo", 25)

            objUDFEngine.AddAlphaField("@SST_GAT1", "whs", "whs", 25)
            objUDFEngine.AddAlphaField("@SST_GAT1", "price", "price", 25)
            objUDFEngine.AddAlphaField("@SST_GAT1", "tcode", "tcode", 25)

            '*******Gate Entry Subcontractor****************
            objUDFEngine.CreateTable("SST_OGATSUB", "Gate Entry Sub Contractor", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "VCode", "Vendor Code", 20)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "VGroupCode", "Vendor Group Code", 50)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "VDesc", "Vendor Description", 100)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "DocNum", "DocNum", 8)
            objUDFEngine.addField("@SST_OGATSUB", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "O,C", "Open,Closed", "O")
            objUDFEngine.AddDateField("@SST_OGATSUB", "DocDate", "Doc Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "PONum", "PO Number", 15)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "PODate", "PO Date", 30)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "WbNo", "Way Bill No", 50)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "Mode", "Mode of transport", 25)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "Recby", "Received by", 50)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "Vechno", "Vechicle No", 15)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "Rmrks", "Remarks", 250)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "RefNo", "Invoice Ref no", 25)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "PosStat", "Posting Status", 3)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "Docentry", "Docentry", 15)
            objUDFEngine.AddAlphaField("@SST_OGATSUB", "Series", "Series", 15)

            objUDFEngine.CreateTable("SST_SUBGAT1", "Gate Entry Sub-Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "ItmCode", "Item Code", 20)
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "ItmDesc", "Way Bill No", 100)
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "UOM", "UOM", 10)
            objUDFEngine.AddFloatField("@SST_SUBGAT1", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.addField("@SST_SUBGAT1", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "A,R", "Accepted,Rejected", "")
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "GrnNo", "GRN No", 25)
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "BrefNo", "BrefNo", 25)
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "whs", "whs", 25)
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "price", "price", 25)
            objUDFEngine.AddAlphaField("@SST_SUBGAT1", "tcode", "tcode", 25)
            '***********************************************

            objUDFEngine.CreateTable("SST_SETUP", "Set Up", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_SETUP", "RGLWH", "Regular Warehouse", 10)
            objUDFEngine.AddAlphaField("@SST_SETUP", "RJTWH", "Reject Warehouse", 10)
            objUDFEngine.AddAlphaField("@SST_SETUP", "RWKWH", "Rework Warehouse", 10)


            objUDFEngine.CreateTable("SST_LOGIN", "Login", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            objUDFEngine.AddAlphaField("@SST_LOGIN", "UN", "UserName", 50)
            objUDFEngine.AddAlphaField("@SST_LOGIN", "PWD", "Password", 200)


            objUDFEngine.AddAlphaField("PDN1", "InsNo", "Inspection Number", 15)
            objUDFEngine.AddAlphaField("PDN1", "Res", "Reason", 15)
            objUDFEngine.AddAlphaField("PDN1", "EType", "Entry Type", 10)
            objUDFEngine.addField("OIGN", "RFISt", "RFI Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "A,R", "Accepted,Rejected", "")


            objUDFEngine.CreateTable("SST_RFIHDR", "RFI Inspection", SAPbobsCOM.BoUTBTableType.bott_Document)
            objUDFEngine.AddAlphaField("@SST_RFIHDR", "insno", "Inspec.Number", 20)
            objUDFEngine.AddDateField("@SST_RFIHDR", "insdate", "Inspec.Date", SAPbobsCOM.BoFldSubTypes.st_None)
            objUDFEngine.AddAlphaField("@SST_RFIHDR", "RFPNo", "RFP No.", 20)
            objUDFEngine.AddFloatField("@SST_RFIHDR", "RFPqty", "RFP Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddAlphaField("@SST_RFIHDR", "itemcode", "Item code", 20)
            objUDFEngine.AddAlphaField("@SST_RFIHDR", "itemname", "Item Name", 200)
            'objUDFEngine.AddAlphaField("@SST_RFIHDR", "uom", "Item UOM", 5)
            objUDFEngine.AddAlphaField("@SST_RFIHDR", "Status", "Status", 3)
            objUDFEngine.AddAlphaField("@SST_RFIHDR", "Remarks", "Remarks", 250)
            objUDFEngine.addField("@SST_RFIHDR", "DocStat", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "O,C", "Open,Closed", "O")

            objUDFEngine.CreateTable("SST_RFIDTL", "RFI Inspection-Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            objUDFEngine.AddAlphaField("@SST_RFIDTL", "CatCode", "Category Code", 20)
            objUDFEngine.AddAlphaField("@SST_RFIDTL", "pcode", "Parameter Code", 20)
            objUDFEngine.AddAlphaField("@SST_RFIDTL", "paradesc", "Parameter Desc", 250)
            objUDFEngine.AddAlphaField("@SST_RFIDTL", "uomcode", "Uom code", 20)
            objUDFEngine.AddAlphaField("@SST_RFIDTL", "uom", "Uom Name", 40)
            objUDFEngine.AddFloatField("@SST_RFIDTL", "usl", "USL", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_RFIDTL", "lsl", "LSL", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("@SST_RFIDTL", "value", "Actual Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddNumericField("@SST_RFIDTL", "smpsize", "Sample No ", 5)
            objUDFEngine.AddAlphaField("@SST_RFIDTL", "Status", "Status", 30)

        Catch ex As Exception

        End Try
    End Sub
    Private Sub CreateUDOS()
        Dim Ct(0) As String

        ''***********************************************MASTER UDO*********************************************
        SBO_Application.SetStatusBarMessage("Creating UDOs Please wait!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        createUDO("SST_QCREASON", "SST_REASON", "ReasonMaster", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        createUDO("SST_PRDCAT", "SST_PRDCAT", "Rejection Reason Category", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_PARACAT", "SST_PARACAT", "Parameter Category", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_QCPARAMETER", "SST_PARAM", "QC Param", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        createUDO("SST_QCUOM", "SST_UOM", "UOM Master", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("ss_cat", "ss_cat", "Category Master", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_SAMPLINGLEVEL", "SST_SMPLVL", "Sampling Level Master", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_ACCEPETLIMIT", "SST_ACCEPTLMT", "Acceptance Quality Limit", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_MODE", "SST_MOD", "Mode of Transport", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_STAGE", "SST_STG", "Stages", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_SETUP", "SST_SET", "Set UP", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        createUDO("SST_REF", "SST_REF", "Reference", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, True, False)
        createUDO("SST_LOGIN", "SST_LOGIN", "Login", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        ReDim Ct(1)
        Ct(0) = "SST_SLMDTL"
        createUDO("SST_SLMHDR", "SST_SLM", "Sampling Level Master", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Ct(0) = "SST_AQLDTL"
        createUDO("SST_AQLHDR", "SST_AQL", "Acceptance Quality Limit Master", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)

        Ct(0) = "SST_PLANDTL"
        createUDO("SST_PLANHDR", "SST_PLAN", "Sampling Plan-Prod", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Ct(0) = "SST_QCSTANDDTL"
        createUDO("SST_QCSTANDHDR", "SST_STAND", "Standard Parameter", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Ct(0) = "SST_PRDSTDDTL"
        createUDO("SST_PRDSTDHDR", "SST_PSTAN", "Prod.Std Parameter", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Ct(0) = "SST_NPLANDTL"
        createUDO("SST_NPLANHDR", "SST_NPLAN", "Sampling Plan-Inward", Ct, SAPbobsCOM.BoUDOObjType.boud_MasterData, False, True)
        Ct(0) = "SST_NQCDTL"
        createUDO("SST_NQCHDR", "SST_NINSP", "Inward Inspection", Ct, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
        Ct(0) = "SST_PRDCONDTL"
        createUDO("SST_PRDCONHDR", "SST_PCONS", "Prod.Consolidate", Ct, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
        Ct(0) = "SST_CONSDTL"
        createUDO("SST_CONSHDR", "SST_CONS", "Inward Consolidate", Ct, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
        Ct(0) = "SST_PRDQCDTL"
        createUDO("SST_PRDQCHDR", "SST_PRDINSP", "PRD.Inspection", Ct, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
        Ct(0) = "SST_GAT1"
        createUDO("SST_OGAT", "SST_GAT", "Gate Entry", Ct, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)

        Ct(0) = "SST_SUBGAT1"
        createUDO("SST_OGATSUB", "SST_GATSUB", "Gate Entry Sub Contractor", Ct, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)

        Ct(0) = "SST_RFIDTL"
        createUDO("SST_RFIHDR", "SST_RINSP", "RFI Inspection", Ct, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)

    End Sub
    Private Sub createUDO(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal childTable() As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        Dim i As Integer
        Dim lRetCode As Long


        oUserObjectMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If Not oUserObjectMD.GetByKey(udocode) Then
            oUserObjectMD.Code = udocode
            oUserObjectMD.Name = udoname
            oUserObjectMD.ObjectType = type
            oUserObjectMD.TableName = tblname
            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            If DfltForm = True Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO

                oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                oUserObjectMD.FormColumns.Add()
                oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                oUserObjectMD.FormColumns.Add()
            End If
            If FindForm = True Then
                If type = SAPbobsCOM.BoUDOObjType.boud_MasterData Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.FindColumns.ColumnAlias = "Code"
                    oUserObjectMD.FindColumns.ColumnDescription = "Code"
                    oUserObjectMD.FindColumns.Add()
                Else
                    Select Case udocode
                        Case "SST_GAT"
                            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                            oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                            oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                            oUserObjectMD.FindColumns.Add()
                        Case "SST_GATSUB"
                            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                            oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                            oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                            oUserObjectMD.FindColumns.Add()
                        Case "SST_NINSP"
                            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                            oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                            oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                            oUserObjectMD.FindColumns.Add()
                        Case "SST_CONS"
                            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                            oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                            oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                            oUserObjectMD.FindColumns.Add()
                        Case "SST_PCONS"
                            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                            oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                            oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                            oUserObjectMD.FindColumns.Add()
                        Case "SST_PRDINSP"
                            oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                            oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                            oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                            oUserObjectMD.FindColumns.Add()
                    End Select
                End If
            End If
            If childTable.Length > 0 Then
                For i = 0 To childTable.Length - 2
                    If Trim(childTable(i)) <> "" Then
                        oUserObjectMD.ChildTables.TableName = childTable(i)
                        oUserObjectMD.ChildTables.Add()
                    End If
                Next
            End If
            lRetCode = oUserObjectMD.Add()
            If lRetCode <> 0 Then
                ' MsgBox("error" + CStr(lRetCode))
                'MsgBox(objAddOn.objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = SBO_Application.Forms.AddEx(creationPackage)
            End If
        End If

    End Sub

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                SBO_Application.SetStatusBarMessage("A Shut Down Event has been caught" & _
                Environment.NewLine() & "Terminating Add On...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                SBO_Application.SetStatusBarMessage("A Company Change Event has been caught" & _
                Environment.NewLine() & "Terminating Add On...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                SBO_Application.SetStatusBarMessage("A Server Terminition Event has been caught" & _
                Environment.NewLine() & "Terminating Add On...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                System.Windows.Forms.Application.Exit()
        End Select
    End Sub


    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Dim FormType As String
        FormType = SBO_Application.Forms.ActiveForm.Type
        Try

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
    End Sub

#End Region
#Region ""
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        Try

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

#End Region

    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.RightClickEvent
        If eventInfo.BeforeAction Then

        End If
    End Sub
End Class
