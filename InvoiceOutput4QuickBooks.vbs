'********************************************************************************************
' Custom function: DoTimecard (exports data to Excel and print previews)
'********************************************************************************************
    'This procedure fills an Excel spreadsheet with data from SQL Server Database
    'to create an output file to import into QuiclBooks 
    	Dim ExcelApp, xlSheet, row, col



    'Create an Excel object. This will load a new instance of Excel even if Excel is already running.

    	Set ExcelApp = CreateObject("Excel.Application")
    	ExcelApp.Visible = True  'Make Excel visible through the Application object
    	ExcelApp.Workbooks.Add
    	Set xlSheet = ExcelApp.Workbooks(1).Worksheets("Sheet1")
    	xlSheet.Name = "Bill"

    'Start filling column headers


	ExcelApp.cells(1,1).formula = "Bill number"
	ExcelApp.cells(1,2).formula = "Line.Amount"
	ExcelApp.cells(1,3).formula = "Vendor"
	ExcelApp.cells(1,4).formula = "AP Account"
	ExcelApp.cells(1,5).formula = "Currency"
	ExcelApp.cells(1,6).formula = "TxnDate"
	ExcelApp.cells(1,7).formula = "Due date"
	ExcelApp.cells(1,8).formula = "Exchange Rate"
	ExcelApp.cells(1,9).formula = "Global tax calculation"
	ExcelApp.cells(1,10).formula = "Update Bill"
	ExcelApp.cells(1,11).formula = "Line: Account name"
	ExcelApp.cells(1,12).formula = "Line: Billable status"
	ExcelApp.cells(1,13).formula = "Line: Class"
	ExcelApp.cells(1,14).formula = "Line: Customer"
	ExcelApp.cells(1,15).formula = "Line: Description"
	ExcelApp.cells(1,16).formula = "Line: Product name"
	ExcelApp.cells(1,17).formula = "Item: Quantity"
	ExcelApp.cells(1,19).formula = "Line: Tax code"
	ExcelApp.cells(1,20).formula = "Item: Unit price"
	ExcelApp.cells(1,21).formula = "Line: Markup percent"
	ExcelApp.cells(1,22).formula = "Memo"
	ExcelApp.cells(1,23).formula = "Sales term"
	ExcelApp.cells(1,24).formula = "Txn date"
	ExcelApp.cells(1,25).formula = "Total tax"
	ExcelApp.cells(1,26).formula = "Transaction Location Type"
	ExcelApp.cells(1,27).formula = "Attachment: File Link"
	ExcelApp.cells(1,28).formula = "Attachment: Id"
	ExcelApp.cells(1,29).formula = "Attachment: Attach to email"


    'Get database record set

	set rs1 = CreateObject("ADODB.Recordset")
	rs1.ActiveConnection = "Driver={SQL Server};Server=DataServer\SQL2017;database=intranet;uid=webuser;pwd=info4web"
	rs1.Source = "SELECT IA.ApprovalID,IA.ApprovalGID,IA.VendorName,IA.InvoiceDueDate,IAI.AccountNo,IAI.ItemDesc,IAI.Qty,IAI.TotalCost,IA.Hold,IA.Notes,IAI.AccountDesc,IAF.ImageNo,IAF.ImageName "
	rs1.Source = rs1.Source &"FROM InvoiceApprovalItems as IAI INNER JOIN "
	rs1.Source = rs1.Source &"InvoiceApprovalFiles as IAF ON IAI.ApprovalGID = IAF.RecNo "
	rs1.Source = rs1.Source &"INNER JOIN InvoiceApproval as IA ON IAI.ApprovalGID = IA.ApprovalGID WHERE IA.Company = 'EVENTMGMT' AND (IA.Exported = '' or IA.Exported is NULL)"
	rs1.CursorType = 0
	rs1.CursorLocation = 2
	rs1.LockType = 3
	rs1.Open


row = 1

    'Loop through record set filling cells

Do while not rs1.eof

    'Grab Items for each Invoice Approval

	ApprovalID = rs1("ApprovalID")
	ApprovalGID = rs1("ApprovalGID")

	set rs2 = CreateObject("ADODB.Recordset")
	rs2.ActiveConnection = "Driver={SQL Server};Server=DataServer\SQL2017;database=intranet;uid=webuser;pwd=info4web"
	rs2.Source = "SELECT * FROM InvoiceApprovalItems WHERE ApprovalGID = 'ApprovalGID' ORDER BY ItemID"
	rs2.CursorType = 0
	rs2.CursorLocation = 2
	rs2.LockType = 3
	rs2.Open

    'Set Variables

	row = row + 1


	ApprovalGID = rs1("ApprovalGID")
	BillNumber = rs1("ApprovalGID")
	LineAmount = rs1("TotalCost")
	Vendor = rs1("VendorName")
	APAccount = ""
	Currency1 = ""
	TxnDate = ""
	DueDate  = rs1("InvoiceDueDate")
	ExchangeRate = ""
	GlobalTaxCalculation = ""
	UpdateBill = ""
	LineAccountName  = rs1("AccountDesc")
	LineBillableStatus  = rs1("Hold")
	LineClass = ""
	LineCustomer  = rs1("AccountNo")
	LineDescription  = rs1("ItemDesc")
	LineProductName = ""
	ItemQuantity  = rs1("Qty")
	LineTaxCode = ""
	ItemUnitPrice = ""
	LineMarkupPercent = ""
	Memo  = rs1("Notes")
	SalesTerm = ""
	TxnDate = ""
	TotalTax = ""
	TransactionLocationType = ""
	Location = ""
	AttachmentFileLink = "https://intranet.wisdells.com/invoiceapproval/ViewImage.asp?ImageNo="&rs1("ImageNo")&"&ImageName="&rs1("ImageName")&""
	AttachmentId = ""
	AttachmentAttach2Email = ""




    'Start filling cells
    
	ExcelApp.cells(row,1).formula = BillNumber
	ExcelApp.cells(row,2).formula = LineAmount
	ExcelApp.cells(row,3).formula = Vendor
	ExcelApp.cells(row,4).formula = APAccount
	ExcelApp.cells(row,5).formula = Currency1
	ExcelApp.cells(row,6).formula = TxnDate
	ExcelApp.cells(row,7).formula = DueDate 
	ExcelApp.cells(row,8).formula = ExchangeRate
	ExcelApp.cells(row,9).formula = GlobalTaxCalculation
	ExcelApp.cells(row,10).formula = UpdateBill
	ExcelApp.cells(row,11).formula = LineAccountName 
	ExcelApp.cells(row,12).formula = LineBillableStatus 
	ExcelApp.cells(row,13).formula = LineClass
	ExcelApp.cells(row,14).formula = LineCustomer 
	ExcelApp.cells(row,15).formula = LineDescription 
	ExcelApp.cells(row,16).formula = LineProductName
	ExcelApp.cells(row,17).formula = ItemQuantity 
	ExcelApp.cells(row,18).formula = LineTaxCode
	ExcelApp.cells(row,19).formula = ItemUnitPrice
	ExcelApp.cells(row,20).formula = LineMarkupPercent
	ExcelApp.cells(row,21).formula = Memo 
	ExcelApp.cells(row,22).formula = SalesTerm
	ExcelApp.cells(row,23).formula = TxnDate
	ExcelApp.cells(row,24).formula = TotalTax
	ExcelApp.cells(row,25).formula = TransactionLocationType
	ExcelApp.cells(row,26).formula = Location
	ExcelApp.cells(row,27).formula = AttachmentFileLink
	ExcelApp.cells(row,28).formula = AttachmentId
	ExcelApp.cells(row,29).formula = AttachmentAttach2Email



    'Update exported field
    
	set rs = CreateObject("ADODB.Recordset")
	rs.ActiveConnection = "Driver={SQL Server};Server=DataServer\SQL2017;database=intranet;uid=webuser;pwd=info4web"
	rs.Source = "UPDATE InvoiceApproval SET Exported = 'Y' WHERE ApprovalGID = '"&ApprovalGID&"'"
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType = 3
	rs.Open
	msgbox rs.Source

rs1.movenext
Loop
	Set ExcelApp = Nothing
	rs1.Close
 