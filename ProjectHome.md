# Budgeting for Mint and Wesabe using Excel #
This project includes an Excel workbook that you can use to track your spending. It loads your transaction data from [Mint](https://www.mint.com/) or [Wesabe](http://www.wesabe.com/) and shows your daily spending categorized using your Mint categories or Wesabe tags. You can enter monthly or weekly budget amounts for your tags and see how your spending compares to your budget.

## Getting Started ##
Choose the website that you want to work with. For Mint, download [mint-budgeting-v1\_3.xls](http://wesabe-excel-budget.googlecode.com/files/mint-budgeting-v1_3.xls); for Wesabe, download [wesabe-budgeting-v1\_2\_1.xls](http://wesabe-excel-budget.googlecode.com/files/wesabe-budgeting-v1_2_1.xls). If Excel prompts you to enable or disable macros, click the Enable Macros button. Newer versions of Excel will open the workbook in Protected View, click Enable Editing and Enable Content. On the Budget sheet, enter a few of your Mint categories or Wesabe tags in row 1, columns D though N. Enter monthly budgets for those tags in row 2.

When working with Mint, go to the [Transactions](https://wwws.mint.com/transaction.event) page in your web browser and use the Export all transactions link at the bottom of the page to save a copy of your Mint transactions.csv file. For Wesabe, the workbook will download transactions for you.

In the Excel workbook, press Ctrl+Shift+D to load your transactions. For Mint, you will be promted to choose a file; select the transactions.csv file you downloaded. Keep this transactions.csv file up to date by periodically downloading a new version from Mint. For Wesabe, you will be prompted for your username and password; enter your Wesabe credentials and click OK. The status bar at the bottom of Excel will say, "Connecting to Web...," then, "Retrieving data from web site...." When it says, "Ready," your transaction data has been downloaded.

View the other worksheets to see how your recent spending compares to your budget. After you tag more transactions in Wesabe go back to Excel and press Ctrl+Shift+D to download your latest data.

Don't worry about cells that contain `#REF!` errors, you haven't downloaded data for those dates yet.

For more information, visit the UsageTips or the ImplementationDetails pages. If you have problems, please file an [issue report](http://code.google.com/p/wesabe-excel-budget/issues/entry).

The ChangeLog lists all changes from previous versions.

## Security ##
When you use this workbook with Wesabe, you will be prompted for your Wesabe username and password. You can be confident that it is safe to enter your Wesabe password because it is going directly to Wesabe. The workbook contains an [Excel Web Query](http://office.microsoft.com/en-us/excel/HA010450851033.aspx) on the Transactions worksheet. It connects to https://www.wesabe.com/transactions.xml using Internet Explorer behind the scenes. The prompt for your password comes from the Excel Web Query, the workbook's macro code never sees your credentials.