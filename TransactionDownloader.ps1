using namespace System.Xml.Linq
using namespace System.Collections.Generic
<#
	.SYNOPSIS
		Download transaction data from PocketSmith.
	
	.DESCRIPTION
		Downloads transaction and budget event data using the PocketSmith api at https://api.pocketsmith.com
	
	.PARAMETER UserId
		Pocketsmith account user Id. This is an integer.
	
	.PARAMETER ApiKey
		PocketSmith Api key.
	
	.PARAMETER StartDate
		Start date of data to return.
	
	.PARAMETER EndDate
		End date of data to return.
	
	.PARAMETER DataType
		The data type you wish to return. Default is "All".
	
	.PARAMETER LoadTransactionFile
		A description of the LoadTransactionFile parameter.
	
	.EXAMPLE
				PS C:\> .\TransactionDownloader.ps1 -UserId $value1 -ApiKey 'Value2'
	
	.NOTES
		Additional information about the file.
#>
[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true)]
	[int]$UserId,
	[Parameter(Mandatory = $true)]
	[string]$ApiKey,
	[string]$StartDate = '01/01/2015',
	[string]$EndDate = '12/31/2021',
	[ValidateSet('All', 'Transactions', 'BudgetEvents')]
	$DataType = 'All',
	[switch]$LoadTransactionFile
)

#Script Constants
$script:userId = $UserId
$script:apiKey = $ApiKey
$script:startDate = [DateTime]::Parse($StartDate).ToString("yyyy-MM-dd")
$script:endDate = [DateTime]::Parse($EndDate).ToString("yyyy-MM-dd")
$baseUri = "https://api.pocketsmith.com/v2/users/"

Add-Type -AssemblyName "System.Windows.Forms"


#Script Classes
class FileManager
{
    [System.Xml.Linq.XDocument]$DataFile
    [string]$DataFilePath

    FileManager()
    {
    
    }

    Load()
    {
        $consoleResponse = $null
        Write-Host "Select the file you wish to load."

        $fileBrowser = [System.Windows.Forms.OpenFileDialog]::new()
        $fileBrowser.Filter = "Xml Files (*.xml) | *.xml"
        [System.Windows.Forms.DialogResult]$fileDialogResult = $FileBrowser.ShowDialog()

        if ($fileDialogResult -eq [System.Windows.Forms.DialogResult]::OK) 
        {
            $this.DataFilePath = $fileBrowser.FileName
            try {
                $this.DataFile = [System.Xml.Linq.XDocument]::Load($this.DataFilePath)
            }
            catch [Exception]
            {
                do {
                    Write-Host "File could not be loaded."
                    $consoleResponse = Read-Host "Would you like to create a new file? [Y]es or [N]o"
    
                } while (($consoleResponse -ne "Y") -and ($consoleResponse -ne "N"))

                if($consoleResponse -eq "Y")
                {
                    $this.Create()
                }
                if($consoleResponse -eq "N")
                {
                    Read-Host "Transaction file will not be created. Press ENTER to close"
                    exit
                }
            }
        }
    }

    Create()
    {
        Write-Host "Creating transaction file..."

        $fileBrowser = [System.Windows.Forms.SaveFileDialog]::new()
        $fileBrowser.Filter = "Xml File (*.xml) | *.xml"
        $fileBrowser.AddExtension = ".xml"
        $fileBrowser.ShowDialog()
        $this.DataFilePath = $fileBrowser.FileName

        $this.DataFile = [XDocument]::new(
            [XDeclaration]::new('1.0', 'utf-8', 'yes'),
            [XElement]::new(
                [XName]'Data',
            [XElement]::new([XName]'Transactions'),
            [XElement]::new([XName]'BudgetEvents')
        ))

        $this.DataFile.Save($this.DataFilePath)
    }

    Save()
    {
        $this.DataFile.Save($this.DataFilePath)
    }

    AddTransaction([psobject]$Transaction)
    {
        #Replace null values with ""
        $Transaction.psobject.Properties | ForEach-Object{
            
            if(($_.Value -eq $null))
            {
                $_.Value = ""
            }
        }
        #Replace null values in transaction_account with ""
        $Transaction.transaction_account.psobject.Properties | ForEach-Object {
            
            if($_.Value -eq $null)
            {
                $_.Value = ""
            }
        }
        #Replace null values in institution with ""
        $Transaction.transaction_account.institution.psobject.Properties | ForEach-Object {
           
            if($_.Value -eq $null)
            {
                $_.Value = ""
            }
        }

        #Handle null labels

        if($Transaction.labels[0].Length -gt 0)
        {
            $transactionLabels = [string]::Join(',', $Transaction.labels)
        }
        else {
            $transactionLabels = ""
        }
        
        #Handle null transaction categories.
        if($Transaction.category -eq "")
        {
            $transactionCategory = [XElement]::new('category', "")
        }
        else{

            $Transaction.category.psobject.Properties | ForEach-Object {
           
                if($_.Value -eq $null)
                {
                    $_.Value = ""
                }
            }

            $transactionCategory = [XElement]::new('category',
            @(
                [XAttribute]::new('id', $Transaction.category.id),
                [XAttribute]::new('title', $Transaction.category.title),
                [XAttribute]::new('colour', $Transaction.category.colour),
                [XAttribute]::new('is_transfer', $Transaction.category.is_transfer),
                [XAttribute]::new('is_bill', $Transaction.category.is_bill),
                [XAttribute]::new('refund_behaviour', $Transaction.category.refund_behaviour),
                [XAttribute]::new('parent_id', $Transaction.category.parent_id),
                [XAttribute]::new('roll_up', $Transaction.category.roll_up),
                [XAttribute]::new('created_at', $Transaction.category.created_at),
                [XAttribute]::new('updated_at', $Transaction.category.updated_at)
                #Child categories not included. 
            ))
        }

        $newElement = [XElement]::new(
            'Transaction',
            @(
                [XAttribute]::new('id', $Transaction.id),
                [XAttribute]::new('payee', $Transaction.payee),
                [XAttribute]::new('original_payee', $Transaction.original_payee),
                [XAttribute]::new('date', $Transaction.date),
                [XAttribute]::new('upload_source', $Transaction.upload_source),
                [XAttribute]::new('closing_balance', $Transaction.closing_balance),
                [XAttribute]::new('cheque_number', 
                [Regex]::Match($Transaction.cheque_number, '\d+')),
                [XAttribute]::new('memo', $Transaction.memo),
                [XAttribute]::new('amount', $Transaction.amount),
                [XAttribute]::new('amount_in_base_currency', $Transaction.amount_in_base_currency),
                [XAttribute]::new('type', $Transaction.type),
                [XAttribute]::new('is_transfer', $Transaction.is_transfer),
                [XAttribute]::new('needs_review', $Transaction.needs_review),
                [XAttribute]::new('status', $Transaction.status),
                [XAttribute]::new('note', $Transaction.note),
                [XAttribute]::new('labels', $transactionLabels),
                [XAttribute]::new('created_at', $Transaction.created_at),
                [XAttribute]::new('updated_at', $Transaction.updated_at),

                #Insert previously created category.
                $transactionCategory,

                [XElement]::new('transaction_account', 
                @(
                    [XAttribute]::new('id', $Transaction.transaction_account.id),
                    [XAttribute]::new('name', $Transaction.transaction_account.name),
                    [XAttribute]::new('number', $Transaction.transaction_account.number),
                    [XAttribute]::new('type', $Transaction.transaction_account.type),
                    [XAttribute]::new('currency_code', $Transaction.transaction_account.currency_code),
                    [XAttribute]::new('current_balance', $Transaction.transaction_account.current_balance),
                    [XAttribute]::new('current_balance_in_base_currency', $Transaction.transaction_account.current_balance_in_base_currency),
                    [XAttribute]::new('current_balance_exchange_rate', $Transaction.transaction_account.current_balance_exchange_rate),
                    [XAttribute]::new('current_balance_date', $Transaction.transaction_account.current_balance_date),
                    [XAttribute]::new('safe_balance', $Transaction.transaction_account.safe_balance),
                    [XAttribute]::new('safe_balance_in_base_currency', $Transaction.transaction_account.safe_balance_in_base_currency),
                    [XAttribute]::new('starting_balance', $Transaction.transaction_account.starting_balance),
                    [XAttribute]::new('starting_balance_date', $Transaction.transaction_account.starting_balance_date),
                    [XAttribute]::new('created_at', $Transaction.transaction_account.created_at),
                    [XAttribute]::new('updated_at', $Transaction.transaction_account.updated_at),

                    [XElement]::new('institution',
                    @(
                        [XAttribute]::new('id', $Transaction.transaction_account.institution.Id),
                        [XAttribute]::new('title', $Transaction.transaction_account.institution.title),
                        [XAttribute]::new('currency_code', $Transaction.transaction_account.institution.currency_code),
                        [XAttribute]::new('created_at', $Transaction.transaction_account.institution.created_at),
                        [XAttribute]::new('updated_at', $Transaction.transaction_account.institution.updated_at)
                    ))
                ))
            ))

            $this.DataFile.Element("Data").Element("Transactions").Add($newElement)

    
    }

    AddBudgetEvent([psobject]$budgetEvent)
    {
        #Replace null values with ""
        $budgetEvent.psobject.Properties | ForEach-Object{
            
            if(($_.Value -eq $null))
            {
                $_.Value = ""
            }
        }

        #Replace null values in category with ""
        $budgetEvent.category.psobject.Properties | ForEach-Object{

            if ($_.Value -eq $null)
            {
                $_.Value = ""
            }
        }

        #Replace null values in scenario with ""
        $budgetEvent.scenario.psobject.Properties | ForEach-Object{
            
            if($_.Value -eq $null)
            {
                $_.Value = ""
            }
        }

        $newElement = [XElement]::new('BudgetEvent',
        @(
            [XAttribute]::new('id', $budgetEvent.id),
            [XAttribute]::new('amount', $budgetEvent.amount),
            [XAttribute]::new('amount_in_base_currency', $budgetEvent.amount_in_base_currency),
            [XAttribute]::new('currency_code', $budgetEvent.currency_code),
            [XAttribute]::new('date', $budgetEvent.date),
            [XAttribute]::new('colour', $budgetEvent.colour),
            [XAttribute]::new('note', $budgetEvent.note),
            [XAttribute]::new('repeat_type', $budgetEvent.repeat_type),
            [XAttribute]::new('repeat_interval', $budgetEvent.repeat_interval),
            [XAttribute]::new('series_id', $budgetEvent.series_id),
            [XAttribute]::new('series_start_id', $budgetEvent.series_start_id),
            [XAttribute]::new('infinite_series', $budgetEvent.infinite_series),

                [XElement]::new('category',
                @(
                    [XAttribute]::new('id', $budgetEvent.category.id),
                    [XAttribute]::new('title', $budgetEvent.category.title),
                    [XAttribute]::new('colour', $budgetEvent.category.colour),
                    [XAttribute]::new('is_transfer', $budgetEvent.category.is_transfer),
                    [XAttribute]::new('is_bill', $budgetEvent.category.is_bill),
                    [XAttribute]::new('refund_behaviour', $budgetEvent.category.refund_behaviour),
                    [XAttribute]::new('parent_id', $budgetEvent.category.parent_id),
                    [XAttribute]::new('roll_up', $budgetEvent.category.roll_up),
                    [XAttribute]::new('created_at', $budgetEvent.category.created_at),
                    [XAttribute]::new('updated_at', $budgetEvent.category.updated_at)
                )),
                [XElement]::new('scenario',
                @(
                    [XAttribute]::new('id', $budgetEvent.scenario.id),
                    [XAttribute]::new('title', $budgetEvent.scenario.title),
                    [XAttribute]::new('description', $budgetEvent.scenario.description),
                    [XAttribute]::new('interest_rate', $budgetEvent.scenario.interest_rate),
                    [XAttribute]::new('interest_rate_repeat_id', $budgetEvent.scenario.interest_rate_repeat_id),
                    [XAttribute]::new('type', $budgetEvent.scenario.type),
                    [XAttribute]::new('minimum_value', $budgetEvent.scenario.minimum_value),
                    [XAttribute]::new('maximum_value', $budgetEvent.scenario.maximum_value),
                    [XAttribute]::new('achieve_date', $budgetEvent.scenario.achieve_date),
                    [XAttribute]::new('starting_balance', $budgetEvent.scenario.starting_balance),
                    [XAttribute]::new('starting_balance_date', $budgetEvent.scenario.starting_balance_date),
                    [XAttribute]::new('closing_balance', $budgetEvent.scenario.closing_balance),
                    [XAttribute]::new('closing_balance_date', $budgetEvent.scenario.closing_balance_date),
                    [XAttribute]::new('current_balance', $budgetEvent.scenario.current_balance),
                    [XAttribute]::new('current_balance_in_base_currency', $budgetEvent.scenario.current_balance_in_base_currency),
                    [XAttribute]::new('current_balance_exchange_rate', $budgetEvent.scenario.current_balance_exchange_rate),
                    [XAttribute]::new('current_balance_date', $budgetEvent.scenario.current_balance_date),
                    [XAttribute]::new('safe_balance', $budgetEvent.scenario.safe_balance),
                    [XAttribute]::new('safe_balance_in_base_currency', $budgetEvent.scenario.safe_balance_in_base_currency),
                    [XAttribute]::new('created_at', $budgetEvent.scenario.created_at),
                    [XAttribute]::new('updated_at', $budgetEvent.scenario.updated_at)
                ))
        ))

        $this.DataFile.Element("Data").Element("BudgetEvents").Add($newElement)
    }
}

class RestClient
{
    [string]$BaseUri
    [string]$FirstPage
    [string]$LastPage
    [string]$NextPage
    [string]$CurrentPage
    $Headers
    
    RestClient([string]$baseUri)
    {
        $this.BaseUri = $baseUri

        $this.Headers = @{
            "X-Developer-Key" = $script:apiKey
            "accept"          = "application/json"
        }

        $this.CurrentPage = $this.BaseUri

    }

    [psobject]Run([string]$uri)
    {
        $responseHeaders = $null
        
        #Get Links
        $results = Invoke-RestMethod -Headers $this.Headers -Uri $uri -ResponseHeadersVariable responseHeaders
        $links = $responseHeaders["Link"][0]
    
        $this.FirstPage = [Regex]::Match($links, '(?<=<)[^<]+(?=>;\srel=\"first\")').Value
        $this.NextPage = [Regex]::Match($links, '(?<=<)[^<]+(?=>;\srel=\"next\")').Value
        $this.LastPage = [Regex]::Match($links, '(?<=<)[^<]+(?=>;\srel=\"last\")').Value

        return $results
    }

}

class TransactionDownloader
{
    [string]$BaseUri
    [int]$TotalTransactions

    TransactionDownloader([string]$BaseUri)
    {
        $this.BaseUri = $BaseUri + $script:userId + '/transactions?' + 'per_page=100'
    }

    Run()
    {
        Write-Host "Downloading Transactions..."
        $restClient = [RestClient]::new($this.BaseUri)

        do {
            
                if($restClient.NextPage -ne $null)
                {
                    $restClient.CurrentPage = $restClient.NextPage
                }
            

                #For Testing
                $page = [Regex]::Match($restClient.CurrentPage, '(?<=page=)(\d+)(?=&)').Value
                Clear-Host
                Write-Host "PocketSmith Transaction Downloader"
                Write-Host "Processing Transaction Page $page"

                $transactions = $restClient.Run($restClient.CurrentPage)
                $transactions | ForEach-Object {
                    $transaction = $_            
                    
                    $script:fileManager.AddTransaction($transaction)
                
            }
            

        } until ($restClient.CurrentPage -eq $restClient.LastPage)

        $script:fileManager.Save()
    }
}

class BudgetDownloader
{
    [string]$BaseUri
    [int]$TotalBudgets

    BudgetDownloader([string]$BaseUri)
    {
        $this.BaseUri = $BaseUri + $script:userId + '/events?' + 'per_page=100' + "&start_date=$script:startDate" + "&end_date=$script:endDate"
    }

    Run()
    { 
        Write-Host "Downloading Budget Events..."
        $restClient = [RestClient]::new($this.BaseUri)

        do {
            if($restClient.NextPage -ne $null)
            {
                $restClient.CurrentPage = $restClient.NextPage
            }
        

            #For Testing
            $page = [Regex]::Match($restClient.CurrentPage, '(?<=page=)(\d+)(?=&)').Value
            Clear-Host
            Write-Host "PocketSmith Transaction Downloader"
            Write-Host "Processing Budget Event Page $page"

            $events = $restClient.Run($restClient.CurrentPage)

            $events | ForEach-Object {
                $budgetEvent = $_         
                
                $script:fileManager.AddBudgetEvent($budgetEvent)
            
            }
        } until ($restClient.CurrentPage -eq $restClient.LastPage)

        $script:fileManager.Save()
    }
}

# Script Body
Clear-Host

$script:fileManager = [FileManager]::new()
if($LoadTransactionFile)
{
    $fileManager.Load()    
}
else
{
   $fileManager.Create() 
}

switch ($DataType) {
    'Transactions' { 
        $transactionDownloader = [TransactionDownloader]::new($baseUri)
        $transactionDownloader.Run()
    }
    'BudgetEvents' {
        $budgetDownloader = [BudgetDownloader]::new($baseUri)
        $budgetDownloader.Run()
    }
    'All' {
        $transactionDownloader = [TransactionDownloader]::new($baseUri)
        $transactionDownloader.Run()

        $budgetDownloader = [BudgetDownloader]::new($baseUri)
        $budgetDownloader.Run()
    }
}


Write-Host "All downloads have completed successfully. Exiting..."
        Pause

