using namespace System.Xml.Linq
[CmdletBinding()]
param
(
	#[Switch]$LoadTransactionFile,
	[Parameter(Mandatory = $true)]
	[string]$UserId,
	[Parameter(Mandatory = $true)]
	[string]$ApiKey
)

#Script Constants
$script:userId = $UserId
$script:apiKey = $ApiKey
$script:baseUri = "https://api.pocketsmith.com/v2/users/"

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
            [XElement]::new([XName]'Transactions')
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

        $Transaction.transaction_account.psobject.Properties | ForEach-Object {
            
            if($_.Value -eq $null)
            {
                $_.Value = ""
            }
        }

        $Transaction.transaction_account.institution.psobject.Properties | ForEach-Object {
           
            if($_.Value -eq $null)
            {
                $_.Value = ""
            }
        }

        $newElement = [XElement]::new(
            'Transaction',
            @(
                [XAttribute]::new('id', $Transaction.id),
                [XAttribute]::new('payee', $Transaction.payee),
                [XAttribute]::new('original_payee', $Transaction.original_payee),
                [XAttribute]::new('date', $Transaction.date),
                [XAttribute]::new('upload_source', $Transaction.upload_source),
                [XAttribute]::new('category', $Transaction.category),
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
                [XAttribute]::new('labels', $Transaction.labels.ToString()),
                [XAttribute]::new('created_at', $Transaction.created_at),
                [XAttribute]::new('updated_at', $Transaction.updated_at),

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

        $this.Save()
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

        $this.CurrentPage = $this.BaseUri + $script:userId + '/transactions?' + 'per_page=100'

    }

    [psobject]Run()
    {
        $responseHeaders = $null

        if($this.NextPage -ne $null)
        {
            $this.CurrentPage = $this.NextPage
        }
        
        #For Testing Only
        #$pageNumber = [Regex]::Match($this.CurrentPage, '(?<=page=)(\d+)(?=&)').Value
        #Write-Host "Processing Page $pageNumber"
        
        #Get Links
        $results = Invoke-RestMethod -Headers $this.Headers -Uri $this.CurrentPage -ResponseHeadersVariable responseHeaders
        $links = $responseHeaders["Link"][0]
    
        $this.FirstPage = [Regex]::Match($links, '(?<=<)[^<]+(?=>;\srel=\"first\")').Value
        $this.NextPage = [Regex]::Match($links, '(?<=<)[^<]+(?=>;\srel=\"next\")').Value
        $this.LastPage = [Regex]::Match($links, '(?<=<)[^<]+(?=>;\srel=\"last\")').Value

        return $results
    }

}

class TransactionDownloader
{
    Run()
    {
        Write-Host "Downloading Transactions..."
        $restClient = [RestClient]::new($script:baseUri)

        do {

            $transactions = $restClient.Run()

            $transactions | ForEach-Object {
              $transaction = $_
              
              $script:fileManager.AddTransaction($transaction)
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

$transactionDownloader = [TransactionDownloader]::new()
$transactionDownloader.Run()

Write-Host "All transactions have downloaded successfully. Exiting..."
        Pause

