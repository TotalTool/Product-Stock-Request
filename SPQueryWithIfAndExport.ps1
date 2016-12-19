# Setup the correct modules for SharePoint Manipulation 
if ( (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null ) 
{     
   Add-PsSnapin Microsoft.SharePoint.PowerShell 
} 

$web = Get-SPWeb http://sharepoint.totaltool.int/sales
$list = $web.Lists["Product Stock Request"]

$spQuery = New-Object Microsoft.SharePoint.SPQuery
$spQuery.ViewAttributes = "Scope='Recursive'";
$spQuery.RowLimit = 2000
$caml = '' 
$spQuery.Query = $caml 
$ExportCollection = @()
do
{
    $listItems = $list.GetItems($spQuery)
    $spQuery.ListItemCollectionPosition = $listItems.ListItemCollectionPosition
    foreach($item in $listItems){
        $Rep = $item["Author"]
        $PN1 = $item["_x0031__x002e__x0020_Product_x00"]
        $PN2 = $item["_x0032__x002e__x0020_Product_x00"]
        $PN3 = $item["_x0033__x002e__x0020_Product_x00"]
        $Status = $item["Status"]
        $StartDate = $item["Created"]
        if(($Status -eq "Approved") -and ($StartDate -lt (Get-Date).AddDays(-90))){
            #Write-host $StartDate, $Status
            $expobj = ""|select Rep, PN1, PN2, PN3, Status, StartDate
            $expobj.Rep = $item["Author"]
            $expobj.PN1 = $item["_x0031__x002e__x0020_Product_x00"]
            $expobj.PN2 = $item["_x0032__x002e__x0020_Product_x00"]
            $expobj.PN3 = $item["_x0033__x002e__x0020_Product_x00"]
            $expobj.Status = $item["Status"]
            $expobj.StartDate = $item["Created"]
            #Write-Host $expobj
            $ExportCollection += $expobj
            }
        
    }
}
while ($spQuery.ListItemCollectionPosition -ne $null)
$filePath = "H:\temp\PSR\PSR.csv"
$ExportCollection |Export-Csv -NoTypeInformation -Path $filePath
"Exported to: " + $filePath 