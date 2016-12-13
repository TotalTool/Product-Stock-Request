$web = Get-SPWeb http://sharepoint.totaltool.int/sales
$list = $web.Lists["Product Stock Request"]

$spQuery = New-Object Microsoft.SharePoint.SPQuery
$spQuery.ViewAttributes = "Scope='Recursive'";
$spQuery.RowLimit = 2000
$caml = '' 
$spQuery.Query = $caml 

do
{
    $listItems = $list.GetItems($spQuery)
    $spQuery.ListItemCollectionPosition = $listItems.ListItemCollectionPosition
    foreach($item in $listItems)
    {
        $Rep = $item["Author"]
        $PN1 = $item["_x0031__x002e__x0020_Product_x00"]
        $PN2 = $item["_x0032__x002e__x0020_Product_x00"]
        $PN3 = $item["_x0033__x002e__x0020_Product_x00"]
        $Status = $item["Status"]
        $StartDate = $item["Created"]
        Write-host $status
    }
}
while ($spQuery.ListItemCollectionPosition -ne $null)