# Point this script at an etsy shop and it will attempt to catalog the sales from the shop and create an output that can be read
# Andy
# 11-27-22
# 1.0.1

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# setup variables/"constants"

$outputDir = "C:\temp\EtsyScriptOutput\"

$etsyPaginationEnding = "?ref=pagination&page="

$browserHeaderInfo = "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:99.0) Gecko/20100101 Firefox/99.0"

$organizedPowershellObject  = @()

[double]$runningShopProfit = 0

###################

Clear-Variable collatedSalesInfoFull,matches,etsySellerRawRequestResponse,etsySellerContentPageRegex,etsySellerMaxPagination,sortedObjectOfEtsySales,listOfShops -ErrorAction SilentlyContinue

function collectShopNames
    {
        Write-Host "Enter in shop names or URL's to shops separated by commas:"
        $input = Read-Host
        $listOfEntries = $input.split(",")
        foreach ($entry in $listOfEntries)
            {
                if ($entry -like "*.com*")
                    {
                        [array]$listOfShops += $entry -match 'etsy\.com\/shop\/(.*?)\?' | % {$matches[1]}
                    }
                Else
                    {
                        [array]$listOfShops += $entry
                    }

            }
            Clear-Variable matches -ErrorAction SilentlyContinue
            return($listOfShops)
    }

Foreach ($shopName in collectshopnames)
{
    Clear-Variable salesinfoexposed -ErrorAction SilentlyContinue

Try # Did the etsy seller even expose sales info?
    {
        $etsySellerRawRequestResponse = iwr -uri "https://www.etsy.com/shop/$($shopName)/sold" -UserAgent "$browserHeaderInfo"
    }
Catch
    {
        Write-Host "Seller $shopName doesn't share sales info"
        $salesInfoExposed = $false
    }

if ($salesInfoExposed -ne $false) # if sales info is exposed continue, otherwise skip this loop
    {
        $etsySellerContentPageRegex = $etsySellerRawRequestResponse.content | Select-String -pattern 'class\=\"page\-(\d+)' -AllMatches
        [int]$etsySellerMaxPagination = ($etsySellerContentPageRegex.Matches | %{$_.groups[1].value})[-1]


        # Start collecting sales info
        $collatedSalesInfoFull += $etsySellerRawRequestResponse.links | where {$_.href -like "*listing*"} #first page with no pagination

        for ($i = 2; $i -le $etsySellerMaxPagination; $i++)
            {
                Clear-Variable tempEtsySellerPage -ErrorAction SilentlyContinue
                Write-Progress -Activity "Gathering Data" -percentcomplete $(($i / $etsySellerMaxPagination) * 100)
                do
                {
                write-host "Loop $i"
                $thisLoopDone = $True
                Try {
                        $tempEtsySellerPage = Invoke-Webrequest -uri "https://www.etsy.com/shop/$($shopName)/sold$($etsyPaginationEnding)$($i)" -UserAgent "$browserHeaderInfo" -Verbose
                    }
                Catch 
                    {
                        Write-Host "Error on page read: $($_.Exception.Response.StatusCode.Value__) - Waiting 61 seconds..."
                        $thisLoopDone = $False
                        sleep -Seconds 61 # I think this is correct limit for etsy, seems to work after two 30 sec segments...
                    }
                } while($thisLoopDone -eq $False)
        
                $collatedSalesInfoFull += $tempEtsySellerPage.links | where {$_.href -like "*listing*"} # next page of pagination sale data
                Sleep -seconds 3 # slow down by a touch
            }

        $sortedObjectOfEtsySales = ($collatedSalesInfoFull | Select-Object 'data-listing-id',Title | Group-Object -Property 'data-listing-id'| Where {$_.count -gt 2} | Sort-Object -Property Count -Descending)



        Foreach ($itemFound in $sortedObjectOfEtsySales)
            {
                Clear-Variable tempobject -ErrorAction SilentlyContinue
                $tempObject = New-Object -TypeName psobject

                Do
                    {
                    $thisInvokeWorked = $True
                    Try {
                            $tempItemPage = Invoke-Webrequest -uri "https://www.etsy.com/listing/$($itemFound.Name)" -UserAgent "$browserHeaderInfo" -Verbose
                        }
                    Catch 
                        {
                            Write-Host "Error on item page read: $($_.Exception.Response.StatusCode.Value__) - Waiting 61 seconds..."
                            $thisInvokeWorked = $False
                            sleep -Seconds 61 # I think this is correct limit for etsy, seems to work after two 30 sec segments...
                        }
                    } while($thisInvokeWorked -eq $False)
                Try
                    {
                        $priceString = $($tempItemPage.content -match '\$(\d+[\.\d+]+)[\s\S].*\<\/p\>' | % {$matches[1]})
                    }
                Catch
                    {
                        $priceString = "Not Available"
                    }
                Add-Member -InputObject $tempObject -MemberType NoteProperty -Name NumberSold -Value $itemFound.Count
                Add-Member -InputObject $tempObject -MemberType NoteProperty -Name ID -Value $itemFound.Name
                Add-Member -InputObject $tempObject -MemberType NoteProperty -Name Price -Value $priceString
                Add-Member -InputObject $tempObject -MemberType NoteProperty -Name Title -Value $itemFound.group.title[0]
                Add-Member -InputObject $tempObject -MemberType NoteProperty -Name itemURL -Value "https://www.etsy.com/listing/$($itemFound.Name)"
                Add-Member -InputObject $tempObject -MemberType NoteProperty -Name Seller -Value "$($shopName)"
                Add-Member -InputObject $tempObject -MemberType NoteProperty -Name SellerURL -Value "https://www.etsy.com/shop/$($shopName)"
                $organizedPowershellObject += $tempObject
                sleep -Seconds 1
            }
    }
}

New-Item -Path $($outputDir) -ItemType Directory -Force

# Add up total profit based on items multiplied by number sold for each

Foreach ($uniqueItemOnShop in $organizedPowershellObject)
    {
        if ($uniqueItemOnShop.Price -ne "Not Available") { [double]$runningShopProfit += [double]$($uniqueItemOnShop.Price)*$($uniqueItemOnShop.NumberSold) }
    }

$lastProfitObject = New-Object -TypeName psobject
Add-Member -InputObject $lastProfitObject -MemberType NoteProperty -Name EstimatedShopProfits -Value "$runningShopProfit"
$organizedPowershellObject += $lastProfitObject

# Output to file
$organizedPowershellObject | Export-Clixml "$($outputDir)$(Get-Date -UFormat %m-%d-%y-T%H-%M-%S).xml"