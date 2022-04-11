$url = "https://docs.microsoft.com/en-us/windows/release-health/release-information"

if ( -Not $wr_response ) {
    $wr = [System.Net.WebRequest]::Create($url)
    $wr_response = $wr.GetResponse()
    $wr_html = (New-Object System.IO.StreamReader ($wr_response.GetResponseStream())).ReadToEnd()
}

Write-Host -NoNewline "Get Tables... "
$tables = $wr_html | Select-String -Pattern '<h(.|\n)*?<\/table>' -AllMatches | % { $_.Matches } | % { $_.Value }
Write-Host $tables.Count
#$tables | Out-GridView -OutputMode Multiple

#

$sac = @()

if ( $tables.Count -ge 1 ) {
    
    $table = $tables[0]
    #$table | Out-GridView -OutputMode Multiple

    $title = $table | Select-String -Pattern '(?<=>)(.*?)(?=</h4>)' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    $trs = $table | Select-String -Pattern '<tr(.|\n)*?<\/tr>' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Skip 1
    #$trs | Out-GridView -OutputMode Multiple

    $trs | % {
        
        $tr = $_
        $tds = $tr | Select-String -Pattern '(?<=>)(.*?)(?=</td>)' -AllMatches  | % { $_.Matches } | % { $_.Value }
        #$tds | Out-GridView -OutputMode Multiple

        if ( $tds.Count -ge 7 ) {
            
            $version = $tds[0].Replace("<td>","")
            $servicingoption = $tds[1]            
            
            $availabilitydate = $null
            try { $availabilitydate = ([datetime]$tds[2]).ToShortDateString() } catch {}
                       
            $latestrevisiondate = $null
            try { $latestrevisiondate = ([datetime]$tds[3]).ToShortDateString() } catch {}

            $osbuild = $tds[4]

            $supportend_home = $null
            try { $supportend_home = ([datetime]$tds[5]).ToShortDateString() } catch {}
            if ( $tds[5].Trim() -like "End*" ) { $supportend_home = "EndOfService" }

            $supportend_enterprise = $null
            try { $supportend_enterprise = ([datetime]$tds[6]).ToShortDateString() } catch {}
            if ( $tds[6].Trim() -like "End*" ) { $supportend_enterprise = $true }

            $sac += New-Object -TypeName PSObject -Property (@{
                "Version" = $version;
                "ServicingOption" = $servicingoption;
                "AvailabilityDate" = $availabilitydate;                    
                "LatestRevisionDate" = $latestrevisiondate;
                "OSBuild" = $osbuild;            
                "SupportEndHome" = $supportend_home;
                "SupportEndEnterprise" = $supportend_enterprise          
            }) | Select-Object Version, ServicingOption, AvailabilityDate, LatestRevisionDate, OSBuild, SupportEndHome, SupportEndEnterprise
        }
    }
}
#$sac | Out-GridView -Title $title -OutputMode Multiple

$ltsc = @()

if ( $tables.Count -ge 2 ) {
    
    $table = $tables[1]
    #$table | Out-GridView -OutputMode Multiple

    $title = $table | Select-String -Pattern '(?<=>)(.*?)(?=</h4>)' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    $trs = $table | Select-String -Pattern '<tr(.|\n)*?<\/tr>' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Skip 1
    #$trs | Out-GridView -OutputMode Multiple

    $trs | % {
        
        $tr = $_
        $tds = $tr | Select-String -Pattern '(?<=>)(.*?)(?=</td>)' -AllMatches  | % { $_.Matches } | % { $_.Value }
        #$tds | Out-GridView -OutputMode Multiple

        if ( $tds.Count -ge 7 ) {
            
            $version = $tds[0].Replace("<td>","").Split(" ")[0]

            $servicingoption = $tds[1]            
            
            $availabilitydate = $null
            try { $availabilitydate = ([datetime]$tds[2]).ToShortDateString() } catch {}
            
            $latestrevisiondate = $null
            try { $latestrevisiondate = ([datetime]$tds[3]).ToShortDateString() } catch {}

            $osbuild = $tds[4]

            $supportend_mainstream = $null
            try { $supportend_mainstream = ([datetime]$tds[5]).ToShortDateString() } catch {}
            if ( $tds[5].Trim() -like "End*" ) { $supportend_mainstream = "EndOfService" }

            $supportend_enterprise = $null
            try { $supportend_enterprise = ([datetime]$tds[6].Split("(")[0]).ToShortDateString() } catch {}
            if ( $tds[6].Trim() -like "End*" ) { $supportend_enterprise = $true }

            $ltsc += New-Object -TypeName PSObject -Property (@{
                "Version" = $version;
                "ServicingOption" = $servicingoption;
                "AvailabilityDate" = $availabilitydate;                        
                "LatestRevisionDate" = $latestrevisiondate;
                "OSBuild" = $osbuild;        
                "SupportEndMainstream" = $supportend_mainstream;
                "SupportEndEnterprise" = $supportend_enterprise                
            }) | Select-Object Version, ServicingOption, AvailabilityDate, LatestRevisionDate, OSBuild, SupportEndMainstream, SupportEndEnterprise
        }
    }
}
#$ltsc | Out-GridView -Title $title -OutputMode Multiple

#

$releases = @()

if ( $tables.Count -ge 3 ) {
    
    $tables = $wr_html | Select-String -Pattern '<strong(.|\n)*?<\/table>' -AllMatches | % { $_.Matches } | % { $_.Value }
    #$tables = $tables | Select-Object -Skip 2
    #$tables | Out-GridView -OutputMode Multiple

    $tables | % {
        
        $table = $_
        #$table = $tables[0]
        #$table | Out-GridView -OutputMode Multiple
        
        $title = $table | Select-String -Pattern '(?<=>)(.*?)(?=</strong>)' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1
        $version = $title.Split(" ")[1]

        $endofservice = $false
        if ( $table -like "*End of servic*" ) { $endofservice = $true }

        $trs = $table | Select-String -Pattern '<tr(.|\n)*?<\/tr>' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Skip 1
        #$trs | Out-GridView -OutputMode Multiple

        $builds = @()

        $trs | % {
        
            $tr = $_
            $tds = $tr | Select-String -Pattern '(?<=>)(.*?)(?=</td>)' -AllMatches  | % { $_.Matches } | % { $_.Value }
            #$tds | Out-GridView -OutputMode Multiple

            if ( $tds.Count -ge 3 ) {           

                $servicingoptions = @()
                try { 
                
                    ($tds[0].Split(';') | % { $_.Replace("&bull", "").Replace("<span>", "").Replace("</span>", "").Trim() }) | % {
                        
                        $servicingoption = $_

                        $supportend_home = $null
                        $supportend_mainstream = $null
                        $supportend_enterprise = $null

                        $servicingoptions_sac = $sac | Where-Object { ($_.ServicingOption -eq $servicingoption) -and ($_.OSBuild -eq $osbuild) }
                        $supportend_home = $servicingoptions_sac.SupportEndHome
                        $supportend_enterprise = $servicingoptions_sac.SupportEndEnterprise

                        if ( -Not $servicingoptions_sac ) { 
                            $servicingoptions_ltsc = $ltsc | Where-Object { ($_.ServicingOption -like "*$servicingoption*") -and ($_.OSBuild -eq $osbuild) }
                            $supportend_mainstream = $servicingoptions_ltsc.SupportEndMainstream
                            $supportend_enterprise = $servicingoptions_ltsc.SupportEndEnterprise
                        }
                                             
                        $servicingoptions += New-Object -TypeName PSObject -Property (@{
                            "ServicingOption" = $servicingoption;
                            "SupportEndHome" = $supportend_home;
                            "SupportEndMainstream" = $supportend_mainstream;
                            "SupportEndEnterprise" = $supportend_enterprise;                       
                        }) | Select-Object ServicingOption, SupportEndHome, SupportEndMainstream, SupportEndEnterprise

                    }
                
                } catch {}
                
                           
                $availabilitydate = $null
                try { $availabilitydate = ([datetime]$tds[1]).ToShortDateString() } catch {}

                $osbuild = $tds[2] 

                $kbarticleid = $null
                try { $kbarticleid = $tds[3] | Select-String -Pattern '(\d+)' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Last 1  } catch {}

                $builds += New-Object -TypeName PSObject -Property (@{
                    "ServicingOptions" = $servicingoptions
                    "AvailabilityDate" = $availabilitydate;              
                    "OSBuild" = $osbuild;                          
                    "KBArticleID" = $kbarticleid;
                    
                           
                }) | Select-Object ServicingOptions, AvailabilityDate, OSBuild, KBArticleID
                #$builds | Out-GridView -Title $title -OutputMode Multiple
            }
        }

        $osbuild = $builds[0].OSBuild.Split(".")[0]
        $releases += New-Object -TypeName PSObject -Property (@{
            "Version" = $version;
            "OSBuild" = $osbuild;
            "EndOfService" = $endofservice;
            "Builds" = $builds                
        }) | Select-Object Version, OSBuild, EndOfService, Builds

    }
}
#$releases | Out-GridView -Title "Select release" -OutputMode Single

#

Write-Host ""
Write-Host "All builds:"
$releases.Builds | Sort-Object -Property OSBuild | Format-Table

Write-Host "All releases:"
$releases | Sort-Object -Property Version | Format-Table

Write-Host "Releases with no servicing options for Home usage:"
$sac | Where-Object SupportEndHome -eq "EndOfService" | Sort-Object -Property Version | Format-Table

Write-Host "Releases with no servicing options for Home/Enterprise usage:"
$releases | Where-Object EndOfService | Sort-Object -Property Version | Format-Table

Write-Host "Releases no servicing options for Home usage:"
$sac | Where-Object SupportEndHome -ne "EndOfService" | Sort-Object -Property Version | Format-Table

Write-Host "Long-Term servicing options/builds:"
$ltsc | Where-Object SupportEndEnterprise -ne "EndOfService" | Sort-Object -Property Version |  Format-Table

#

do {

    $release = $releases | Out-GridView -Title "Select Release from $url (Last Updated: $($wr_response.LastModified.ToString()))" -OutputMode Single
    if ( -Not $release ) { exit 0 }

    do {
        $build = $release.Builds | Out-GridView -Title "Select Build from Release $($release.Version)" -OutputMode Single
        if ( $build ) { $servicingoption = $build.ServicingOptions | Out-GridView -Title "View ServicingOptions from Build $($build.OSBuild)" -OutputMode Single }
    } while ( $build )

} while ( $true )
