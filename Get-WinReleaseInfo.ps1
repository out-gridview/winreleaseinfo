$url = "https://winreleaseinfoprod.blob.core.windows.net/winreleaseinfoprod/en-US.html"

if ( -Not $wr_response ) {
    $wr = [System.Net.WebRequest]::Create($url)
    $wr_response = $wr.GetResponse()
    $wr_html = (New-Object System.IO.StreamReader ($wr_response.GetResponseStream())).ReadToEnd()
}

Write-Host -NoNewline "Get Tables... "
$tables = $wr_html | Select-String -Pattern '<span(.|\n)*?<\/table>' -AllMatches | % { $_.Matches } | % { $_.Value }
Write-Host $tables.Count
#$tables | Out-GridView -OutputMode Multiple

#

$sac = @()

if ( $tables.Count -ge 1 ) {
    
    $table = $tables[0]
    #$table | Out-GridView -OutputMode Multiple

    $title = $table | Select-String -Pattern '(?<=>)(.*?)(?=</span>)' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

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

            $osbuild = $tds[3]

            $latestrevisiondate = $null
            try { $latestrevisiondate = ([datetime]$tds[4]).ToShortDateString() } catch {}

            $supportend_home = $null
            try { $supportend_home = ([datetime]$tds[5]).ToShortDateString() } catch {}
            if ( $tds[5].Trim() -eq "End of service" ) { $supportend_home = "EndOfService" }

            $supportend_enterprise = $null
            try { $supportend_enterprise = ([datetime]$tds[6]).ToShortDateString() } catch {}
            if ( $tds[6].Trim() -eq "End of service" ) { $supportend_enterprise = $true }

            $microsoftrecommends = $false
            try { if ( $tds[7].Trim() -eq "Microsoft recommends" ) { $microsoftrecommends = $true } } catch {}

            $sac += New-Object -TypeName PSObject -Property (@{
                "Version" = $version;
                "ServicingOption" = $servicingoption;
                "AvailabilityDate" = $availabilitydate;
                "OSBuild" = $osbuild;                
                "LatestRevisionDate" = $latestrevisiondate;
                "SupportEndHome" = $supportend_home;
                "SupportEndEnterprise" = $supportend_enterprise;
                "MicrosoftRecommends" = $microsoftrecommends                
            }) | Select-Object Version, ServicingOption, AvailabilityDate, OSBuild, LatestRevisionDate, SupportEndHome, SupportEndEnterprise, MicrosoftRecommends
        }
    }
}
#$sac | Out-GridView -Title $title -OutputMode Multiple

$ltsc = @()

if ( $tables.Count -ge 2 ) {
    
    $table = $tables[1]
    #$table | Out-GridView -OutputMode Multiple

    $title = $table | Select-String -Pattern '(?<=>)(.*?)(?=</span>)' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    $trs = $table | Select-String -Pattern '<tr(.|\n)*?<\/tr>' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Skip 1
    #$trs | Out-GridView -OutputMode Multiple

    $trs | % {
        
        $tr = $_
        $tds = $tr | Select-String -Pattern '(?<=>)(.*?)(?=</td>)' -AllMatches  | % { $_.Matches } | % { $_.Value }
        #$tds | Out-GridView -OutputMode Multiple

        if ( $tds.Count -eq 7 ) {
            
            $version = $tds[0].Replace("<td>","").Split(" ")[0]

            $servicingoption = $tds[1]            
            
            $availabilitydate = $null
            try { $availabilitydate = ([datetime]$tds[2]).ToShortDateString() } catch {}
            
            $osbuild = $tds[3]

            $latestrevisiondate = $null
            try { $latestrevisiondate = ([datetime]$tds[4]).ToShortDateString() } catch {}

            $supportend_mainstream = $null
            try { $supportend_mainstream = ([datetime]$tds[5]).ToShortDateString() } catch {}
            if ( $tds[5].Trim() -eq "End of service" ) { $supportend_mainstream = "EndOfService" }

            $supportend_enterprise = $null
            try { $supportend_enterprise = ([datetime]$tds[6]).ToShortDateString() } catch {}
            if ( $tds[6].Trim() -eq "End of service" ) { $supportend_enterprise = $true }

            $ltsc += New-Object -TypeName PSObject -Property (@{
                "Version" = $version;
                "ServicingOption" = $servicingoption;
                "AvailabilityDate" = $availabilitydate;
                "OSBuild" = $osbuild;                
                "LatestRevisionDate" = $latestrevisiondate;
                "SupportEndMainstream" = $supportend_mainstream;
                "SupportEndEnterprise" = $supportend_enterprise                
            }) | Select-Object Version, ServicingOption, AvailabilityDate, OSBuild, LatestRevisionDate, SupportEndMainstream, SupportEndEnterprise
        }
    }
}
#$ltsc | Out-GridView -Title $title -OutputMode Multiple

#

$releases = @()

if ( $tables.Count -ge 3 ) {

    $tables | Select-Object -Skip 2 | % {

        $table = $_
        #$table | Out-GridView -OutputMode Multiple
        
        $title = $table | Select-String -Pattern '(?<=>) (.*?)(?=\))' -AllMatches | % { $_.Matches } | % { "$($_.Value.Trim()))" } | Select-Object -First 1
        $version = $title.Split(" ")[1]

        $endofservice = $false
        if ( $table -like "*End of service*" ) { $endofservice = $true }

        $trs = $table | Select-String -Pattern '<tr(.|\n)*?<\/tr>' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Skip 1
        #$trs | Out-GridView -OutputMode Multiple

        $builds = @()

        $trs | % {
        
            $tr = $_
            $tds = $tr | Select-String -Pattern '(?<=>)(.*?)(?=</td>)' -AllMatches  | % { $_.Matches } | % { $_.Value }
            #$tds | Out-GridView -OutputMode Multiple

            if ( $tds.Count -ge 3 ) {
                
                $osbuild = $tds[0]

                $availabilitydate = $null
                try { $availabilitydate = ([datetime]$tds[1]).ToShortDateString() } catch {}

                $servicingoptions = @()
                try { 
                
                    ($tds[2].Split(';') | % { $_.Replace("&bull", "").Replace("<span>", "").Replace("</span>", "").Trim() }) | % {
                        
                        $servicingoption = $_

                        $supportend_home = $null
                        $supportend_mainstream = $null
                        $supportend_enterprise = $null
                        $microsoftrecommends = $false

                        $servicingoptions_sac = $sac | Where-Object { ($_.ServicingOption -eq $servicingoption) -and ($_.OSBuild -eq $osbuild) }
                        $supportend_home = $servicingoptions_sac.SupportEndHome
                        $supportend_enterprise = $servicingoptions_sac.SupportEndEnterprise
                        $microsoftrecommends = $servicingoptions_sac.MicrosoftRecommends

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
                            "MicrosoftRecommends" = $microsoftrecommends                           
                        }) | Select-Object ServicingOption, SupportEndHome, SupportEndMainstream, SupportEndEnterprise, MicrosoftRecommends

                    }
                
                } catch {}
                 
                $kbarticleid = $null
                try { $kbarticleid = $tds[3] | Select-String -Pattern '(\d+)' -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Last 1  } catch {}

                $builds += New-Object -TypeName PSObject -Property (@{
                    "OSBuild" = $osbuild;
                    "AvailabilityDate" = $availabilitydate;                    
                    "KBArticleID" = $kbarticleid;
                    "ServicingOptions" = $servicingoptions
                           
                }) | Select-Object OSBuild, AvailabilityDate, KBArticleID, ServicingOptions
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

Write-Host "Recommended servicing option/build:"
$sac | Where-Object MicrosoftRecommends | Format-Table

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
