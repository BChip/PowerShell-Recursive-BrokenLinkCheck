Stop-Process -name WINWORD
Clear-Host
$path = 'C:\'
$files = Get-Childitem $path -Include *.docx,*.doc -Recurse | Where-Object { !($_.psiscontainer) }
$application = New-Object -comobject word.application
$application.visible = $False
'Location:' + "," + "Text:" + "," + "Link:" + "," + "Try Failed?" + "," + "Error Report" + "," + "SharePoint Access:" | Add-Content -path 'deadLinks.csv'

Function Get-StringMatch
{
    # Loop through all *.doc files in the $path directory
    Foreach ($file In $files)
    {
        Write-Host $file
        $document = $application.documents.open($file.fullname,$false,$true)
        $hyperlinks = @($document.Hyperlinks) 
        foreach($hyperlink In $hyperlinks) 
        {
            
            try{
                #check for links that are blank
                if(!$hyperlink.Address){
                    continue
                }
                #ignore mail address
                elseif(($hyperlink.Address -match "mailto:") -eq $true){
                    continue
                }
                #check if website works
                elseif(($hyperlink.Address -match "http") -eq $true){
                    $r = Invoke-WebRequest $hyperlink.Address -UseDefaultCredentials
                    if ($r.StatusCode -eq 200) {
                        if($r.Content.IndexOf("Sorry,") -ne -1){
                            $file.fullname + "," + $hyperlink.TextToDisplay + "," + $hyperlink.Address + "," + "NULL" + "," + "NULL" + "," + "ACCESS DENIED" | Add-Content -path 'deadLinks.csv' 
                        } 
                    }
                }
                else{
                    #try and test the path to a directory
                    if ((Test-Path $hyperlink.Address -errorAction Stop) -eq $false) { 
                        $file.fullname + "," + $hyperlink.TextToDisplay + "," + $hyperlink.Address | Add-Content -path 'deadLinks.csv' 
                    }
                }
            }catch{
                $file.fullname + "," + $hyperlink.TextToDisplay + "," + $hyperlink.Address + "," + "True" + "," + $_.Exception.Message | Add-Content -path 'deadLinks.csv' 
                continue
            }
            
        }
	    $document.close()
    }
    
    $application.quit()
}

Get-StringMatch

