Clear-Host
Stop-Process -name WINWORD
$path = 'C:\'
$files = Get-Childitem $path -Include *.docx,*.doc -Recurse | Where-Object { !($_.psiscontainer) }
$application = New-Object -comobject word.application
$application.visible = $False
'Location:'| Add-Content -path 'deadLinks.csv'

Function Get-StringMatch
{
    # Loop through all *.doc files in the $path directory
    Foreach ($file In $files)
    {
        $document = $application.documents.open($file.fullname,$false,$true)
        $hyperlinks = @($document.Hyperlinks) 
        foreach($hyperlink In $hyperlinks) 
        {
            
            try{ 
                if(!$hyperlink.Address){
                    continue
                }
                #ignore mail address
                elseif(($hyperlink.Address -match "mailto:") -eq $true){
                    continue
                }
                #check if website works
                elseif(($hyperlink.Address -match "http") -eq $true){
                    $HTTP_Request = [System.Net.WebRequest]::Create($hyperlink.Address)
                    $HTTP_Response = $HTTP_Request.GetResponse()
                    $HTTP_Status = [int]$HTTP_Response.StatusCode
                    if ($HTTP_Status -eq 200) { 
                        Write-Host "Site is OK!" 
                    }
                    else {
                        $file.fullname + "," + $hyperlink.TextToDisplay + "," + $hyperlink.Address | Add-Content -path 'deadLinks.csv'
                    }
                    $HTTP_Response.Close()
                }
                else{
                    #try and test the path to a directory
                
                    if ((Test-Path $hyperlink.Address -errorAction Stop) -eq $false) { 
                        $file.fullname + "," + $hyperlink.TextToDisplay + "," + $hyperlink.Address | Add-Content -path 'deadLinks.csv' 
                    }
                }
            }catch{
                $file.fullname + "," + $hyperlink.TextToDisplay + "," + $hyperlink.Address | Add-Content -path 'deadLinks.csv' 
                continue
            }
            
        }
	    $document.close()
    }
    
    $application.quit()
}

Get-StringMatch

