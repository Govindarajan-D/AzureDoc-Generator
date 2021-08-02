#Check-DocModule checks for installation of DocumentFormat.OpenXML
function Check-DocModule{
    if(Test-Path "C:\Program Files\PackageManagement\NuGet\Packages\Open-XML-SDK.2.9.1\lib\net46\DocumentFormat.OpenXml.dll"){
         [Reflection.Assembly]::LoadFile("C:\Program Files\PackageManagement\NuGet\Packages\Open-XML-SDK.2.9.1\lib\net46\DocumentFormat.OpenXml.dll")
         return 1
    }
    elseif(Test-Path $UserPath){
        [Reflection.Assembly]::LoadFile("C:\Users\",$env:UserName,"\AppData\Local\PackageManagement\NuGet\Packages\DocumentFormat.OpenXml.2.13.0\lib\net46","DocumentFormat.OpenXml.dll" -join "")
        return 1
    }
    else{
        Write-Host "DocumentFormat.OpenXml is missing. Please install the package"
        return 0

    }
}
#Create-Documentation checks for resource group name for filtering
function Create-Documentation{
    param(
        [ValidateNotNullOrEmpty()]
        $ResourceGroup = "All"
    )
    Connect-AzAccount
    if($ResourceGroup -ne "All"){
        $AzureResources = Get-AzResource -ResourceGroupName $ResourceGroup
        Start-Documentation -AzureResources $AzureResources
    }
}
#Start-Documentation starts the documentation process. It first converts the Powershell output to HTML first and then to DOCX
function Start-Documentation{
    param(
        $AzureResources
    )

    $HTMLObj = foreach($obj in $AzureResources){
       $obj | ConvertTo-Html -As List -Fragment -Property Name,ResourceGroupName,ResourceType,Kind,Location,ResourceId,Tags -PreContent "<h2>$($obj.Name)</h2>"
    }

$header = @"
    <style>

        h1 {

            font-family: Arial, Helvetica, sans-serif;
            color: #e68a00;
            font-size: 28px;

        } 
    </style>

”@

    $Document = ConvertTo-Html -Body $HTMLObj -Head $header -Title "Azure Documentation" -PostContent "<p> Created Date: $(Get-Date)</p>"

    $Document | Out-File AzureDocumentation.html

    $UserPath = "C:\Users\",$env:UserName,"\AppData\Local\PackageManagement\NuGet\Packages\DocumentFormat.OpenXml.2.13.0\lib\net46" -join ""

    #HTML is converted to AltChunk and then inserted to newly created DOCX file
    #AltChunk allows to insert formatted HTML into DOCX directly

    if(Check-DocModule -eq 1){
        $FilePath = (Get-Location).ToString()

        $DocumentPath = ($FilePath,"\AzureDocumentation.docx" -join "").ToString()

        $NewDocument = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Create($DocumentPath,[DocumentFormat.OpenXml.WordprocessingDocumentType]::Document)

        $NewDocument.AddMainDocumentPart()

        $MainDocumentPart = $NewDocument.MainDocumentPart

        $NewDocument.MainDocumentPart.Document = New-Object DocumentFormat.OpenXml.Wordprocessing.Document

        $NewDocument.MainDocumentPart.Document.Body = New-Object DocumentFormat.OpenXml.Wordprocessing.Body

        $AltChunkID = "AltChunkId1";
        
        $Chunk = $MainDocumentPart.AddAlternativeFormatImportPart([DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPartType]::Xhtml, $AltChunkID)

        $AltChunk = New-Object DocumentFormat.OpenXml.Wordprocessing.AltChunk

        $HTMLPath = (Get-Location).ToString()
        
        $FileStream = [System.IO.File]::Open(($HTMLPath,"\AzureDocumentation.html" -join "").ToString(),[System.IO.FileMode]::Open)

        $Chunk.FeedData($FileStream)

        $AltChunk.Id = $AltChunkID
    
        $MainDocumentPart.Document.Body.Append($AltChunk)
    
        $MainDocumentPart.Document.Save()

        $NewDocument.Close()

        $FileStream.Close()
    }
}

$cmd, $params = $args
& $cmd @params