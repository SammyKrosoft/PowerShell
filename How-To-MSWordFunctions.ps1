#region Below functions from the following link:
#https://blog.pythian.com/automation-powershell-and-word-templates-let-the-technician-do-tech/

#Function to open a word document:
Function OpenWordDoc($Filename){
    $Word=NEW-Object –comobject Word.Application
    Return $Word.documents.open($Filename)
}

#Function to save a word document as:
Function SaveAsWordDoc($Document, $FileName){
    $Document.Saveas([REF]$Filename)
    $Document.close()
}

#Function to replace a text tag (<TAG_YOUR_NAME>) with something else
Function ReplaceTag($Document, $FindText, $ReplaceWithText){
    $FindReplace=$Document.ActiveWindow.Selection.Find
    $matchCase = $false;
    $matchWholeWord = $true;
    $matchWildCards = $false;
    $matchSoundsLike = $false;
    $matchAllWordForms = $false;
    $forward = $true;
    $format = $false;
    $matchKashida = $false;
    $matchDiacritics = $false;
    $matchAlefHamza = $false;
    $matchControl = $false;
    $read_only = $false;
    $visible = $true;
    $replace = 2;
    $wrap = 1;

    $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl) | Out-Null
}


#Add an image to a bookmark
Function AddImage($Document, $BookmarkName, $ReplaceWithImage){
    $FindReplace=$Document.ActiveWindow
    $FindReplace.Selection.GoTo(-1,0,0,$Document.Bookmarks.item("$BookmarkName"))
    $FindReplace.Selection.InlineShapes.AddPicture("$replacewithImage")
}

    <#Sample using the above functions:

    $TemplateFile = “C:\reports\Template_Report.docx”
    $FinalFile = “C:\reports\FinalReport.docx”

    # Open template file
    $Doc=OpenWordDoc -Filename $TemplateFile

    # Replace text tags
    ReplaceTag –Document $Doc -FindText ‘<client_name>’ -replacewithtext “Pythian”
    ReplaceTag –Document $Doc -FindText ‘<server_name>’ -replacewithtext “WINSRV001”

    # Add image
    AddImage –Document $Doc -BookmarkName ‘img_SomeBookmark’ -ReplaceWithImage “C:\reports\img.png”

    # Save FInal Report
    SaveAsWordDoc –document $Doc –Filename $FinalFile
    #>
    #endregion

#endregion End of Functions from link indicated.

    <# BELOW : Tried to check if doc called is already opened and trying to use the 
    opened one ... not very accurrate method, commenting for now.
    Error handling is in the content of the  $Doc = $MSWord.Documents.Open($Docfile) 
    result : if $false => nothing is in $Doc, meaning the opening failed.#>
    <#
    Title1 "Checking for already opened documents"
    Try {
        $MSWord = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')
        $DocCount = $MSWord.Documents.count
        write-host "Currently $DocCount Word docs opened... checking if a doc named $DocName already exists..." -BackgroundColor yellow -ForegroundColor red
        $CountDocs = 0
        Foreach ($Doc in $MSWord.Documents) {
            $COuntDocs++
            If ($($Doc.Name) -eq $DocNAme) {
            Write-Host "A document with the same name is already opened ... please close it first" -ForegroundColor Red -BackgroundColor Yellow
            Exit
            }
        }
        Write-Host "Found $CountDocs currently opened, no docs with name $DocName opened. Creating a new Word instance !"
        Write-Host "Cleaning variables..."
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSword)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable MSword
    }
    Catch {
        Write-Host "No Word instance opened. Creating a new Word Instance"
    }
#>


    #Text FormFields are also referred to as text Bookmarks in Word
    #The advantage of BookMarks is that it's quicker to find and to update, and you can sort by name and by position
    #in the word document.
    #
    #$Bookmarks = $Doc.Bookmarks | Sort Start

    #Foreach ($BM in $Bookmarks) {
    #    #$BM.Name | Out-Host
    #    $BM | fl * | out-host
    #}
    #
    #Using FormFields for now until I get what's the best to use in Word automation...
