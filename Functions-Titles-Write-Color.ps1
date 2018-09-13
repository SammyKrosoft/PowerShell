Function Title1 ($title, $TotalLength = 100, $Back = "Yellow", $Fore = "Black") {
    $TitleLength = $Title.Length
    [string]$StarsBeforeAndAfter = ""
    $RemainingLength = $TotalLength - $TitleLength
    If ($($RemainingLength % 2) -ne 0) {
        $Title = $Title + " "
    }
    $Counter = 0
    For ($i=1;$i -le $(($RemainingLength)/2);$i++) {
        $StarsBeforeAndAfter += "*"
        $counter++
    }
    
    $Title = $StarsBeforeAndAfter + $Title + $StarsBeforeAndAfter
    Write-host
    Write-Host $Title -BackgroundColor $Back -foregroundcolor $Fore
    Write-Host
    
}
Function LogMag ($Message){
    Write-Host $message -ForegroundColor Magenta
}

Function LogGreen ($message){
    Write-Host $message -ForegroundColor Green
}

Function LogYellow ($message){
    Write-Host $message -ForegroundColor Yellow
}

Function LogBlue ($message){
    Write-Host $message -ForegroundColor Blue
}

#Examples
cls

Title1 "Part 1 - Checking mailboxes"

For ($i=0;$i -le 10;$i++){
    LogGreen "Mailbox $i - ok"
}

Title1 "Part 2 - Checking databases"

For ($i=0;$i -le 5;$i++){
    LogGreen "Database $i - ok"
}
