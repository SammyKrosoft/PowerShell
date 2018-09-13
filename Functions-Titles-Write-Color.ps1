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