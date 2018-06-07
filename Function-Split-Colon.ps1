Function Split-ListColon {
    param(
        [string]$StringToSplit,
        [switch]$Noquotes
    )
    $TargetSplit = $StringToSplit.Split(',')
    $ListItems = ""
    If ($NoQuotes){
        For ($i = 0; $i -lt $TargetSplit.Count - 1; $i++) {$ListItems += $TargetSplit[$i].trim() + (", ")}
        $ListItems += $TargetSplit[$TargetSplit.Count - 1].trim()
    } Else {
        For ($i = 0; $i -lt $TargetSplit.Count - 1; $i++) {$ListItems += ("""") + $TargetSplit[$i].trim() + (""", ")}
        $ListItems += ("""") + $TargetSplit[$TargetSplit.Count - 1].trim() + ("""")
    }
    Return $ListItems
}