
Function Say {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")]
        [String]$Msg
    )
    $InstalledVoices = @()
    Add-Type -AssemblyName System.Speech
    $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
    # $InstalledVoices = $Speak.GetInstalledVoices().VoiceInfo
    # $InstalledVoices
    # Select by hint like this ('Male/Female', 'NotSet/Child/Teen/Adult/Senior',[int32]'Position which voices are ordered','fr/en')
    $Speak.SelectVoiceByHints(0,0,0,'en')
        $Speak.Speak($Msg)
}
cls

$msg = 'Hello world ! I love you !'
Say $msg