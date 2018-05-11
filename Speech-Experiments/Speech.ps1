
Function Speak {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")]
        [String]$Msg
    )
    $InstalledVoices = @()
    Add-Type -AssemblyName System.Speech
    $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
    $InstalledVoices = $Speak.GetInstalledVoices().VoiceInfo
    $InstalledVoices
    #Select by hint like this ('Male/Female', 'NotSet/Child/Teen/Adult/Senior')
    $Speak.SelectVoiceByHints('male','Senior',0,'en')
        $Speak.Speak($Msg)
}
cls

$msg = 'Hello world ! I love you !'
Speak $msg