
Function Speak {
    [CmdletBinding(DefaultParameterSetName = "NormalRun")]
    Param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "NormalRun")]
        [String]$Msg
    )
    $InstalledVoices = @()
    Add-Type -AssemblyName System.Speech
    $Speak = New-Object system.Speech.Synthesis.SpeechSynthesizer
    $Speak.GetInstalledVoices() | Foreach {
        $Object = New-Object -TypeName PSObject
        $Object | Add-Member -MemberType NoteProperty -Name AdditionalInfo -Value $($_.VoiceInfo.AdditionalInfo)
        $Object | Add-Member -MemberType NoteProperty -Name Gender -Value $($_.VoiceInfo.Gender)
        $Object | Add-Member -MemberType NoteProperty -Name Name -Value $($_.VoiceInfo.Name)
        $Object | Add-Member -MemberType NoteProperty -Name Culture -Value $($_.VoiceInfo.Culture)
        $Object | Add-Member -MemberType NoteProperty -Name ID -Value $($_.VoiceInfo.ID)
        $InstalledVoices += $Object
    }
    $Speak.Speak($Msg)
}
cls

Speak "Hello world"