Function MessageBox ($Title = "Validation",$msg="Usage: MessageBox -Title ""Title"" -msg ""Your message"" -Button ""Ok/OkCancel/YesNo"" -Icon ""Question/Information/Error/Warning""",$Button = "YesNo",$Icon = "Question") {
    return [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)
}

Add-Type -AssemblyName presentationframework, presentationcore
MessageBox