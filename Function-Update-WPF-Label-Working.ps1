<#
    .NOTES
        You must define $wpf as $wpf XAML reader first
        NOTE: your WPF form must be called "WForm" if we use the below example
        Also, the below example uses a label object named "lblStatus"
        

#>

Function Working-Label {
    # Trick to enable a Label to update during work :
    # Follow with "Dispatcher.Invoke("Render",[action][scriptblobk]{})" or [action][scriptblock]::create({})
    $wpf.lblStatus.Content = "Working ..."
    # $wpf.WForm.Dispatcher.Invoke("Render",[action][scriptblock]::create({}))
    $wpf.WForm.Dispatcher.Invoke("Render",[action][scriptblock]{})
}
