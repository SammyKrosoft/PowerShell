<#
.SYNOPSIS
    This command line will set the Powershell prompt to PS> and the Window Title to the current working directory

.DESCRIPTION
    By default the PowerShell prompt is set to your current work directory or PS Drive location - and the current
    directory prompt can be very long if you are working in a deep sub-directories structure, making your command lines
    difficult to follow sometimes. 
    
    We can change that to your liking, for example to set the prompt to a simple "PS>" (or string you like), and 
    showing the current working directory in the PowerShell Window title. That Window title will update each time
    you change your working location - that is possible thanks to that PowerShell "Prompt" built-in function that
    you can customize to whatever you want it to be !

    As mentionned above, the PowerShell Built-In function called "Prompt" is a function that is called automatically 
    everytime you hit "Enter" on a PowerShell host. Setting the Prompt function to get the current directoy or 
    PS Drive location and update the current PowerShell window will update your current window with the current
    directory or PS Drive location everytime you do a "CD <Directory or PSDrive>".

    In the sample prompt here we are using $Host.UI.RawUI.WindowTitle to change the current PowerShell Window
    title, and we follow the function definition with the "PS>" string.

    See the Related Links section for the URL to the PowerShell prompt built-in function, or to open directly the Microsoft
    documentation about the Prompt, just type

        Get-Help .\This-Script -Online

    
.INPUTS
    None.

.OUTPUTS
    None.

.EXAMPLE
    Function Prompt {"PS>"}
    
    This will just set the user prompt to "PS>"


.EXAMPLE
    Function prompt {$Host.UI.RawUI.WindowTitle = $(Get-Location);"PS>"}

    This will set the Window title to the current location ($(Get-Location)) and will also set the user prompt to "PS>".
    Note the 2 commands inside the curly brackets, separated by a ";"
    
.NOTES
    None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_prompts


#>


function prompt {$Host.UI.RawUI.WindowTitle = $(Get-Location);"PS>"}