# Load the form:
# Older way >>>>> $wpf.EventCollectWindow.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.EventCollectWindow.Dispatcher.InvokeAsync({
    $wpf.EventCollectWindow.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null