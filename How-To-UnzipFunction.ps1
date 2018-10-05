# Unzip function : needs .NET Framework 4.5 for it to work !
function unzip {
    param( [string]$ziparchive, [string]$extractpath )
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory( $ziparchive, $extractpath )
}
