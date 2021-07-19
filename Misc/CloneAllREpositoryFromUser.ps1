[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Unj-GitCloneAllBy($User, $Proxy = $null) {
    (curl -Proxy $Proxy "https://api.github.com/users/$User/repos?page=1&per_page=100").Content | ConvertFrom-Json | %{ $_.clone_url } | %{ & git clone $_ 2>&1 } | % { $_.ToString() }
      # workaround git printing to stderr by @wekempf aka William Kempf
      # https://github.com/dahlbyk/posh-git/issues/109#issuecomment-21638678
}

Unj-GitCloneAllBy -User SammyKrosoft