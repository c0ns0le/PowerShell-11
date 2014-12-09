param($path)
([Environment]::UserDomainName + "\" + [Environment]::UserName) | Out-File ./TESTOUTPUT.txt -Encoding ASCII
Add-Content -Path ./TESTOUTPUT.txt -Value "$env:userdomain\$env:username"
[Security.Principal.WindowsIdentity]::GetCurrent().Name | Out-File ./TESTOUTPUT.txt -Append -Encoding ASCII
Add-Content -Path ./TESTOUTPUT.txt -Value "Path: $path"
