$Autoticket="#######"
ForEach ($Object in @(
"Example User 1"
"Example User 2"
)) {
$Autolocate=$Object
Write-Host
Get-User-Info
Offboard-User
}
