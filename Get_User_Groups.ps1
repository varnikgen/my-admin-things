Import-module Activedirectory
$User= read-host "Введите логин пользователя"
Get-ADPrincipalGroupMembership $User | Sort-Object | select -ExpandProperty name
Write-Host " `n `nДля выхода нажмите клавишу Enter..."
Read-Host