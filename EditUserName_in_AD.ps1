cls
$users = Get-ADUser -filter * -Properties SamAccountName -SearchBase "OU=KHV-Users,OU=KHV-Office,OU=Office,DC=bam,DC=loc"

foreach ($user in $users){
    if ($user.Enabled -eq "True"){
        $name = $user.Name.Split(" ")
        if ($name.Count -eq 3){
            if($user.SamAccountName -eq "n.varfolomeev"){
                Set-ADUser $user.SamAccountName -givenName $name[1] -Surname $name[0]
            }
        }
    }
}
