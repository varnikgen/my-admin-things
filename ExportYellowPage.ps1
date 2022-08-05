cls
$user_id = ""
$num = 0
$users = Get-ADUser -filter * -Properties Company,Enabled,mail,Name,DisplayName,SamAccountName,telephoneNumber,SID,Title,UserPrincipalName -SearchBase "OU=Office,DC=bam,DC=loc"
$userstnd = Get-ADUser -filter * -Properties Company,Enabled,mail,Name,DisplayName,SamAccountName,telephoneNumber,SID,Title,UserPrincipalName -SearchBase "OU=УК-БСМ-Users-Тында,OU=KHV-Users-Outside,OU=KHV-Office,OU=Office,DC=bam,DC=loc"
$admins = "i.zheltov", "n.varfolomeev", "d.savvin", "a.akinshin", "a.solopov", "a.senatorov", "k.avershina", "k.zhukova", "i.zhuravlev", "a.kulukin", "a.melnik"
$t1 = "primary_key;name;first_name;org_id;email;status;function;phone;mobile_phone;notify;employee_number`n"
$t2 = ""
$companies = "УК БСМ Хабаровск", "УК БСМ Комсомольск-на-Амуре", "УК БСМ Тында", "УК БСМ Москва"

foreach ($user in $users){
    if (($user.Company -in $companies)  -and $user.Enabled -eq "True" -and $fullName.Count -eq 3 -and $user.SamAccountName -notin $admins){
        $num ++
        $user_id
        $phones = ""
        $mobilPhone = ""
        if($user.telephoneNumber.Count -ne 0){
            $phones = $user.telephoneNumber.Split('[,;]')
            if($phones.Count -eq 2){
                $mobilPhone = $phones[1].Trim()
            }
        }
        $t1 += ""+$user.DisplayName +";"+$user.mail+";"+$user.Title+";"+$user.telephoneNumber+";"+"`n"
        $t1
    }
}
$num

$t1 > C:\temp\yellowpages.csv