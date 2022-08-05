cls
#$company = ""
#$mail = ""
#$fullName = ""
#$firstName = ""
#$lastName = ""
#$phone = ""
#$sid = ""
#$function = ""
#$loginName = ""
$user_id = ""
$num = 0
#$a = Get-Content -Path C:\temp\users.csv
$users = Get-ADUser -filter * -Properties Company,Enabled,mail,Name,DisplayName,SamAccountName,telephoneNumber,SID,Title,UserPrincipalName -SearchBase "OU=Office,DC=bam,DC=loc"
$userstnd = Get-ADUser -filter * -Properties Company,Enabled,mail,Name,DisplayName,SamAccountName,telephoneNumber,SID,Title,UserPrincipalName -SearchBase "OU=УК-БСМ-Users-Тында,OU=KHV-Users-Outside,OU=KHV-Office,OU=Office,DC=bam,DC=loc"
$admins = "i.zheltov", "n.varfolomeev", "d.savvin", "a.akinshin", "a.solopov", "a.senatorov", "k.avershina", "k.zhukova", "i.zhuravlev", "a.kulukin", "a.melnik"
$t1 = "primary_key;name;first_name;org_id;email;status;function;phone;mobile_phone;notify;employee_number`n"
$t2 = ""
$companies = "УК БСМ Хабаровск", "УК БСМ Комсомольск-на-Амуре", "УК БСМ Тында", "УК БСМ Москва"

foreach ($user in $users){
    $fullName = $user.Name.Split(" ")
    if (($user.Company -in $companies)  -and $user.Enabled -eq "True" -and $fullName.Count -eq 3 -and $user.SamAccountName -notin $admins){
        $num ++
        $_,$_,$_,$_,$_,$_,$_,$user_id = ($user.SID).ToString().Split("-")
        $user_id
        $phones = ""
        $mobilPhone = ""
        if($user.telephoneNumber.Count -ne 0){
            $phones = $user.telephoneNumber.Split('[,;]')
            if($phones.Count -eq 2){
                $mobilPhone = $phones[1].Trim()
            }
        }
        $t1 += $user_id+";"+$fullName[0]+";"+$fullName[1]+";2;"+$user.mail+";Active;"+$user.Title+";"+$phones[0]+";"+$mobilPhone+";Yes;"+$user_id+"`n"
        $t2 += $user_id+";"+$user_id+";"+$user.SamAccountName+";RU RU;profileid->name:Portal user`n"
    }
}
$num

#foreach ($b in $a){
#    $s = $b.split(";")
#    if ($s[0] -eq "УК БСМ Хабаровск" -and $s[1] -eq "True"){
#        $num ++
#        $company = 2
#        $mail = $s[2]
#        $fullName = $s[3]
#        $loginName = $s[5]
#        $phone,$_ = $s[6].Split(",")
#        $sid = $s[7]
#        $function = $s[8]
#        $lastName, $firstName, $_ = $fullName.Split(" ")
#        $_,$_,$_,$_,$_,$_,$_,$user_id = $sid.Split("-")
#        $t1 += $num.ToString()+";$lastName;$firstName;$company;$mail;Active;$function;$phone;Yes;$user_id`n"
#        $t2 += $num.ToString()+";$user_id;$loginName;RU RU;profileid->name:Portal user`n"
#    }
#}
$t1 > C:\temp\persons.csv
$t2 > C:\temp\usersLDAP.csv