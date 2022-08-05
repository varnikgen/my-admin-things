import csv, re
from ldap3 import Server, Connection, SUBTREE


LDAP_SERVER = "bam.loc"
USERNAME = "itop@BAM"
WORD = "password"
SEARCH_TREE = "OU=Office,DC=bam,DC=loc"
PERSONS_FILENAME = "c:\\temp\\persons.csv"
USERS_FILENAME = "c:\\temp\\users.csv"

attr = ['Company','userAccountControl','mail','DisplayName','SamAccountName','telephoneNumber','objectSid','Title','UserPrincipalName', "physicalDeliveryOfficeName", "streetAddress"]
admins = ["i.zheltov", "n.varfolomeev", "d.savvin", "a.akinshin", "a.senatorov", "k.avershina", "k.zhukova", "i.zhuravlev", "a.kulukin", "a.melnik", "g.gusev", "v.danenko"]
companies = {"УК БСМ Хабаровск":2, "УК БСМ Комсомольск-на-Амуре":2, "УК БСМ Тында":2, "УК БСМ Москва":2, "АО БСМ Хабаровск":3}
persons = []
users = []
p_ao = {}
c_ao = {}
num = 0
disabled = [514, 66050]

# соединяюсь с сервером.
server = Server(LDAP_SERVER)
# соедининие для сбора пользователей
conn = Connection(server, user=USERNAME, password=WORD)
conn.bind()
# читаю список пользоватей со "всеми" атрибутами
conn.search(SEARCH_TREE, '(objectCategory=person)', SUBTREE, attributes=attr)
person_list = conn.entries

# формирую из полученных данных список пользователей
for item in person_list:
    # УК
    if (item.SamAccountName.value not in admins) and (item.Company.value in companies) and (len(item.DisplayName.value.split(" ")) == 3) and (item.mail.value):
        # Подсчёт пользователей
        num +=1
        # Забираю уникальный номер пользователя
        sid = item.objectSid.value.split("-")[-1]
        # Разбиваю ФИО на 3 части
        fullname = item.DisplayName.value.split(" ")
        # Извлекаю из поля только то что может быть телефоном
        phone = ""
        if item.telephoneNumber.value: 
            ph=re.findall(r'\d{5}|\+?\d[\s|\(|\-]?\d{3}[\s|\)|\-]?[\s]?\d{3}[\s|\-]?\d{2}[\s|\-]?\d{2}', item.telephoneNumber.value) 
            phone = " ".join(ph)
        # Получаем номер кабинета
        cabinet = item.physicalDeliveryOfficeName.value if item.physicalDeliveryOfficeName.value != None else ""
        # Определяем расположение
        if item.streetAddress.value and "Фабричный" in item.streetAddress.value:
            location = 1
        elif item.streetAddress.value and "Тургенева" in item.streetAddress.value:
            location = 2
        elif item.streetAddress.value and "Гамарника" in item.streetAddress.value:
            location = 3
        else:
            location = 0

        # Сборданных для контактов в iTop
        persons.append({
            "primary_key": item.objectSid.value.split("-")[-1],
            "name": fullname[0],
            "first_name": fullname[1],
            "org_id": companies[item.Company.value],
            "email": item.mail.value,
            "status": "Inactive" if item.userAccountControl.value in disabled else "Active",
            "function": item.Title.value,
            "phone": phone,
            "notify": "Yes",
            "employee_number": sid,
            "cabinet_number": cabinet,
            "location_id": location})
        # Сборданных для пользователей в iTop
        users.append({
            "primary_key": item.objectSid.value.split("-")[-1],
            "contactid": item.objectSid.value.split("-")[-1],
            "login": item.SamAccountName.value,
            "language": "RU RU",
            "profile_list": "profileid->name:Portal user"
        })

# Заношу данные в файл persons.csv
columns = ["primary_key", "name", "first_name", "org_id", "email", "status", "function", "phone", "notify", "employee_number", "cabinet_number", "location_id"]
col = ";".join(columns)
with open(PERSONS_FILENAME, "w+", newline="") as file:
    file.write(col + "\n") #
    writer = csv.DictWriter(file, delimiter=";", fieldnames=columns)
    writer.writerows(persons)

# Заношу данные в файл users.csv
columns = ["primary_key", "contactid", "login", "language", "profile_list"]
col = ";".join(columns)
with open(USERS_FILENAME, "w+", newline="") as file:
    file.write(col + "\n") #
    writer = csv.DictWriter(file, delimiter=";", fieldnames=columns)
    writer.writerows(users)
print(num)
