import csv
from sys import argv
from ldap3 import Server, Connection, SUBTREE
from ldap3.extend.microsoft.addMembersToGroups import ad_add_members_to_groups as addUsersToGroups


LDAP_SERVER = "bam.loc"
USERNAME = "_arch@BAM"
WORD = argv
GROUPOU = "OU=Группы распространения 1С (GPO),OU=Office,DC=bam,DC=loc"
OBJECTCLASS = 'groupOfNames'

base_with_users_dict = {}

server = Server(LDAP_SERVER)
conn = Connection(server, user=USERNAME, password=WORD)
conn.bind()
conn.search(GROUPOU, '(objectCategory=Group)', SUBTREE, attributes={'cn'})
group_list = []

for item in conn.entries:
    if item.cn.value:
        group_list.append(item.cn.value)
print(group_list)

attr = ['userAccountControl','DisplayName','distinguishedName']
conn.search('OU=Office,DC=bam,DC=loc', '(objectCategory=person)', SUBTREE, attributes=attr)
disabled = [514, 66050]
users_dict = {}
for item in conn.entries:
    if item.userAccountControl.value not in disabled:
        users_dict[item.DisplayName.value] = item.distinguishedName.value

def add_user_in_group(user, users_dict, groupDN):
    keys = users_dict.keys()
    if user in keys:
        addUsersToGroups(conn, users_dict.get(user),groupDN)


with open ('users_in_bases.csv', newline='', encoding="utf-8") as csvfile:
    #linereader = csv.reader(csvfile, delimiter=';', quotechar='|')
    base_dict = csv.DictReader(csvfile, dialect='excel', delimiter=';')
    dataset = [dict(row) for row in base_dict]
    for row in dataset:
        user = row.get("User")
        for key in row.keys():
            group_name = "1c_"+key
            groupDN = 'cn='+group_name+','+GROUPOU
            if key != "User":
                if row.get(key) == "1":
                    if base_with_users_dict.get(key):
                        base_with_users_dict[key].append(user)
                        add_user_in_group(user=user, users_dict=users_dict, groupDN=groupDN)
                    else:
                        base_with_users_dict[key] = [user]
                        if group_name not in group_list:
                            attr = {
                                'cn': group_name,
                                'description': 'Группа безопасности для базы '+ key
                            }
                            conn.add(groupDN, 'Group', attr)
                            group_list.append(group_name)
                        
