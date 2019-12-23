
import pandas as pd
from pyad import pyad


from pyad.adgroup import ADGroup

path = r"C:\Users\austin.niemann\Desktop\AddStudents\BlankFillLoadFile.csv"

df = pd.read_csv(path, encoding='UTF-8')


df.dropna(axis=0, how="any", thresh=None, subset=None, inplace=True)

size = df.shape
students = max(size)

condition_First = 0
intRow_First = 0
intColumn_First = 0

firstName = (df.iloc[intRow_First, intColumn_First])
lastName = (df.iloc[0, 1])
userName = (df.iloc[0, 2])
group1 = "pldc student"


intRow_First += 1
intColumn_First += 0
condition_First += 1
print(firstName)
print(lastName)
print(userName)

user = firstName + "." + lastName
print(user)

pyad.set_defaults(ldap_server="rti.loc", username="austin.niemann", password="Dillon2018")

ou = pyad.adcontainer.ADContainer.from_dn("OU=WLC, OU=1st Battalion, DC=rti, DC=loc")
new_user = pyad.aduser.ADUser.create(user, ou, password="password", enable=True)


def force_pwd_change_on_login(self):
    self.update_attribute("PwdLastSet", 0)



