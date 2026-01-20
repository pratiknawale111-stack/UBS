# from cryptography.fernet import Fernet
#
# with open("creds.key", "rb") as f:
#     key = f.read()
#
# fernet = Fernet(key)
#
# with open("creds.txt", "rb") as f:
#     data = f.read()
#
# encrypted = fernet.encrypt(data)
#
# with open("creds.enc", "wb") as f:
#     f.write(encrypted)
#
# print("✅ creds.txt encrypted → creds.enc")



#2
from cryptography.fernet import Fernet

with open("D:/UBS MDAL/creds.key", "rb") as f:
    key = f.read()

fernet = Fernet(key)

with open("D:/UBS MDAL/creds.txt", "rb") as f:
    data = f.read()

encrypted = fernet.encrypt(data)

with open("creds.enc", "wb") as f:
    f.write(encrypted)

print("✅ creds.txt encrypted → creds.enc")
