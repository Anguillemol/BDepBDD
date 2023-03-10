from  cryptography.fernet import Fernet

"""
key = Fernet.generate_key()

with open('key.key', 'wb') as key_file:
    key_file.write(key)

fernet = Fernet(key)
encrypted_username = fernet.encrypt(b"acae250d-01e9-4f32-9d65-e06fa388ff60")
encrypted_password = fernet.encrypt(b"8FG7d+Es/DYXCJWN8spbNV6qyU5TQqUsoKmg5HLsHw4=")

with open('config.cfg', 'wb') as config_file:
    config_file.write(encrypted_username + b'\n')
    config_file.write(encrypted_password + b'\n')

print(key)

"""
with open('key.key', 'rb') as key_file:
    key = key_file.read()

print(key)




with open('config.cfg', 'rb') as config_file:
    encrypted_user = config_file.readline()
    encrypted_password = config_file.readline()
fernet = Fernet(key)
username = fernet.decrypt(encrypted_user).decode()
password = fernet.decrypt(encrypted_password).decode()
print("username: " + username)
print("password: " + password)

