'''

Module for user facing utility classes and functions

'''

class Utility:

    def __init__(self):
        self.__password = 'smartsheet'
        self.__access_key = '<your access token here>'
        self.sheet_name = None

    def password_check(self):
        pw = input('Enter password: ')
        if pw != self.__password:
            print('Incorrect password.')
            exit()

    @property
    def access_key(self):
        return self.__access_key
