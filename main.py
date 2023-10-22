import re
import pyodbc
from tabulate import tabulate
from win32api import SetConsoleTitle
from colorama import init, Fore, Back, Style


def main():
    """Проверка sms и email рассылки клиентам"""

    def header():
        """Заглавие программы"""
        init()
        SetConsoleTitle('Dispatch SMS and EMAIL v1.2')
        print(Back.YELLOW + Fore.BLACK + 'Check dispatch sms and email'.center(32) + Style.RESET_ALL)

    def sms(number):
        """Проверка sms рассылки"""
        command_sms = f"""
        SELECT [idsender]
              ,[phonenumber]
              ,[ordernumber]
              ,[patient]
              ,[sentdate]
          FROM [ServingResults].[dbo].[Sms_Sent]
          where ordernumber = {number}
        """

        print(Fore.YELLOW + f'{number} - dispatch sms' + Fore.RESET)
        sms_header = ('idsender', 'phonenumber', 'ordernumber', 'patient', 'sentdate')
        sms_data = cursor.execute(command_sms).fetchall()
        if sms_data:
            print(Fore.GREEN, end='')
            print(tabulate(sms_data, headers=sms_header))
            print(Fore.RESET)
        else:
            print(Fore.RED + 'Not found!' + Fore.RESET, '\n')

    def email(number):
        """Проверка email рассылки"""
        command_email = f"""
        SELECT [idsender]
              ,[addressto]
              ,[ordernumber]
              ,[sentdate]
          FROM [ServingResults].[dbo].[Private_Pages]
          where ordernumber = {number}
        """

        print(Fore.YELLOW + f'{number} - dispatch email' + Fore.RESET)
        email_header = ('idsender', 'addressto', 'ordernumber', 'sentdate')
        email_data = cursor.execute(command_email).fetchall()
        if email_data:
            print(Fore.BLUE, end='')
            print(tabulate(email_data, headers=email_header))
            print(Fore.RESET)
        else:
            print(Fore.RED + 'Not found' + Fore.RESET, '\n')

    def search_by_mail(mail):
        """Поиск рассылки по почте"""
        command_email = f"""
                SELECT [idsender]
                      ,[addressto]
                      ,[ordernumber]
                      ,[sentdate]
                  FROM [ServingResults].[dbo].[Private_Pages]
                  where addressto = '{mail}'
                """
        print(Fore.YELLOW + f'{mail} - search dispatch by mail' + Fore.RESET)
        email_header = ('idsender', 'addressto', 'ordernumber', 'sentdate')
        email_data = cursor.execute(command_email).fetchall()
        if email_data:
            print(Fore.MAGENTA, end='')
            print(tabulate(email_data, headers=email_header))
            print(Fore.RESET)
        else:
            print(Fore.RED + 'Not found' + Fore.RESET, '\n')

    def find_data_email():
        """Поиск номеров заказа почтового адреса в тексте письма"""
        nonlocal list_data
        print(Fore.GREEN + '\nInput email text and write "end":' + Fore.RESET)
        text = list(iter(input, 'end'))
        print()
        text_convert_string = ' '.join(text)
        list_data = re.findall(r'[19]\d{9}|[\w.-]+@[\w.]+', text_convert_string)
        print(Fore.CYAN, end='')
        print(list_data, '\n', Fore.RESET)

    header()
    list_data = []

    # Основной цикл
    while True:
        find_data_email()
        if list_data:
            for data in set(list_data):
                if data.isdigit():
                    sms(data)
                    email(data)
                else:
                    if data.isascii():
                        search_by_mail(data)
                    else:
                        print(Fore.RED, end='')
                        print('Russian letter -', data)
                        print(Fore.RESET, end='')
                print()
        else:
            print(Fore.RED + 'Order number or mail not found!' + Fore.RESET, '\n')


if __name__ == '__main__':
    with pyodbc.connect('Driver={SQL Server};'
                        'Server=Server;'
                        'Trusted_Connection=YES') as connectDB:
        cursor = connectDB.cursor()
        main()

