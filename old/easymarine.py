import os
print('\nHi welcome to Easymarine! What do you want to run?')

while True:
    print('1. Customer create')
    print('2. Invoice sender')
    print('If you dont have added your .csv-file, please do that before making your choice!')
    choice = input('Enter the number of the script you want to run (q to quit): \n')
    os.system('python csv_changer.py')

    if choice == '1':
        os.system('python customer.py')
    elif choice == '2':
        os.system('python invoice.py')
    elif choice.lower() == 'q':
        break
    else:
        print('Invalid choice, please try again.')

    print('\nNice! Anything more?\n')
