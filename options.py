from contact import *
from os import *


def option(answer):
    data = []
    filename = input("Type the filename: ")
    templFilename = filename
    filename += ".xlsx"
    file_exist = path.exists(filename)
    title = input("Type the worksheet title [press enter to skip]: ")
    file = Modify(filename, title)
    # local labels
    labels = ['LAST NAME', 'FIRST NAME', 'CONTACT NUMBER', 'EMAIL', 'HOME ADDRESS']

    # create a file
    if answer == 1:
        if file_exist:
            print(f"\nThe filename \"{templFilename}\" is already exists!\n")

        else:
            print('-' * 40)
            print("Create a file".center(40))
            print('-' * 40)
            n = int(input("Enter the number of person: "))
            for i in range(n):
                print(f'\nPERSON {i + 1}')
                person = {'NO.': i + 1}
                p = {label: input(f'{label}: ') for label in labels}
                person.update(p)
                data.append(person)
            file.create(data)
            print("\nData successfully saved!\n")

    # view all contacts
    elif answer == 2 and file_exist and path.getsize(filename) > 0:
        print('-' * 40)
        print("View all contacts".center(40))
        print('-' * 40)
        data = file.view()
        for idx, dict in enumerate(data):
            for label in file.labels:
                print(f'{label}: {dict[label]}')
            print('\n')

    # search contact
    elif answer == 3 and file_exist and path.getsize(filename) > 0:
        print('-' * 40)
        print("Search contact".center(40))
        print('-' * 40)
        key = input("Search: ").split()
        data = file.search(key)
        labels = file.labels
        if data:
            print(f"\n\"{''.join(key).upper()}\" is found in the contact records!\n")
            for idx, d in enumerate(data):
                for i, label in enumerate(labels):
                    print(f'{label}: {d[i]}')
            print('\n')
        else:
            print(f"\n\"{''.join(key).upper()}\" is not found in the contact records!\n")

    # add contact
    elif answer == 4 and file_exist and path.getsize(filename) > 0:
        print('-' * 40)
        print("Add contact".center(40))
        print('-' * 40)
        n = int(input("Enter the number of person/s to be added: "))
        n_persons = file.get_npersons()
        for i in range(n):
            n_persons += 1
            print(f'\nPERSON {n_persons}')
            person = {'NO.': n_persons}
            p = {label: input(f'{label}: ') for label in labels}
            person.update(p)
            data.append(person)
        file.add(data)
        print("\nData successfully saved!\n")

    # update
    elif answer == 5 and file_exist and path.getsize(filename) > 0:
        print('-' * 40)
        print("Update contact".center(40))
        print('-' * 40)
        key = input("Search: ").split()
        data = file.search(key)
        labels = file.labels
        if data:
            print(f"\n\"{''.join(key).upper()}\" is found in the contact records!\n")
            for idx, d in enumerate(data):
                for i, label in enumerate(labels):
                    print(f'{label}: {d[i]}')
            print('\n')
            num = int(input("Enter the NO. you want to change: "))
            label = input("[LAST NAME] [FIRST NAME] [CONTACT NUMBER] [EMAIL] [HOME ADDRESS]\n"
                          "Choose what do you want to change: ").upper()
            new_data = input(f"{label}: ")
            file.update(num, label, new_data)
            print("\nData successfully changed!\n")

        else:
            print(f"\n\"{''.join(key).upper()}\" is not found in the contact records!\n")

    # delete
    elif answer == 6 and file_exist and path.getsize(filename) > 0:
        print('-' * 40)
        print("Delete contact".center(40))
        print('-' * 40)
        answer1 = int(input(f"[1] Delete all contacts\n"
                            f"[2] Delete a contact\n"
                            f"Enter the # of choice: "))
        if answer1 == 1:
            file.delete_all()
            print("\nAll contacts have been successfully deleted!")
        else:
            data = file.view()
            for idx, dict in enumerate(data):
                for label in file.labels:
                    print(f'{label}: {dict[label]}')
                print('\n')
            num = int(input("Enter the NO. you want to change: "))
            file.delete(num)
            print("\nContact has been successfully deleted!")

    else:
        print("\nThe file does not exists!\n")
