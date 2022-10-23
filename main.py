import os

import time
from options import option

# main loop
run = True
while run:
    # home page
    print('-' * 40)
    print("Contact Information Management".center(40))
    print('-' * 40)
    print("[1]: Create file\n"
          "[2]: View all contacts\n"
          "[3]: Search contact\n"
          "[4]: Add contact\n"
          "[5]: Update contact\n"
          "[6]: Delete contact\n")
    #try:
    answer = int(input("Enter the # of your choice: "))
    if 6 >= answer > 0:
        os.system('cls')
        # calling the option function
        option(answer)
        # after the option
        input("Press enter to continue . . .")
        os.system('cls')
        print('[1]: Back to main menu\n'
              '[2]: Exit')
        answer1 = int(input("Enter your choice: "))
        run = False if answer1 == 2 else True

    else:
        print(f"There's no option {answer}")
        time.sleep(1)
        os.system('cls')


    #except ValueError:
        #print("Invalid option, choose again. ")
        #time.sleep(1)
        #os.system('cls')

# exit program
print("\nExiting . . .")
time.sleep(2)
