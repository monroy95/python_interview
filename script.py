from file_observer.observer import folder_observer

if __name__ == '__main__':
    path = './workspace/'
    res = input('Do you want to choose a folder to monitor (y/n): ')

    if res.lower() == 'y':
        path = str(input('Enter the folder path: '))

    folder_observer(path)
