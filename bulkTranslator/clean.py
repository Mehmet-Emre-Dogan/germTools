import re
from msvcrt import getch
import ctypes

ctypes.windll.kernel32.SetConsoleTitleW("Word list cleaner")

# Maximize window
consoleWindow = ctypes.windll.kernel32.GetConsoleWindow()
SW_MAXIMIZE = 3
ctypes.windll.user32.ShowWindow(consoleWindow, SW_MAXIMIZE)

print("Reading words.txt")
words = []
prohibitedWords = ["", " ", "\n", "\r", "\r\n"]
processedWords = []

try:
    with open("words.txt", "r", encoding="utf-8") as fil:
        words = fil.readlines()
        print("File loaded succesfully")
except (Exception, OSError, FileNotFoundError) as ex:
    print(f"An error occured: {str(ex)}")
print("Cleaning words...")
print( "*"*50, "\n")

try:
    for word in words:
        word = re.sub(r"(\b(der|die|das)\b|,)|(\b|-en|-er|-e|-nen|-s|-n)", "", word)
        word = re.sub('[-=/.,?\"]', "", word)
        if not word in prohibitedWords:
            processedWords.append(word)
            print(word.strip())
except Exception as ex:
    print(ex)
else:
    print("\n", "*"*50, sep="")
    print("Cleaner has finished job. Writing words to 'words.txt'")
    try:
        with open("words.txt", "w", encoding="utf-8") as fil:
            fil.writelines(processedWords)
            print("Words have been written to file successfully.")
    except:
        print("Unable to write. An error occured.")

print("Press any key to exit...")
garbage = getch()