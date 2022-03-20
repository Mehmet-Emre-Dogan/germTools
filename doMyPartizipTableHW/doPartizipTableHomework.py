try:
    # other libraries
    import warnings
    from time import sleep
    import datetime
    from msvcrt import getch
    import ctypes

    # Selenium Components
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.options import Options

    # pandas
    import pandas as pd

except ImportError as impex:
    print(f"During library import, an error occured: {str(impex)}")

################################# CONFIGURE THEM #################################
DEBUG = False
HEADLESS = True
PATH = "E:\Programs\chromedriver_win32\chromedriver.exe" # chrome driver's path
SLEEP_BTWN_TRIALS = 0.25
SLEEP_BTWN_ITERS = 0.002
##################################################################################

ctypes.windll.kernel32.SetConsoleTitleW("Automated homework doer: Partizip2 to infinitiv and more...")
pd.set_option("max_columns", None) # Show all rows and coloumns
pd.set_option("max_rows", None)    # Show all rows and coloumns
pd.options.display.max_colwidth = 50

def convert(seconds): 
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60

    if minutes < 10:
        if seconds < 10:
            return f"{hour}:0{minutes}:0{seconds}"
        else:
            return f"{hour}:0{minutes}:{seconds}"
    else:
        if seconds < 10:
            return f"{hour}:{minutes}:0{seconds}"
        else:
            return f"{hour}:{minutes}:{seconds}"

print("Reading partizip2Input.txt")
try:
    with open("partizip2Input.txt", "r", encoding="utf-8") as fil:
        words = fil.readlines()
        print("File loaded succesfully")
        print(f"{len(words)} word has been loaded.")
except (Exception, OSError, FileNotFoundError) as ex:
    print(f"An error occured: {str(ex)}")

myOptions = Options()
if HEADLESS:
    myOptions.add_argument("--headless") # Get rid of browser window
myOptions.add_argument("--start-maximized")
myOptions.add_argument("--no-sandbox") # Keep it. Otherwise you may face with problems
myOptions.add_argument("--log-level=3") # Get rid of unnecessary logs in the terminal about chrome
myOptions.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:98.0) Gecko/20100101 Firefox/98.0") # Pretend to we are in thorough firefox as headless chrome causes errors
myBrowser = webdriver.Chrome(PATH, chrome_options=myOptions)
myBrowser.set_window_size(1920, 1080) 

def isLoaded():
    pageState = myBrowser.execute_script('return document.readyState;')
    return pageState == 'complete'

def waitUntilLoaded():
    while not isLoaded():
        print("Waiting page to load")
        sleep(SLEEP_BTWN_TRIALS)
    return

mainUrl = "https://en.bab.la"
conjPath = "/conjugation/german/"
dictiPath = "/dictionary/german-english/"
myBrowser.get(mainUrl)
waitUntilLoaded()

myDict = {"infitive": [], "helpingVerb": [], "partizip2": [], "English": []}

startTime = datetime.datetime.now()
customDT = startTime.strftime('%H:%M:%S')
print(f"{customDT} - Started to do your homework")
for i, word in enumerate(words):
    word = word.strip() # get rid of '\n' at the end
    # wordAscii = word.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace( "ß", "ss")
    habenOrSein = "error helping verb"
    infinitive = "error infinitive"
    engTranslt = "translation error"
    sleep(SLEEP_BTWN_ITERS)
    # Conjugate
    try:
        myBrowser.get(mainUrl + conjPath + word)
        waitUntilLoaded()
        myBrowser.execute_script("window.scrollTo(0, 600)") 
        infinitive = str(myBrowser.find_element_by_xpath('/html/body/main/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/ul/li').text).strip()
        habenOrSein = myBrowser.find_element_by_xpath('/html/body/main/div[2]/div/div[2]/div[2]/div[2]/div/div[4]/div[1]/div[2]').text.split()[0].strip()
        if DEBUG:
            print(f"infinitive: {infinitive}  helpingVerb: {habenOrSein}")
        if habenOrSein == "habe":
            habenOrSein = "haben"
        elif habenOrSein == "bin":
            habenOrSein = "sein"
    except Exception as ex:
        if DEBUG:
            print(ex)
    sleep(SLEEP_BTWN_ITERS)
    # Translate
    try:
        myBrowser.get(mainUrl + dictiPath + infinitive)
        waitUntilLoaded()
        myBrowser.execute_script("window.scrollTo(0, 600)") 
        firstRow = myBrowser.find_element_by_xpath('/html/body/main/div[2]/div/div[2]/div[1]/div/div[2]/div[2]/ul')
        aTags = firstRow.find_elements_by_tag_name("a")
        aTagsCou = len(aTags)
        if DEBUG:
            print(f"a tag count: {aTagsCou}")
        engTranslt = ""
        for j, li in enumerate(aTags):
            engTranslt += li.text
            if j < (aTagsCou - 1):
                engTranslt += " / "
    except Exception as ex:
        if DEBUG:
            print(ex)

    #.get_attribute("innerHTML").strip()
    myDict["partizip2"].append(word)
    myDict["infitive"].append(infinitive)
    myDict["helpingVerb"].append(habenOrSein)
    myDict["English"].append(engTranslt)
    print(f"{i + 1}".rjust(3) + f" - {word} : {engTranslt}")

endTime = datetime.datetime.now()
runTime =  endTime - startTime
customDT = endTime.strftime('%H:%M:%S')
print(f"{customDT} - Homework has been done.")
print(f"Taken time: {convert(runTime.seconds)}")
print('\n')
df = pd.DataFrame(myDict)
df["Number"] = df.index + 1
df = df[["Number", "infitive", "helpingVerb", "partizip2", "English"]]
print(df.to_string(index=False))
print("Writing to excel file...")
customDT = datetime.datetime.now().strftime('%Y-%m-%d_%H.%M.%S')

"""Excel Writing"""

# Please see the below sources for further information
# https://stackoverflow.com/questions/22831520/how-to-do-excels-format-as-table-in-python
# https://xlsxwriter.readthedocs.io/example_pandas_table.html
# https://xlsxwriter.readthedocs.io/working_with_tables.html
# https://stackoverflow.com/questions/17326973/is-there-a-way-to-auto-adjust-excel-column-widths-with-pandas-excelwriter

writer = pd.ExcelWriter(f"Translations_{customDT}.xlsx", engine='xlsxwriter')
df.to_excel(writer, sheet_name="Translations", index=False, startrow=1, header=False)
workbook = writer.book
worksheet = writer.sheets['Translations']
(rowCou, ColCou) = df.shape
columnSettings = [{'header': column} for column in df.columns]
worksheet.add_table(0, 0, rowCou, ColCou - 1, {'columns': columnSettings, 'style': 'Table Style Medium 4'})
for i, column in enumerate(df.columns):
    colLen = df[column].astype(str).str.len().max()
    colLen = max(colLen, len(column)) # colLen is the maximum length of the rows in this column. And len(column) is the length of the header of this column
    worksheet.set_column(first_col=i, last_col=i, width=colLen)
writer.save()

print("Writing completed. Press any key to exit...")
getch()
print("Closing browser, please wait...")
myBrowser.quit()