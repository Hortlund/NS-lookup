# Imports super important stuffs
import xlrd
import dns.resolver
import xlsxwriter
import os
import platform
import time

#Global Variables
showServer = dns.resolver.Resolver()

def instructions():
    
    print("**Boi Automation tool**\n")
    print("CURRENT SERVERS USED FOR LOOKUP:",showServer.nameservers,"\n")
    print("What tool do you want to use?\n")
    print("1. NS Lookup\n")
    print("2. A Lookup\n")
    print("3. MX Lookup\n")
    print("4. PTR Lookup\n")
    print("5. CNAME Lookup\n")
    print("6. Domain Lookup(Everything)\n")
    print("7. Server Settings\n")
    print("H. For Help\n")
    print("0. Exit\n")


def osRecognitionClear():
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')

def nsLookUp(mainChoice):

    excelBook = input("Enter the search path of the excel sheet you want to use: ")

    # b variable takes the Excel file of choice, "request" was just the one i was workign with
    b = xlrd.open_workbook(excelBook)

    #Starts checking on the first sheet in Excel if more are opened
    s = b.sheet_by_index(0)

    # Starts checking domain one column down, as the first were just "Domains", this depends how your Excel file is formated.
    domains = s.col_values(0,1)

    workbook = xlsxwriter.Workbook("result.xlsx")
    worksheet = workbook.add_worksheet()

    row = 0
    column = 0
    #for loop to iterate thru all of the domains in the Excel file.
    for domain in domains:
    
        #Creating a new instance of the imported dns resolver
        #myResolver = dns.resolver.Resolver()
        #Try except, as the loop will stop otherwise when it doesnt finds any NS
        try:
            #Querys the domains for their NS and stores the result in myAnswers
            myAnswers = showServer.query(domain, "NS")
            #For every domain in myAnswers, prints the domains and its NS to console
            for data in myAnswers:
                worksheet.write(row, column, domain)
                worksheet.write(row, column + 1 , str(data))
                row += 1
                #print(data)
        #When no lookup is found, prints domain and that no NS were found, This can be wrong, depending on circumstances, double check these manually, shouldnt be many
        except Exception as e:
            worksheet.write(row, column, domain)
            worksheet.write(row, column + 1 , str(e))
            row += 1
            print(str(e))
            time.sleep(1)

    workbook.close()

    osRecognitionClear()
    print("Query Completed")
    time.sleep(2)
    osRecognitionClear()

def serverSettings():
    osRecognitionClear()
    print("Default servers:",showServer.nameservers,"\n")
    serverSet = input("Change servers? y/n\n")
    if (serverSet.upper() == "Y"):
        checkServers = []
        checkServersInput = input("Set first source NS: ")
        checkServers.append(checkServersInput)
        print("Server", checkServers[0], "added!")
        checkServersInput = input("Set second source NS: ")
        checkServers.append(checkServersInput)
        print("Server", checkServers[1], "added!")
        print(checkServers)
        showServer.nameservers = checkServers
        osRecognitionClear()
        start()

    elif (serverSet.upper() == "N"):
        osRecognitionClear()
        start()

def helper():
    osRecognitionClear()
    print("**Help Section**\n")
    print("What do you want to know, no documentation written yet...\n")
    print("Enter to get back")
    input()
    osRecognitionClear()
    start()
    
def start():
    mainLoop = True
    while (mainLoop):
        instructions()
        mainChoice = input()
        if (mainChoice == "0"):
            quit()
        elif (mainChoice.upper() == "H"):
            helper()
        elif (mainChoice == "7"):
            serverSettings()
        nsLookUp(mainChoice)
start()



