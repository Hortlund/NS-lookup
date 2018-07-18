# Imports super important stuffs
import xlrd
import dns.resolver

# b variable takes the Excel file of choice, "request" was just the one i was workign with
b = xlrd.open_workbook("request.xls")

#Starts checking on the first sheet in Excel if more are opened
s = b.sheet_by_index(0)

# Starts checking domain one column down, as the first were just "Domains", this depends how your Excel file is formated.
domains = s.col_values(0,1)

#for loop to iterate thru all of the domains in the Excel file.
for domain in domains:
    
    #Creating a new instance of the imported dns resolver 
    myResolver = dns.resolver.Resolver()
    #Try except, as the loop will stop otherwise when it doesnt finds any NS
    try:
        #Querys the domains for their NS and stores the result in myAnswers
        myAnswers = myResolver.query(domain, "NS")
        #For every domain in myAnswers, prints the domains and its NS to console
        for data in myAnswers:
            print(domain, data)
    #When no lookup is found, prints domain and that no NS were found, This can be wrong, depending on circumstances, double check these manually, shouldnt be many
    except:
        print(domain, "No NS lookup found")



