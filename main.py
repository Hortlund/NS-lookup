import xlrd
from lxml import html
import requests
import dns.resolver

#myResolver = dns.resolver.Resolver()

b = xlrd.open_workbook("request.xls")

s = b.sheet_by_index(0)

domains = s.col_values(0,1)

#print (domains)

for domain in domains:

    myResolver = dns.resolver.Resolver()
    try:
        myAnswers = myResolver.query(domain, "NS")
        for data in myAnswers:
            print(domain, data)
    except:
        print(domain, "No NS lookup found")



