import zeep, wex

nice_file = 'CNS_Companies.xlsx'

print('Connecting to API...')

wsdl = 'https://cns.omxgroup.com/wsdl/CnsIntegrationMetadataService.wsdl'

while True:
    try:
        client = zeep.Client(wsdl=wsdl)
        break
    except:
        input('Unable to connect. Are you connected to internet? Press ENTER to retry.')

all_companies = []

exchanges = client.service.getAllExchanges()
for ex in exchanges:
##    print(ex['id'],ex['name'],ex['saxxessId'] )
    print('Getting companies from: ' + ex['name'] + '...')
    companies = client.service.getListedCompanies(ex['id'])
    for company in companies:
        all_companies.append([company['name'],company['targinIssuerSign'],company['id'],ex['name']])

print('Writing to file...')

while True:
    try:
        wex.wtfl(nice_file,{'Companies':[['Name','Symbol','DPID','Market']] + all_companies})
        input('\nFile Generated, press ENTER to quit.')
        break
    except PermissionError:
        input('Please close the file ' + nice_file + ' first! Press ENTER to retry.')
