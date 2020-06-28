import requests, json, csv

# Need token with Mail.Read permissions
token = "placeholder_token"


indexloop = 100
graphUrl = "https://graph.microsoft.com/v1.0/me/messages?$select=sender,subject,receivedDateTime&$top=100"
headers = { 
  "Authorization": "Bearer " + token 
}
scores = {}

def extract(url, indexloop):
  response = requests.get(url, data='',headers=headers)
  r = response.json()

  if "error" in r:
    print(json.dumps(r, indent=2))
    exit()

  for i in r["value"]:
    if "sender" in i:
      if "emailAddress" in i["sender"]:
        if "address" in i["sender"]["emailAddress"]:
          if i["sender"]["emailAddress"]["address"] in scores:
            scores[i["sender"]["emailAddress"]["address"]] += 1
          else:
            scores[i["sender"]["emailAddress"]["address"]] = 1  

  print('Collecting around %d emails...' % indexloop, end='\r')
  indexloop += 100
  
  # Verify in still more emails to crawl
  if "@odata.nextLink" in r:
    extract(r['@odata.nextLink'], indexloop)
  else:
    print('End of collecting with around %d emails' % indexloop)
    print('From %d different senders.' % len(scores))
    convert_to_csv(scores)

def convert_to_csv(i):
  print("Generate csv file with email addresses")
  with open('output.csv', 'w') as csvoutput:
    writer = csv.writer(csvoutput, delimiter=',', quotechar='"',lineterminator='\n')
    for r in i:
      writer.writerow([r, i[r]])
  print("Done")

extract(graphUrl, indexloop)
