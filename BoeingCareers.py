
import sys
import csv
import mailbox
from email.header import decode_header
from datetime import datetime

def get_content(inMessage):
  htmlContent = ''
  payload = inMessage.get_payload()
  if isinstance(payload, str):
    htmlContent += payload
  else:
    for inMessage in payload:
      htmlContent += get_content(inMessage)
  return htmlContent
    
# input
# -----
# standalone
#infile = sys.argv[1]
# IDE
infile = "BoeingCareers.mbox"

# output
# ------
# standalone
#outfile = sys.argv[2]
# IDE
outfile = "BoeingCareers.csv"

rowCount = 0

writer = csv.writer(open(outfile, "w"))
writer.writerow(['date', 'from', 'to', 'jobTitle', 'requisitionNumber', 'PreviewURL', 'content'])

for index, message in enumerate(mailbox.mbox(infile)):
  objDate = datetime.strptime(message['date'], '%a, %d %b %Y %H:%M:%S %z')
  strDate = datetime.strftime(objDate,'%m/%d/%Y')
  theSubject = decode_header(message['subject'])[0][0]
  theBody = get_content(message)
  # title and requisitionNumber processing
  if theSubject == 'Job Posting Notification - The Boeing Company':
    temp = theBody[theBody.find('A position matching your profile for a position of ')+51:theBody.find(' has just been posted at our company.')]
    jobTitle = temp[0:len(temp)-11]
    requisitionNumber = temp[len(temp)-10:len(temp)]
  else:
    jobTitle = theSubject[0:theSubject.find('(')-1]
    requisitionNumber = theSubject[theSubject.find('(')+1:theSubject.find('(')+11]
  # content processing
  if theBody.find('To preview the job description') < 0:
    previewURL = ""
  else:
    previewURL = theBody[theBody.find('To preview the job description')+47:theBody.find('click here')-34]
  row = [
    strDate,
    #message['from'].strip('>').split('<')[-1],
    'Boeing',
    #message['to'],
    'Me',
    jobTitle,
    requisitionNumber,
    previewURL,
    theBody[0:5]
  ]
  writer.writerow(row)
  rowCount += 1

print("Done. " + str(rowCount) + " messages processed.")