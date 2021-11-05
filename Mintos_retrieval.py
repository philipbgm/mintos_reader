import imaplib
import email
from email.header import decode_header
import re
from datetime import datetime, date, timedelta
import csv

def clean(text):
  # clean text for creating a folder
  return "".join(c if c.isalnum() else "_" for c in text)

def parse_encoded(string):
  stringsplit = string.split()
  stringneu = ""
  for i in stringsplit:
    if not re.findall("([=?].*[?=])", i):
      stringneu = stringneu + i + " "
    else:
      restring = re.findall("([=?].*[?=])", i)
      encoded_string = restring[0]
      decoded_string, charset = decode_header(encoded_string)[0]
      if charset is not None:
        decoded_string_single = decoded_string.decode(charset)
      else:
        decoded_string_single = decoded_string
      stringneu = stringneu + decoded_string_single + " "
  stringneu = stringneu.rstrip()
  return stringneu

def check_data_len(data):
  row_index = 0
  for row in data:
        if row:  # avoid blank lines
            row_index += 1
  return row_index

def read_credentials():
  fcred = open("credentials.txt", "r")
  lines = fcred.readlines()
  username = re.findall("([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", lines[0])[0]
  password = re.findall("(.+)", lines[1])[0]
  return username,password

debug = 0

# account credentials
try:
  username,password = read_credentials()
  # create an IMAP4 class with SSL 
  mail = imaplib.IMAP4_SSL("outlook.office365.com")
  # authenticate & login
  mail.login(username, password)
  mail.select("inbox")
except:
  print("Fehler beim Abruf der Zugangsdaten")

ldata = []
lheader = ['Art', 'Datum', 'Anfangssaldo', 'Endsaldo', 'Investitionen', 'Secondary', 'Interest', 'Backbuy', 'Backbuy_Interest', 'Secondary_offset', 'Delay_balance', 'Delay', 'Redeption', 'Eingang', 'Tilgung aus R', 'Tilgung aus Kreditr', 'Total', 'Wert']

# Look for Data-CSV file
try:
  # read last date
  with open('data.csv', 'r', newline='') as csvfile:
    reader = csv.reader(csvfile, delimiter=';')
    # Check whether columns are valid
    lheader_check = []
    lheader_check = next(reader)
    ldata = list(reader)
    len_data = len(ldata)
    try:
      if lheader_check == lheader:
        start_date = ldata[len_data-1][1]
        print("Data file erkannt und gültig, Abfrage wird am %s gestartet." % (start_date))
      else:
        print("Data file nicht erkannt, es wird das Postfach vollständig abgeglichen.")
    except:
      print("Error in accessing data-file")
except:
  print("Data file nicht erkannt, es wird das Postfach vollständig abgeglichen.")

# Handle Message
try:
  if start_date:
    mail_start_date = datetime.strptime(start_date, '%d.%m.%Y')+timedelta(2)
    resp_code, mails = mail.search(None, '(FROM "support@mintos.com")', 'SUBJECT "Mintos-Zusammenfassung"', '(since "%s")' % (mail_start_date.strftime('%d-%b-%Y')))
  else:
    resp_code, mails = mail.search(None, '(FROM "support@mintos.com")', 'SUBJECT "Mintos-Zusammenfassung"')
except:
  print("Fehler beim Abrufen der Mails von Mintos")

# Handle Message
#resp_code, mails = mail.search(None, '(FROM "support@mintos.com")', 'SUBJECT "Mintos-Zusammenfassung"')
for mail_id in mails[0].decode().split():
  ldata_single = []
  #if parse_encoded(message.get("Subject")) == "Ihre tägliche Mintos-Zusammenfassung":
  #print("======================================================\n".format(mail_id))
  try:
    resp_code, mail_data = mail.fetch(mail_id, '(RFC822)') ## Fetch mail data.
    message = email.message_from_bytes(mail_data[0][1]) ## Construct Message from mail data
    #print("From       : {}".format(message.get("From")))
    #print("To         : {}".format(message.get("To")))
    #print("Bcc        : {}".format(message.get("Bcc")))
    ##if debug == 1:
    ##  print("Date Mail : {}".format(message.get("Date")))
    #print("Subject    : {}".format(parse_encoded(message.get("Subject"))))
    ##datetime_conv = datetime.strptime(message.get("Date"), '%a, %d %b %Y %H:%M:%S %z (%Z)')
    ##ldata_single.append(datetime_conv.strftime('%d.%m.%Y'))
    ldata_single.append("Zinsen")
    if debug == 1:
      print("Art:  Zinsen")
  
    for part in message.walk():
      if part.get_content_type() == "text/plain":
        body_lines = part.as_string().split("\n")
        print("\n".join(body_lines[:12])) ### Print first 12 lines of message
    
    if message.get_content_type() == "text/html":
      body = message.get_payload(decode=True).decode()
      #print('Message %s\n%s\n' % (mail_id, mail_data[0][1]))

      #ldata_single 1
      try:
        obj_body_match = re.findall("Anfangssaldo (\d{2}[.]{1}\d{2}[.]{1}\d{4})", body)
        ldata_single.append(obj_body_match[0])
        if debug == 1:
          print("Date: ",obj_body_match[0])
      except:
        ldata_single.append(float(0.00))
        print("Date Error")

      #ldata_single 2
      try:
        obj_body_match = re.search("Anfangssaldo \d{2}[.]{1}\d{2}", body)
        obj_body_match_beginning = obj_body_match.span()[1]
        beginning_num = body[obj_body_match_beginning:]
        beginning_num = re.findall("([-]*\d+[.]{1}\d{2})", beginning_num)[:1]
        ldata_single.append(float(beginning_num[0]))
        if debug == 1:
          print("Begin: ",beginning_num[0])
      except:
        ldata_single.append(float(0.00))
        print("Anfangssaldo Error")
      
      #ldata_single 3
      try:
        obj_body_match = re.search("Endsaldo \d{2}[.]{1}\d{2}", body)
        obj_body_match_end = obj_body_match.span()[1]
        end_num = body[obj_body_match_end:]
        end_num = re.findall("([-]*\d+[.]{1}\d{2})", end_num)[:1]
        ldata_single.append(float(end_num[0]))
        if debug == 1: 
          print("Ende: ",end_num[0])
      except:
        ldata_single.append(float(0.00))
        print("Endsaldo Error")
      
      #ldata_single 4
      try:
        #if not re.search("Investitionen in Fina", body):
        obj_body_match = re.search("Investitionen in Darlehen", body)
        obj_body_match_invest = obj_body_match.span()[1]
        invest_num = body[obj_body_match_invest:]
        invest_num = re.findall("([-]*\d+[.]{1}\d{2})", invest_num)[:1]
        ldata_single.append(float(invest_num[0]))
        if debug == 1:
          print("Investitionen in Darlehen: ",invest_num[0])
        #else:
        #  ldata_single.append(float(0.00))
      except:
        ldata_single.append(float(0.00))
      
      #ldata_single 5
      try:
        obj_body_match = re.search("auf dem Sekun", body)
        obj_body_match_second = obj_body_match.span()[1]
        second_num = body[obj_body_match_second:]
        second_num = re.findall("([-]*\d+[.]{1}\d{2})", second_num)[:1]
        ldata_single.append(float(second_num[0]))
        if debug == 1:
          print("Käufe und Verkäufe auf dem Sekundärmarkt: ",second_num[0])
      except:
        ldata_single.append(float(0.00))
      
      #ldata_single 6
      try:
        obj_body_match = re.search("Zinszahlungen[^ ]", body)
        #obj_body_match = re.search("Zinszahlungen</span>", body)
        #if not obj_body_match:
        #  obj_body_match = re.search("Zinszahlungen\r", body)
        obj_body_match_Zinszahlungen = obj_body_match.span()[1]
        zinszahlungen_num = body[obj_body_match_Zinszahlungen:]
        zinszahlungen_num = re.findall("(\d+[.]{1}\d{2})", zinszahlungen_num)[:1]
        ldata_single.append(float(zinszahlungen_num[0]))
        if debug == 1:
          print("Zinszahlungen: ",zinszahlungen_num[0])
      except:
        ldata_single.append(float(0.00))

      #ldata_single 7
      try:
        obj_body_match = re.search("Kreditsumme der R", body)
        obj_body_match_backbuy = obj_body_match.span()[1]
        backbuy_num = body[obj_body_match_backbuy:]
        backbuy_num = re.findall("(\d+[.]{1}\d{2})", backbuy_num)[:1]
        ldata_single.append(float(backbuy_num[0]))
        if debug == 1:
          print("Kreditsumme der Rückkäufe: ",backbuy_num[0])
      except:
        ldata_single.append(float(0.00))

      #ldata_single 8
      try:
        if body.count("Zinszahlungen aus R")==1:
          obj_body_match = re.search("Zinszahlungen aus R", body)
          obj_body_match_backbuy_int = obj_body_match.span()[1]
          backbuy_int_num = body[obj_body_match_backbuy_int:]
          backbuy_int_num = re.findall("(\d+[.]{1}\d{2})", backbuy_int_num)[:1]
          ges = float(backbuy_int_num[0])
          #print(backbuy_int_num[0])
        else:
          obj_body_match = re.search("Zinszahlungen aus R", body)
          obj_body_match_backbuy_int = obj_body_match.span()[1]
          backbuy_int_num = body[obj_body_match_backbuy_int:]
          backbuy_int_num = re.findall("(\d+[.]{1}\d{2})", backbuy_int_num)[:1]
          #print(backbuy_int_num[0])
          backbuy_int_num_2 = body[obj_body_match_backbuy_int:]
          backbuy_int_num_2 = re.findall("(\d+[.]{1}\d{2})", backbuy_int_num_2)[2:3]
          #print(backbuy_int_num_2[0])
          ges = float(backbuy_int_num[0])+float(backbuy_int_num_2[0])
        ldata_single.append(ges)
        if debug == 1:
          print("Zinszahlungen aus Rückkäufen: ",ges)
      except:
        ldata_single.append(float(0.00))

      #ldata_single 9
      try:
        obj_body_match = re.search("Zweitmarkttransaktionen", body)
        obj_body_match_secondary_offset = obj_body_match.span()[1]
        secondary_offset_num = body[obj_body_match_secondary_offset:]
        secondary_offset_num = re.findall("([-]*\d+[.]{1}\d{2})", secondary_offset_num)[:1]
        ldata_single.append(float(secondary_offset_num[0]))
        if debug == 1:
          print("Ab-/Aufschläge bei Zweitmarkttransaktionen: ",secondary_offset_num[0])
      except:
        ldata_single.append(float(0.00))

      #ldata_single 10
      try:
        obj_body_match = re.search("tragungsabgleich der Verzugs", body)
        obj_body_match_delay_balance = obj_body_match.span()[1]
        delay_bal_num = body[obj_body_match_delay_balance:]
        delay_bal_num = re.findall("(\d+[.]{1}\d{2})", delay_bal_num)[:1]
        ldata_single.append(float(delay_bal_num[0]))
        if debug == 1:
          print("Übertragungsabgleich: ",delay_bal_num[0])
      except:
        ldata_single.append(float(0.00))

      #ldata_single 11
      try:
        obj_body_match = re.search("  Verzugsge", body)
        obj_body_match_delay = obj_body_match.span()[1]
        delay_num = body[obj_body_match_delay:]
        delay_num = re.findall("(\d+[.]{1}\d{2})", delay_num)[:1]
        ldata_single.append(float(delay_num[0]))
        if debug == 1:
          print("Verzugsgebühren: ",delay_num[0])
      except:
        ldata_single.append(float(0.00))

      #ldata_single 12
      try:
        obj_body_match = re.search("Tilgungszahlungen[^ ]", body)
        #obj_body_match = re.search("Tilgungszahlungen</span>", body)
        #if not obj_body_match:
        #  obj_body_match = re.search("Tilgungszahlungen\r", body)
        obj_body_match_tilgungszahlungen = obj_body_match.span()[1]
        tilgungszahlungen_num = body[obj_body_match_tilgungszahlungen:]
        tilgungszahlungen_num = re.findall("(\d+[.]{1}\d{2})", tilgungszahlungen_num)[:1]
        ldata_single.append(float(tilgungszahlungen_num[0]))
        if debug == 1:
          print("Tilgungszahlungen: ",tilgungszahlungen_num[0])
      except:
        ldata_single.append(float(0.00))
      
      #ldata_single 13
      try:
        obj_body_match = re.search("Eingehende Zahlungen vom Bankkonto", body)
        obj_body_match_eingang = obj_body_match.span()[1]
        eingang_num = body[obj_body_match_eingang:]
        eingang_num = re.findall("(\d+[.]{1}\d{2})", eingang_num)[:1]
        ldata_single.append(float(eingang_num[0]))
        if debug == 1:
          print("Eingang: ",eingang_num[0])
      except:
        ldata_single.append(float(0.00))
      
      #ldata_single 14
      try:
        obj_body_match = re.search("Investment principal transit reconciliation", body)
        obj_body_match_transit = obj_body_match.span()[1]
        transit_num = body[obj_body_match_transit:]
        transit_num = re.findall("(\d+[.]{1}\d{2})", transit_num)[:1]
        ldata_single.append(float(transit_num[0]))
        if debug == 1:
          print("Investment principal transit reconciliation: ",transit_num[0])
      except:
        ldata_single.append(float(0.00))

      #ldata_single 15
      try:
        obj_body_match = re.search("Tilgungszahlungen aus R", body)
        obj_body_match_til_r = obj_body_match.span()[1]
        til_r_num = body[obj_body_match_til_r:]
        til_r_num = re.findall("(\d+[.]{1}\d{2})", til_r_num)[:1]
        ldata_single.append(float(til_r_num[0]))
        if debug == 1:
          print("Tilgungszahlungen aus Rückkäufen: ",til_r_num[0])
      except:
        ldata_single.append(float(0.00))
      
       #ldata_single 16
      try:
        obj_body_match = re.search("Erhaltene Tilgung aus Kreditr", body)
        obj_body_match_til_k = obj_body_match.span()[1]
        til_k_num = body[obj_body_match_til_k:]
        til_k_num = re.findall("(\d+[.]{1}\d{2})", til_k_num)[:1]
        ldata_single.append(float(til_k_num[0]))
        if debug == 1:
          print("Erhaltene Tilgung aus Kreditrückkauf: ",til_k_num[0])
      except:
        ldata_single.append(float(0.00))


      #ldata_single.append(format(ldata_single[1]+ldata_single[3]+ldata_single[4],'.2f'))
      #ldata_single.append(format(ldata_single[2]-ldata_single[1]+ldata_single[3]+ldata_single[4]-ldata_single[6]+ldata_single[7]-ldata_single[10]-ldata_single[12],'.2f'))
      #ldata_single.append(format(ldata_single[2]-ldata_single[1]+ldata_single[3]+ldata_single[4]-ldata_single[6]-ldata_single[11]-ldata_single[12]-ldata_single[13],'.2f'))
      #ldata_single.append(format(ldata_single[3]-ldata_single[2]+ldata_single[4]+ldata_single[5]-ldata_single[7]-ldata_single[12]-ldata_single[13]-ldata_single[14]-ldata_single[15]-ldata_single[16],'.2f'))
      #ldata_single.append(format(ldata_single[3]+ldata_single[4]+ldata_single[5]+ldata_single[2],'.2f'))
      ldata_single.append(format(-ldata_single[13]-ldata_single[2]+ldata_single[3]-ldata_single[4]-ldata_single[5]-ldata_single[7]-ldata_single[9]-ldata_single[12]-ldata_single[14]-ldata_single[15]-ldata_single[16],'.2f'))
      #-Eingang - Start + Ende - Investitionen - Sekundärmarkt - Investment principal transit - Tilgungszahlungen aus Rückkäufen-Erhaltene Tilgung aus Kreditrückkauf 
      if debug == 1:
        print("Total: ",ldata_single[17])
      ldata.append(ldata_single)
  except:
    print("Cannot read message from ",message.get("Date"))
  if debug == 1:
    print("======================================================\n".format(mail_id))


#Create CSV-File
with open('data.csv', 'w', newline='') as f:
      writer = csv.writer(f, delimiter=';')
      # write the data
      writer.writerow(lheader)
      writer.writerows(ldata)

# close the connection and logout
mail.close()
mail.logout()
