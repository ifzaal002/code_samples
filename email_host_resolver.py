__author__ = 'Afzal Ahmad'

from collections import namedtuple
import os
import re
from datetime import datetime
import dns.resolver  # Requires dnspython
import csv

email_host_regex = re.compile(".*@(.*)$")
gmail_servers_regex = re.compile("(.google.com.|.googlemail.com.)$", re.IGNORECASE)
office365_servers_regex = re.compile("(onmicrosoft.[\w]+.com|.onmicrosoft.com|outlook.com)$|(outlook.com.?.*)", re.IGNORECASE)


def is_gmail(email):
  """ Returns True if the supplied Email address is a @gmail.com Email or is a Google Apps for your domain - hosted Gmail address
  Checks are performed by checking the DNS MX records """
  m = email_host_regex.findall(email)
  if m and len(m) > 0:
    host = m[0]
    if host and host != '':
      host = host.lower()
    if host == "gmail.com":
      return True
    else:
      answers = dns.resolver.query(host, 'MX')
      for rdata in answers:
        m = gmail_servers_regex.findall(str(rdata.exchange))
        if m and len(m) > 0:
          return True
  return False


def is_office365(email):
  """ Returns True if the supplied Email address is a office365 Email or is a
  Checks are performed by checking the DNS MX records """
  m = email_host_regex.findall(email)
  if m and len(m) > 0:
    host = m[0]
    if host and host != '':
      host = host.lower()
    if "onmicrosoft" in host or "outlook.com" in host:
      return True
    else:
      answers = dns.resolver.query(host, 'MX')
      for rdata in answers:
        m = office365_servers_regex.findall(str(rdata.exchange))
        if m and len(m) > 0:
          return True
  return False


def get_file_names(input_dir="input_files", output_dir="output_files"):
  """ Assuming that the input and output directories are at the same level where this scripts resides,
    this method generates a list of namedtuples with values input_file and output_file, each containing
    the input file path and output file path.
  """
  res = []
  seed_dir = os.path.join( os.path.dirname(os.path.abspath(__file__)),input_dir)
  output_dir = os.path.join( os.path.dirname(os.path.abspath(__file__)),output_dir)
  iofiles = namedtuple("iofiles",['input_file','output_file'])
  for root, dirs, files in os.walk(seed_dir):
    for file_name in files:
      input_file_path = os.path.normpath(os.path.join(root, file_name))
      output_file_path = os.path.normpath(os.path.join(output_dir, "%s_"%datetime.now().strftime("%Y%m%d-%H%M%S")+file_name))
      res.append(iofiles(input_file_path,output_file_path))
  return res


def remove_control_chars(text):
  ''' Removes non-printable characters from the given string '''
  control_chars = ''.join(map(unichr, range(0, 32) + range(127, 160)))
  control_char_re = re.compile('[%s]' % re.escape(control_chars))
  return control_char_re.sub('', text)


def listify(item):
  if item == None:
    return []
  if isinstance(item, list) or isinstance(item, tuple):
    return item
  return [item]


def identify_individual_email_id_and_write_output(email, csv_dict_writer, line):
  gmail_id = False
  offic365_id = False
  try:
    gmail_id = is_gmail(email)
  except BaseException as e:
    line['Is Gmail Address'] = 'Domain Lookup Failed'
  try:
    if not gmail_id:
      offic365_id = is_office365(email)
  except BaseException as e:
    line['Is Office 365'] = 'Domain Lookup Failed'
    csv_dict_writer.writerow(line)
    return
  line['Is Gmail Address'] = gmail_id
  line['Is Office 365'] = offic365_id
  csv_dict_writer.writerow(line)


def identify_email_host():
  """ Get all the emails and related information from in.csv and write to output.csv after identifying wheather is it
    gmail or not. And output.csv will hold this info in an additional column.csv
    1. Put incoming csv as in.csv in input_csv directory.csv
    2. Run script.csv
    3. Get output file from output_csv.
  """
  res = get_file_names(input_dir="input_csv",output_dir="output_csv")
  dict()
  for iofiles in res:
    with open(iofiles.input_file, "U") as csv_email_file, open(iofiles.output_file, "wb") as csv_output_file:
      csv_reader = csv.DictReader(csv_email_file)
      csv_writer = None
      output_fields = []
      email_re = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[A-Za-z]{2,6}")
      for row in csv_reader:
        row.update({"Is Gmail Address":"Initialized","Is Office 365": "Initialized"})
        if not output_fields:
          output_fields = row.keys()
          csv_writer = csv.DictWriter(csv_output_file, output_fields)
          csv_writer.writeheader()
        email_list = row.get("Email Address")
        if not email_list:
          email_list = row.get("Email")
        if not email_list:
          print "Warning: No Email address found: %s"%row
          row['Is Gmail Address']= "No Email Address Found"
          row['Is Office 365']= "No Email Address Found"
          csv_writer.writerow(row)
          continue
        if "," in email_list:
          email_list = email_list.split(",")
        for email in listify(email_list):
          email_match = email_re.search(email)  # Remove any invalid character in the email address
          if not email_match:
            continue
          email = email_match.group(0)
          # email = line[0].decode('unicode_escape').encode('ascii','ignore') #Remove any invalid character in the email address
          identify_individual_email_id_and_write_output(email, csv_writer, row)

if __name__ == '__main__':
  identify_email_host()