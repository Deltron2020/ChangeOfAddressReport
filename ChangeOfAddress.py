import sys
sys.path.append(r'\\network_path\ChangeOfAddress\venv\Lib\site-packages')
# https://stackoverflow.com/questions/15514593/importerror-no-module-named-when-trying-to-run-python-script
import pandas as pd
import logging
from pathlib import Path
import re
pd.set_option('display.max_columns',10)

log_file = r'\\network_path\Change_Of_Address\coa_log.log'
logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s')

def send_email(receiver):
    # https://leimao.github.io/blog/Python-Send-Gmail/
    import smtplib, ssl
    from email.message import EmailMessage

    port = 465  # For SSL
    smtp_server = ""
    sender_email = ""  # Enter your address
    password = ""

    msg = EmailMessage()
    msg.set_content(log_file)
    msg['Subject'] = "***Change Of Address Report Failure***"
    msg['From'] = sender_email
    msg['To'] = receiver

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        server.send_message(msg, from_addr=sender_email, to_addrs=receiver)


logging.info('Attempting to Read File...')
try:
    excel_file = Path(sys.argv[1])  # Testing File >> 'ChangeOfAddress_NewFormat.xls'
    abs_path = excel_file.resolve(strict=True)
except FileNotFoundError:
    logging.info('The File Does Not Exist or the Path was Entered Incorrectly')
    send_email('tyler@gmail.com')
    exit()
else:
    logging.info('Beginning Import Process...')

    try:
        original_df = pd.read_excel(io = excel_file,
                                    sheet_name=0,
                                    header=None,
                                    names = ["FullName", "ChangeDateTime", "Old/NewAddress", "DoB"],
                                    index_col=None,
                                    usecols=("A:D"),
                                    skiprows=1,
                                    skipfooter=1
                                    )
        new_df = original_df

        '''
        ===========================================
        ###### Old/NewAddress Field Break Out  - Original Field ######
        ===========================================
        '''
        new_df['Old/NewAddress'] = new_df['Old/NewAddress'].str.replace('OLD Address:','').str.replace('NEW Address:','').str.split(r'\r\n')

        l_OldAddress = []
        l_OldCityStZip = []
        l_NewAddress = []
        l_NewCityStZip = []

        for index, val in enumerate(new_df['Old/NewAddress'].values):
            l = (list(filter(None,val))) # filters out all the blank/empty chars from replacing the '\r', the 'Old/New
            # Address' is now split into 4 new fields
            try:
                l_OldAddress.append(l[0])
            except:
                logging.info(f'The Old Address value at index {index} is missing')
                l_OldAddress.append('')

            try:
                l_OldCityStZip.append(l[1])
            except:
                logging.info(f'The Old City State Zip value at index {index} is missing')
                l_OldCityStZip.append('')

            try:
                l_NewAddress.append(l[2])
            except:
                logging.info(f'The New Address value at index {index} is missing')
                l_NewAddress.append('')

            try:
                l_NewCityStZip.append(l[3])
            except:
                logging.info(f'The New City State Zip value at index {index} is missing')
                l_NewCityStZip.append('')


        new_df = new_df.drop(labels = 'Old/NewAddress', axis = 1)
        new_df.insert(1, 'OldAddress', l_OldAddress)
        new_df.insert(2, 'OldCityStZip', l_OldCityStZip)
        new_df.insert(3, 'NewAddress', l_NewAddress)
        new_df.insert(4, 'NewCityStZip', l_NewCityStZip)
        #print(new_df['OldAddress'])
        #print(new_df)

        l_OldAddressBreakout = []
        l_OldStreetNum = []
        l_OldAddressProcessed = []

        for e, i in enumerate(new_df['OldAddress']):
            #print(i)
            for index, char in enumerate(i):
                try:
                    if char.isnumeric():
                        continue
                    else:
                        l_OldStreetNum.append(i[:index])
                        l_OldAddressBreakout.append(i[index:].strip())
                        break
                except:
                    logging.exception(Exception)
                    continue

            processed = f'{l_OldStreetNum[e]} {l_OldAddressBreakout[e]}'
            l_OldAddressProcessed.append(processed.strip().split(r' '))

        #print(l_OldStreetNum)
        #print(l_OldAddressBreakout)
        #print(l_OldAddressProcessed)

        '''
        ===========================================
        ###### OldAddress Field Break Out ########
        ===========================================
        '''

        l_OldStreetNum = []
        l_OldStreetDir = []
        l_OldStreet = []
        l_OldApt = []


        for index, val in enumerate(l_OldAddressProcessed):
            #print(index, val, len(val))
            try:
                l_OldStreetNum.append(int(val[0]))
            except:
                logging.info(f'The Old Street Num value at index {index} cannot be converted to an integer')
                l_OldStreetNum.append('')

            if val[1] in ('SE','SW','NE','NW','N','S','E','W','North','South','East','West'):
                l_OldStreetDir.append(val[1])
            else:
                logging.info(f'The Old Street Dir value on index {index} does not appear to be a direction')
                l_OldStreetDir.append('')

            if val[len(val)-2] == 'APT':
                try:
                    l_OldApt.append(''.join(val[len(val)-2:]).replace('APT',''))
                    l_OldStreet.append(' '.join(val[2:len(val)-3]))
                except:
                    logging.info(f'The Old Street & Old Apt values on index {index} did not work')

            elif val[len(val)-2] != 'APT' and (val[len(val)-1]).isnumeric():
                try:
                    l_OldApt.append(str(val[len(val)-1]))
                    l_OldStreet.append(' '.join(val[2:len(val)-2]))
                except:
                    logging.info(f'The Old Street & Old Apt values on index {index} did not work')

            else:
                l_OldApt.append('')
                try:
                    if val[1] not in ('SE', 'SW', 'NE', 'NW', 'N', 'S', 'E', 'W', 'North', 'South', 'East', 'West'):
                        l_OldStreet.append(f'{val[1]}')
                    else:
                        l_OldStreet.append(val[2])
                except:
                    l_OldStreet.append('')

        #print(l_OldStreetNum)
        #print(l_OldStreetDir)
        #print(l_OldStreet)
        #print(l_OldApt)
        '''
        ===========================================
        ###### Old City State Zip Field Break Out ########
        ===========================================
        '''

        l_OldCity = []
        l_OldState = []
        l_OldZip = []

        temp_col = new_df['OldCityStZip'].str.split(r',')

        for index, val in enumerate(temp_col.values):
            l = (list(filter(None,val[1].split(r' '))))
            l_OldCity.append(val[0])
            l_OldState.append(l[0])
            l_OldZip.append(l[1])
            #print(index, val, l)

        '''
        ===========================================
        ###### Name Field Break Out #########
        ===========================================
        '''

        temp_col = new_df['FullName'].str.split(r' ')

        l_FirstName = []
        l_MiddleName = []
        l_LastName = []

        for index, val in enumerate(temp_col.values):
            l = (list(filter(None,val)))
            l_woNumbers = [item for item in l if not item.isdigit()]
            l_FirstName.append(l_woNumbers[0])
            l_MiddleName.append(l_woNumbers[1])
            l_LastName.append(l_woNumbers[len(l_woNumbers)-1])
            #print(index, l)

        '''
        ===========================================
        ###### Inserting & Reordering Fields Before Export ######
        ===========================================
        '''

        new_df.insert(1, 'FirstName', l_FirstName)
        new_df.insert(2, 'MiddleName', l_MiddleName)
        new_df.insert(3, 'LastName', l_LastName)

        new_df.insert(6, 'OldStreetNum', l_OldStreetNum)
        new_df.insert(7, 'OldStreetDir', l_OldStreetDir)
        new_df.insert(8, 'OldStreet', l_OldStreet)
        new_df.insert(9, 'OldApt', l_OldApt)

        new_df.insert(10, 'OldCity', l_OldCity)
        new_df.insert(11, 'OldState', l_OldState)
        new_df.insert(12, 'OldZip', l_OldZip)

        cols = new_df.columns.tolist()
        cols = [cols[0], cols[1], cols[2], cols[3], cols[16], cols[4], cols[5],
                cols[6], cols[7], cols[8], cols[9], cols[10], cols[11], cols[12],
                cols[13], cols[14], cols[15]]

        # Reordered the columns in list, assign that column order to replacement df
        new_df = new_df[cols]

        # Exporting the df as csv file
        new_df.to_csv(path_or_buf = r'\\network_path\Change_Of_Address_Import\Extracted_COA_Data.csv',
                      sep='|',
                      header = False,
                      index = False
                      )

        excel_file.unlink()

    except:
        logging.warning('******===PROCESS FAILURE===******')
        logging.exception(Exception)
        send_email('tyler@gmail.com')

logging.info('Process Complete')

exit()
