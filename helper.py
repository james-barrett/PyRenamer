import os
import time
import yaml
import fitz
import logging
from yaml.loader import SafeLoader
from win32com.client import Dispatch


def is_pdf(file):
    split_file_name = os.path.splitext(file)
    file_extension = split_file_name[1]
    file_extension = file_extension.lower()
    if file_extension == ".pdf":
        return True
    else:
        return False


def get_config(current_dir, config_file_name):
    logging.info('Loading config file')
    print('Loading config file')
    config_file = os.path.join(current_dir, config_file_name)
    if not os.path.exists(config_file):
        logging.error('Main config file missing, please create.')
        print('Main config file missing, please create.')
        time.sleep(1)
        quit()

    with open(config_file) as f:
        config = yaml.load(f, Loader=SafeLoader)
        logging.info('Config file loaded successfully.')
        print('Config file loaded successfully.')
        return config


# Get list of files within supplied directory
def scan_for_files(directory):
    for path, subdirs, files in os.walk(directory):
        return files


def get_timestamp():
    t = time.localtime()
    timestamp = time.strftime('%b-%d-%Y_%H%M', t)
    return timestamp


def get_pdf_data(file, file_type):
    uprn_rect = ""
    date_rect = ""

    if file_type == "EIC":
        uprn_rect = (678.0, 148.17999267578125, 740.68798828125, 159.1719970703125)
        #date_rect = (258.0, 513.1800537109375, 298.031982421875, 524.1720581054688)
        date_rect = (116.0, 232.17999267578125, 156.031982421875, 243.1719970703125)
        #116.0, 232.17999267578125, 156.031982421875, 243.1719970703125
        #116.0, 232.17999267578125, 156.031982421875, 243.1719970703125
        cert_num_rect = (610.0, 40.220001220703125, 680.0, 51.23600387573242)
        address_line_1_rect = (578.0, 162.17999267578125, 800.0, 173.1719970703125)
        address_line_2_rect = (553.0, 175.17999267578125, 800.0, 186.1719970703125)
        postcode_rect = (588.0, 189.17999267578125, 670, 200.1719970703125)
    elif file_type == "EICR":
        job_no_rect = (400.0, 135.17999267578125, 460.8079833984375, 146.1719970703125)
        # 408.0, 135.17999267578125, 445.8079833984375, 146.1719970703125
        uprn_rect = (582.0, 150.17999267578125, 650.68798828125, 161.1719970703125)
        #date_rect = (673.0, 464.17999267578125, 713.031982421875, 475.1719970703125)
        # original date_rect = (190.0, 276.17999267578125, 230.031982421875, 287.1719970703125)
        date_rect = (190.0, 276.17999267578125, 277.17596435546875, 287.1719970703125)
        #190.0, 276.17999267578125, 230.031982421875, 287.1719970703125
        #190.0, 276.17999267578125, 230.031982421875, 287.1719970703125
        cert_num_rect = (610.0, 40.220001220703125, 680.0, 51.23600387573242)
        address_line_1_rect = (582.0, 162.17999267578125, 800.0, 174.1719970703125)
        address_line_2_rect = (552.0, 177.17999267578125, 800.0, 188.1719970703125)
        postcode_rect = (588.0, 189.17999267578125, 670.0, 200.1719970703125)
    elif file_type == "MW":
        uprn_rect = (572.0, 155.17999267578125, 640.68798828125, 166.1719970703125)
        date_rect = (96.0, 251.17999267578125, 136.031982421875, 262.1719970703125)
        #96.0, 251.17999267578125, 136.031982421875, 262.1719970703125
        cert_num_rect = (610.0, 41.220001220703125, 680.0, 52.23600387573242)
        address_line_1_rect = (578.0, 168.17999267578125, 800.0, 179.1719970703125)
        address_line_2_rect = (555.0, 181.17999267578125, 800.0, 192.1719970703125)
        postcode_rect = (584.0, 195.17999267578125, 670.0, 206.1719970703125)
    elif file_type == "VIS":
        uprn_rect = (572.0, 150.17999267578125, 640.68798828125, 161.1719970703125)
        date_rect = (696.0, 429.17999267578125, 736.031982421875, 440.1719970703125)
        cert_num_rect = (610.0, 40.220001220703125, 680.0, 51.23600387573242)
        address_line_1_rect = (578.0, 163.17999267578125, 800.0, 174.1719970703125)
        address_line_2_rect = (552.0, 177.17999267578125, 800.0, 188.1719970703125)
        postcode_rect = (578.0, 189.17999267578125, 670.0, 200.1719970703125)

    with fitz.open(file) as doc:
        uprn = clean_text(doc[0].get_textbox(uprn_rect))
        date = format_date(clean_text(doc[0].get_textbox(date_rect)))
        cert_num = clean_text(doc[0].get_textbox(cert_num_rect))

        if file_type == "EICR":
            job_no = doc[0].get_textbox(job_no_rect)
        else:
            job_no = ""

        address = clean_text(doc[0].get_textbox(address_line_1_rect)) \
                  + " " + clean_text(doc[0].get_textbox(address_line_2_rect)) \
                  + " " + clean_text(doc[0].get_textbox(postcode_rect))

    return [uprn, date, address, cert_num, job_no]


def rename_pdf_file(file, uprn, date, type, dir, cert_num):

    naming_convention = ""
    clean_uprn = clean_text(uprn)

    if type == "EICR":
        if any(c.isalpha() for c in uprn):
            naming_convention = "C"
        else:
            naming_convention = "D"

    if len(clean_uprn) < 1:
        clean_uprn = "MISSING"

    old_file = os.path.join(dir, file)
    new_file = os.path.join(dir, clean_uprn + "_" + naming_convention + type + "_" + date + ".pdf")

    try:
        os.rename(old_file, new_file)
    except WindowsError as e:
        logging.debug(e)
        try:
            logging.info('FILE RENAME ERROR: Trying to append certificate number to {}.'.format(new_file))
            print('FILE RENAME ERROR: Trying to append certificate number to {}.'.format(new_file))
            os.rename(old_file,
                      os.path.join(dir,
                                   clean_uprn + "_" +
                                   naming_convention + type + "_" + date + "_" + cert_num + ".pdf"))
        except WindowsError as e:
            logging.info('FILE RENAME ERROR: Failed to rename file on second attempt.')
            print('FILE RENAME ERROR: Failed to rename file on second attempt.')
            logging.debug(e)

    return new_file


def move_processed_file(working_dir, file, sub, cert_num):
    processed_dir = os.path.join(working_dir, '_PROCESSED\\' + sub)
    file_name = os.path.basename(file)
    try:
        logging.info('Moving {} to _PROCESSED {} directory.'.format(file_name, sub))
        print('Moving {} to _PROCESSED {} directory.'.format(file_name, sub))
        os.rename(file, os.path.join(processed_dir, file_name))
    except WindowsError as e:
        logging.debug(e)
        try:
            logging.info('FILE MOVE ERROR: Trying to append certificate number to {}, {}'.format(file_name, sub))
            print('FILE MOVE ERROR: Trying to append certificate number to {}, {}'.format(file_name, sub))
            amended_file_name = os.path.splitext(file_name)[0] + '_' + cert_num + '.pdf'
            os.rename(file, os.path.join(processed_dir, amended_file_name))
        except WindowsError as e:
            logging.info('FILE MOVE ERROR: Failed to move file on second attempt.')
            print('FILE MOVE ERROR: Failed to move file on second attempt.')
            logging.debug(e)


def clean_text(item):
    special_characters = ['!', '#', '$', '%', '&', '@', '[', ']', ']', '/', ',']

    for i in special_characters:
            item = item.replace(i, '')
            item = item.strip()

    return item


def create_accuserv_list(working_dir, data, timestamp):
    # Create file
    # data = [uprn, date, address, cert_num, job_no]
    file = open(os.path.join(working_dir, 'accuserv' + '_' + timestamp + '.txt'), 'a+')
    file.write(data[1] + " : " +
               data[0] + " : " +
               data[2] + " : " +
               data[3] + " : " +
               data[4] + '\r\n')
    file.close()


def email_pdf(file, subject, receivers):
    outlook = Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = "".join(receivers)
    message.Subject = subject
    message.Attachments.Add(Source=file)
    message.Body = "Please Find Attached Your Certificate"
    message.Send()


def format_date(d):

    if '-' in d:
        # date_1 = d[:8]
        date = d[-8:]
        date = date[:4] + date[-2:]
    else:
        if len(d) == 8:
            date = d[:4] + d[-2:]
        else:
            date = 'MISSING'

    return date


def get_file_type(file):

    with fitz.open(file) as doc:

        page = doc[0]

        if not page.is_wrapped:
            page.wrap_contents()

        page_0 = str(page.get_text()).upper()

        if "FIRE DETECTION" in page_0:
            return "DFHN"
        elif "CERTIFICATE OF COMPLIANCE" in page_0:
            return "PARTP"
        if "ELECTRICAL INSTALLATION CONDITION" in page_0:
            return "EICR"
        elif "ELECTRICAL INSTALLATION CERTIFICATE" in page_0:
            return "EIC"
        elif "MINOR ELECTRICAL INSTALLATION" in page_0:
            return "MW"
        elif "DOMESTIC VISUAL CONDITION" in page_0:
            return "VIS"
        else:
            return "UNKNOWN"
