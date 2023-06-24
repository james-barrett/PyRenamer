import fitz
import logging

import yaml
from yaml.loader import SafeLoader

import os
import time


# from win32com.client import Dispatch


def main():
    print("Starting PDF processor.")
    current_dir = os.getcwd()

    dirs = [
        'Certificates',
        'Certificates\\EH',
        'Certificates\\FWT',
        'Certificates\\RR',
        'Certificates\\KB',
        'Certificates\\_PROCESSED',
        'Certificates\\_PROCESSED\\EH',
        'Certificates\\_PROCESSED\\FWT',
        'Certificates\\_PROCESSED\\RR',
        'Certificates\\_PROCESSED\\KB'
    ]

    # Check for missing directory and create if missing.
    dirs_check(dirs, current_dir)

    working_dir = os.path.join(current_dir, 'Certificates')
    print("Working directory is " + working_dir + ".")
    config = get_config(current_dir, "config.yaml")

    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                        datefmt='%a, %d %b %Y %H:%M:%S',
                        filename=os.path.join(current_dir, 'renamer.log'),
                        filemode='a+')

    sub_folders = ['EH', 'FWT', 'RR', 'KB']

    for sub in sub_folders:
        print("Processing files in " + sub + " folder.")
        time.sleep(1)
        process_files(working_dir, config, sub)


def dirs_check(dirs, current_dir):
    for d in dirs:
        if not os.path.exists(os.path.join(current_dir, d)):
            print("Missing directory found, creating {}.".format(os.path.join(current_dir, d)))
            os.mkdir(os.path.join(current_dir, d))


def process_files(working_dir, config, sub):
    # Get file list for EH sub directory
    files = scan_for_files(os.path.join(working_dir, sub))
    file_count = len(files)
    print("Files found: " + str(file_count))
    time.sleep(1)
    timestamp = get_timestamp()
    # Iterate through files
    for file in files:
        # Check if we have a pdf
        if is_pdf(file):
            file_type = get_file_type(file)
            # Check we have a known certificate type
            if file_type != None:
                print("Found " + file_type + " start processing.")
                # Set working folder for EH
                folder_path = os.path.join(working_dir, sub)

                # Helper function to get text rects
                # print(pdf_text_finder(os.path.join(folder_path, file)))

                # Get needed data from the pdf
                print("Grab data from PDF file " + file)
                pdf_data = get_pdf_data(os.path.join(folder_path, file), file_type)
                time.sleep(1)
                # Rename the pdf
                file_new = rename_pdf_file(file, pdf_data[0], pdf_data[1], file_type, folder_path, pdf_data[3])

                if sub == "EH":
                    if config['actions']['auto_email_EH'] == "YES":
                        time.sleep(1)
                        print("Emailing EMPTY HOMES certificate to " + os.path.basename(file_new) + " to " + "".join(
                            config['email_addresses'][sub]))
                        email_pdf(file_new, pdf_data[2], config['email_addresses'][sub], config)
                        if config['actions']['auto_delete_on_send'] == "YES":
                            print("Deleting " + os.path.basename(file_new))
                            delete_file(file_new)
                        else:
                            print("Moving " + os.path.basename(file_new) + " to _PROCESSED folder")
                            move_processed_file(working_dir, file_new, sub)
                        logging.info("Email EH enabled")
                    else:
                        print("Auto emailing for EMPTY HOMES disabled")
                        logging.info("Email EH disabled")
                elif sub == "RR":
                    if config['actions']['auto_email_RR'] == "YES":
                        time.sleep(1)
                        print("Emailing RESPONSIVE REPAIRS certificate to " + os.path.basename(
                            file_new) + " to " + "".join(config['email_addresses'][sub]))
                        email_pdf(file_new, pdf_data[2], config['email_addresses'][sub], config)
                        if config['actions']['auto_delete_on_send'] == "YES":
                            print("Deleting " + os.path.basename(file_new))
                            delete_file(file_new)
                        else:
                            print("Moving " + os.path.basename(file_new) + " to _PROCESSED folder")
                            move_processed_file(working_dir, file_new, sub)
                        logging.info("Email RR enabled")
                    else:
                        print("Auto emailing for RESPONSIVE REPAIRS disabled")
                        logging.info("Email RR disabled")
                elif sub == "FWT":
                    print("Writing data to accu-serv list")
                    create_accuserv_list(working_dir, pdf_data[2], pdf_data[3], pdf_data[4], timestamp)
                    if config['actions']['auto_email_FWT'] == "YES":
                        time.sleep(1)
                        print("Emailing FIXED WIRE TESTING certificate to " + os.path.basename(
                            file_new) + " to " + "".join(config['email_addresses'][sub]))
                        email_pdf(file_new, pdf_data[2], config['email_addresses'][sub], config)
                        if config['actions']['auto_delete_on_send'] == "YES":
                            print("Deleting " + os.path.basename(file_new))
                            delete_file(file_new)
                        else:
                            print("Moving " + os.path.basename(file_new) + " to _PROCESSED folder")
                            move_processed_file(working_dir, file_new, sub)
                        logging.info("Email FWT enabled")
                    else:
                        print("Auto emailing for FIXED WIRE TESTING disabled")
                        logging.info("Email FWT disabled")
                elif sub == "KB":
                    time.sleep(1)
                    if config['actions']['auto_email_KB'] == "YES":
                        time.sleep(1)
                        print("Emailing KITCHEN & BATHROOM certificate to " + os.path.basename(
                            file_new) + " to " + "".join(config['email_addresses'][sub]))
                        email_pdf(file_new, pdf_data[2], config['email_addresses'][sub], config)
                        if config['actions']['auto_delete_on_send'] == "YES":
                            print("Deleting " + os.path.basename(file_new))
                            delete_file(file_new)
                        else:
                            print("Moving " + os.path.basename(file_new) + " to _PROCESSED folder")
                            move_processed_file(working_dir, file_new, sub)
                        logging.info("Email KB enabled")
                    else:
                        print("Auto emailing for KITCHEN & BATHROOM disabled")
                        logging.info("Email KB disabled")
                else:
                    pass
            else:
                pass
        else:
            pass


def get_timestamp():
    t = time.localtime()
    timestamp = time.strftime('%b-%d-%Y_%H%M', t)
    return timestamp


# Create a file of accuserv details for processing
def create_accuserv_list(working_dir, address, cert_no, job_no, timestamp):
    # Create file
    file = open(os.path.join(working_dir, 'accuserv' + '_' + timestamp + '.txt'), 'a+')
    file.write(address + " : " + job_no + " : " + cert_no + '\r\n')
    file.close()


# Get list of files within supplied directory
def scan_for_files(directory):
    for path, subdirs, files in os.walk(directory):
        return files


def delete_file(file):
    os.remove(file)


def move_processed_file(working_dir, file, sub):
    processed_dir = os.path.join(working_dir, '_PROCESSED\\' + sub)
    file_name = os.path.basename(file)
    try:
        os.rename(file, os.path.join(processed_dir, file_name))
    except WindowsError as e:
        delete_file(file)


def get_file_type(file):
    # Split path and file name
    path, file_name = os.path.split(file)
    # Check what type of certificate we have
    if "EIC182C" in file_name:
        return "EIC"
    elif "EICR182C" in file_name:
        return "EICR"
    elif "MWC182C" in file_name:
        return "MW"
    elif "DVCR" in file_name:
        return "VIS"
    else:
        print("Found un-usable file " + os.path.basename(file) + ", already processed?")
        pass


# check that we have a pdf file
def is_pdf(file):
    split_file_name = os.path.splitext(file)
    file_extension = split_file_name[1]
    file_extension = file_extension.lower()
    if file_extension == ".pdf":
        return True
    else:
        return False


# Pass in file and type, returns uprn, date, address, certificate no.
def get_pdf_data(file, file_type):
    uprn_rect = ""
    date_rect = ""

    if file_type == "EIC":
        uprn_rect = (678.0, 148.17999267578125, 740.68798828125, 159.1719970703125)
        date_rect = (258.0, 513.1800537109375, 298.031982421875, 524.1720581054688)
        cert_num_rect = (610.0, 40.220001220703125, 680.0, 51.23600387573242)
        address_line_1_rect = (578.0, 162.17999267578125, 800.0, 173.1719970703125)
        address_line_2_rect = (553.0, 175.17999267578125, 800.0, 186.1719970703125)
        postcode_rect = (588.0, 189.17999267578125, 670, 200.1719970703125)
    elif file_type == "EICR":
        job_no_rect = (400.0, 135.17999267578125, 460.8079833984375, 146.1719970703125)
        # 408.0, 135.17999267578125, 445.8079833984375, 146.1719970703125
        uprn_rect = (582.0, 150.17999267578125, 650.68798828125, 161.1719970703125)
        date_rect = (673.0, 464.17999267578125, 713.031982421875, 475.1719970703125)
        cert_num_rect = (610.0, 40.220001220703125, 680.0, 51.23600387573242)
        address_line_1_rect = (582.0, 162.17999267578125, 800.0, 174.1719970703125)
        address_line_2_rect = (552.0, 177.17999267578125, 800.0, 188.1719970703125)
        postcode_rect = (588.0, 189.17999267578125, 670.0, 200.1719970703125)
    elif file_type == "MW":
        uprn_rect = (572.0, 155.17999267578125, 640.68798828125, 166.1719970703125)
        date_rect = (96.0, 251.17999267578125, 136.031982421875, 262.1719970703125)
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
        date = clean_text(doc[0].get_textbox(date_rect))
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
    time.sleep(1)
    naming_convention = ""
    clean_uprn = clean_text(uprn)

    if type == "EICR":
        if any(c.isalpha() for c in uprn):
            naming_convention = "C"
        else:
            naming_convention = "D"

    uprn_num = "".join(i for i in clean_uprn if not i.isalpha())

    if len(uprn_num) < 1:
        uprn_num = "MISSING"

    old_file = os.path.join(dir, file)
    new_file = os.path.join(dir, uprn_num + "_" + naming_convention + type + "_" + date + ".pdf")

    try:
        os.rename(old_file, new_file)
    except WindowsError as e:
        print("Renaming error possible duplicate file, appending certificate number to file name")
        os.rename(old_file,
                  os.path.join(dir, uprn_num + "_" + naming_convention + type + "_" + date + "_" + cert_num + ".pdf"))
        logging.debug(e)

    logging.info('Renamed : ' + old_file + ' to ' + new_file)
    return new_file


def clean_text(item):
    special_characters = ['!', '#', '$', '%', '&', '@', '[', ']', ']', '_', '/', ',']

    for i in special_characters:
        item = item.replace(i, '')
        item = item.strip()

    return item


def get_config(current_dir, config_file_name):
    print("Loading config file")
    config_file = os.path.join(current_dir, config_file_name)
    if not os.path.exists(config_file):
        print("Main config file missing, please create.")
        time.sleep(1)
        quit()

    with open(config_file) as f:
        config = yaml.load(f, Loader=SafeLoader)
        print("Config file loaded successfully.")
        return config


def email_pdf(file, subject, receivers, config):
    outlook = Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = "".join(receivers)
    message.Subject = subject
    message.Attachments.Add(Source=file)
    message.Body = "Please Find Attached Your Certificate"
    message.Send()


def pdf_text_finder(file):
    doc = fitz.open(file)
    for page in doc:
        wlist = page.get_text_words()
        return wlist


if __name__ == "__main__":
    main()
