# Pyrenamer v1.04
# Program Process
# 1. Check valid directory structure
# 2. Load valid config file
# 3. Collect files from active directory
# 4. Loop through files doing various checks
#       Check for pdf file
#       Get file type - possible rewrite function from using file name to scanning pdf for keywords
#       Retrieve data from pdf - function may need rewrite
#       Rename pdf file, naming convention and area specific needs dw_123_123
#       Process file using options found in config file


import logging
import helper
import dev_helper
import os
import time


def main():

    # Set current working directory same as this python script
    current_dir = os.getcwd()

    # Set up logging
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                        datefmt='%a, %d %b %Y %H:%M:%S',
                        filename=os.path.join(current_dir, 'renamer_process.log'),
                        filemode='a+')

    logging.info('Starting PDF processor.')
    print('Starting PDF processor.')

    # Needed directory structure
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
    for d in dirs:
        if not os.path.exists(os.path.join(current_dir, d)):
            logging.error('Missing directory found, creating {}.'.format(os.path.join(current_dir, d)))
            print('Missing directory found, creating {}.'.format(os.path.join(current_dir, d)))
            os.mkdir(os.path.join(current_dir, d))

    # Set initial working directory
    working_dir = os.path.join(current_dir, dirs[0])
    logging.info('Working directory is {}.'.format(working_dir))
    print('Working directory is {}.'.format(working_dir))

    # Get config file
    config_file = "config.yaml"
    config = helper.get_config(current_dir, config_file)

    # Department folders
    sub_folders = ['EH', 'FWT', 'RR', 'KB']

    # Loop through department folders and process certificates
    for sub in sub_folders:

        active_directory = os.path.join(working_dir, sub)
        logging.info('Processing files in {} folder.'.format(active_directory))
        print('Processing files in {} folder.'.format(active_directory))

        # Get file list from subdirectory
        files = helper.scan_for_files(active_directory)
        file_count = len(files)
        logging.info('Files found: {}'.format(file_count))
        print('Files found: {}'.format(file_count))
        timestamp = helper.get_timestamp()

        # Iterate through files
        for file in files:

            time.sleep(1)

            # Check if we have a pdf
            if not helper.is_pdf(file):
                logging.info('Invalid file found {}'.format(file))
                print('Invalid file found {}'.format(file))
                pass
            else:
                file_type = helper.get_file_type(os.path.join(active_directory, file))

                # Check we have a known certificate type
                if file_type == "UNKNOWN" or file_type == "DFHN" or file_type == "PARTP":
                    logging.info('Found incorrect file type {} - {}.'.format(file_type, os.path.basename(file)))
                    print('Found incorrect file type {} - {}.'.format(file_type, os.path.basename(file)))
                    pass
                else:
                    logging.info('Found {} start processing {}.'.format(file_type, file))
                    print('Found {} start processing {}.'.format(file_type, file))

                    # Set working dir
                    folder_path = os.path.join(working_dir, sub)

                    # Get needed data from the pdf
                    pdf_data = helper.get_pdf_data(os.path.join(folder_path, file), file_type)

                    # Rename the pdf
                    file_new = helper.rename_pdf_file(file, pdf_data[0], pdf_data[1], file_type, folder_path, pdf_data[3])

                    # End of renaming function config specific actions follow.
                    # Get post renaming actions from config

                    if sub == 'FWT' and file_type == 'EICR':
                        logging.info("Writing data to accu-serv list")
                        print("Writing data to accu-serv list")
                        helper.create_accuserv_list(working_dir, pdf_data, timestamp)

                    if config[sub]['auto_email'] == 'YES':
                        time.sleep(1)

                        if file_type == 'EICR':
                            email_append = 'EICR'
                        else:
                            email_append = 'OTHER'

                        logging.info('Emailing {} certificate {} to {}'.format(
                            sub,
                            os.path.basename(file_new),
                            ''.join(config[sub]['email_recipients_' + email_append])))

                        print('Emailing {} certificate {} to {}'.format(
                            sub,
                            os.path.basename(file_new),
                            ''.join(config[sub]['email_recipients_' + email_append])))

                        helper.email_pdf(file_new, pdf_data[2], config[sub]['email_recipients_' + email_append])
                        if config[sub]['delete_on_send'] == 'YES':
                            logging.info('Deleting {}'.format(file_new))
                            print('Deleting {}'.format(file_new))
                            os.remove(file_new)
                        else:
                            logging.info('Moving {} to processed folder'.format(file_new))
                            print('Moving {} to processed folder'.format(file_new))
                            helper.move_processed_file(working_dir, file_new, sub, pdf_data[3])
                    else:
                        logging.info('Auto emailing for {} disabled'.format(sub))
                        print('Auto emailing for {} disabled'.format(sub))
                        logging.info('Moving {} to processed folder'.format(file_new))
                        print('Moving {} to processed folder'.format(file_new))
                        helper.move_processed_file(working_dir, file_new, sub, pdf_data[3])


if __name__ == "__main__":
    main()
