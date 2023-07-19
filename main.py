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

    print("Starting PDF processor.")

    # Set current working directory same as this python script
    current_dir = os.getcwd()

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
            print('Missing directory found, creating {}.'.format(os.path.join(current_dir, d)))
            os.mkdir(os.path.join(current_dir, d))

    # Set initial working directory
    working_dir = os.path.join(current_dir, dirs[0])
    print('Working directory is {}.'.format(working_dir))

    # Get config file
    config_file = "config.yaml"
    config = helper.get_config(current_dir, config_file)

    # Set up logging
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                        datefmt='%a, %d %b %Y %H:%M:%S',
                        filename=os.path.join(current_dir, 'renamer.log'),
                        filemode='a+')

    # Department folders
    sub_folders = ['EH', 'FWT', 'RR', 'KB']

    # Loop through department folders and process certificates
    for sub in sub_folders:

        active_directory = os.path.join(working_dir, sub)
        print('Processing files in {} folder.'.format(active_directory))

        # Get file list from subdirectory
        files = helper.scan_for_files(active_directory)
        file_count = len(files)
        print('Files found: {}'.format(file_count))
        timestamp = helper.get_timestamp()

        # Iterate through files
        for file in files:

            # Check if we have a pdf
            if not helper.is_pdf(file):
                print('Invalid file found {}'.format(file))
                pass
            else:
                file_type = helper.get_file_type(file)

                # Check we have a known certificate type
                if file_type is None:
                    print('Found incorrect file type {}.'.format(os.path.basename(file)))
                    pass
                else:
                    print('Found {} start processing.'.format(file_type))

                    # Set working dir
                    folder_path = os.path.join(working_dir, sub)

                    # Helper function to get text rects used for development
                    # print(pdf_text_finder(os.path.join(folder_path, file)))

                    # Get needed data from the pdf
                    pdf_data = helper.get_pdf_data(os.path.join(folder_path, file), file_type)

                    # Rename the pdf
                    file_new = helper.rename_pdf_file(file, pdf_data[0], pdf_data[1], file_type, folder_path, pdf_data[3])

                    # End of renaming function config specific actions follow.
                    # Get post renaming actions from config

                    if sub == 'FWT':
                        print("Writing data to accu-serv list")
                        helper.create_accuserv_list(working_dir, pdf_data, timestamp)

                    if config[sub]['auto_email'] == 'YES':
                        time.sleep(1)
                        print('Emailing {} certificate {} to {}'.format(
                            sub,
                            os.path.basename(file_new),
                            ''.join(config[sub]['email_recipients'])))
                        helper.email_pdf(file_new, pdf_data[2], config[sub]['email_recipients'], config)
                        if config[sub]['delete_on_send'] == 'YES':
                            print('Deleting {}'.format(file_new))
                            os.remove(file_new)
                        else:
                            print('Moving {} to processed folder'.format(file_new))
                            helper.move_processed_file(working_dir, file_new, sub)
                    else:
                        print('Auto emailing for {} disabled'.format(sub))
                        print('Moving {} to processed folder'.format(file_new))
                        helper.move_processed_file(working_dir, file_new, sub)






if __name__ == "__main__":
    main()
