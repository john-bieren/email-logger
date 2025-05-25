#!/usr/bin/env python3

'''Log email metadata into Excel'''

from datetime import datetime
from os import listdir, path

import pandas as pd
from pypdf import PdfReader
from tqdm import tqdm

from exception_logger import configure_logger

configure_logger()

def process_eml_line(line, df_row, recipients):
    '''Get data from a reconstructed line from an .eml file'''
    line = line.replace("\n", "").replace("\t", "")
    if line.startswith("From:"):
        # remove the <email> if there's an alias before it
        if line[5:].split("<", maxsplit=1)[0].strip('" ') != "":
            df_row['Sender'] = line[5:].split("<", maxsplit=1)[0].strip('" ')
        else:
            df_row['Sender'] = line[5:].strip('" <>')
    elif line.startswith("To: ") or line.startswith("CC: "):
        for recipient in line[4:].split(">, "):
            # remove the <email> if there's an alias before it
            if recipient.split("<", maxsplit=1)[0].strip('" ') != "":
                recipient = f'{recipient.split("<", maxsplit=1)[0].strip('" ')}'
            # quotation marks added because aliases might be "last, first"
            recipients += f', "{recipient.strip('<>"')}"'
    elif line.startswith("Subject:"):
        df_row['Subject'] = line[8:].strip()
    elif line.startswith("Date:"):
        date_time = line.split(", ", maxsplit=1)[-1] # -1 to avoid IndexError if there's no comma
        date_time = date_time.split(" +", maxsplit=1)[0].split(" -", maxsplit=1)[0].strip()
        day, month, the_rest = date_time.split(" ", maxsplit=2)
        df_row['Date and Time'] = f"{month} {day} {the_rest}"
    return df_row, recipients

def process_emls(eml_dir):
    '''Parse .eml file by combining wrapped lines, collect the data'''
    df = pd.DataFrame()
    emls_logged = non_emls = 0
    for file_name in tqdm(listdir(eml_dir)):
        if file_name[-4:] != ".eml":
            non_emls += 1
            continue
        try:
            df_row = pd.DataFrame([file_name[:-4]], columns=['Message No.'])
            file_path = path.join(eml_dir, file_name)
            recipients = combined_line = ""
            log_line = found_first_from_line = False
            with open(file_path, 'r', encoding='utf-8') as file:
                for line in file.readlines():
                    # if the line isn't a continuation of the previous line
                    if all(not line.startswith(s) for s in (' ', '\n', '\t')):
                        if log_line:
                            df_row, recipients = process_eml_line(combined_line, df_row, recipients)
                        combined_line = ""
                        log_line = False
                    # if we have all the info that we want, we can move on from the file
                    have_info = all(c in df_row.columns for c in ('Sender', 'Subject', 'Date and Time'))
                    if recipients and have_info:
                        break
                    if any(line.startswith(s) for s in ('To:', 'CC:', 'Subject:', 'Date:')):
                        log_line = True
                    elif line.startswith("From:"):
                        log_line = True
                        # if there's already been a "from" line, this is now a different email
                        if found_first_from_line:
                            break
                        found_first_from_line = True
                    # combine multi-line headers before processing them
                    combined_line += line
        except Exception as exc:
            raise Exception(f"error thorwn by {file_name}") from exc
        df_row['Recipient(s)'] = recipients.strip(', ')
        df = pd.concat((df, df_row))
        emls_logged += 1
    return df, emls_logged, non_emls

def process_pdfs(pdf_dir, df):
    '''Log the pdfs from the given pdf folder'''
    pdfs_logged = non_pdfs = 0
    for file_name in tqdm(listdir(pdf_dir)):
        if file_name[-4:] != ".pdf":
            non_pdfs += 1
            continue
        file_path = path.join(pdf_dir, file_name)
        try:
            reader = PdfReader(file_path)
            page_count = len(reader.pages)
            df.loc[df['Message No.'] == file_name[:-4], 'Page Count'] = page_count
            pdfs_logged += 1
        except Exception as exc:
            raise Exception(f"error thorwn by {file_name}") from exc
    return df, pdfs_logged, non_pdfs

def save_xlsx(df, log_dir, page_count):
    '''Reindex the dataframe and save it to the Excel spreadsheet'''
    if page_count:
        df = df.reindex(columns=[
                'Message No.', 'Date and Time', 'Page Count', 'Sender', 'Recipient(s)',
                'Subject', 'Exemption', 'Legal Authority'
            ]
        )
    else:
        df = df.reindex(columns=[
                'Message No.', 'Date and Time', 'Sender', 'Recipient(s)',
                'Subject', 'Exemption', 'Legal Authority'
            ]
        )
    spreadsheet_path = path.join(log_dir, "Exemption Log.xlsx")
    with pd.ExcelWriter(spreadsheet_path, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, header=True, index=False)

def log_usage(start_time, run_time, emls_logged, pdfs_logged, eml_dir, pdf_dir, log_dir):
    '''Log info about the usage of the program'''
    file_name = "usage_log.csv"
    try:
        if path.isfile(file_name):
            with open(file_name, "a", encoding='UTF-8') as file:
                file.write(f'{start_time},{run_time},{emls_logged},{pdfs_logged},"{eml_dir}","{pdf_dir}","{log_dir}"\n')
        else:
            with open(file_name, "x", encoding='UTF-8') as file:
                file.write("start time,run time,emls logged,pdfs logged,eml directory,pdf directory,log directory\n")
                file.write(f'{start_time},{run_time},{emls_logged},{pdfs_logged},"{eml_dir}","{pdf_dir}","{log_dir}"\n')
    except PermissionError:
        pass

def main():
    '''Create email log, log usage info'''
    eml_dir = input("Enter the full path to the folder that contains the \033[1mEMLs\033[0m: ").strip('"')
    pdf_dir = input("Enter the full path to the folder that contains the \033[1mPDFs\033[0m (or press enter to skip): ").strip('"')
    log_dir = input("Enter the full path to the folder where the log should be saved: ").strip('"')
    start_time = datetime.now()

    print("Processing .eml files")
    df, emls_logged, non_emls = process_emls(eml_dir)
    if emls_logged == 0:
        raise ValueError(f'directory "{eml_dir}" contains no .eml files')
    if non_emls > 0:
        print(f"Processed {emls_logged} EMLs, skipped {non_emls} other files")
    else:
        print(f"Processed {emls_logged} EMLs")

    if pdf_dir == "":
        page_count = False
        pdfs_logged = 0
    else:
        page_count = True
        print("Processing .pdf files")
        df, pdfs_logged, non_pdfs = process_pdfs(pdf_dir, df)
        if non_pdfs > 0:
            print(f"Processed {pdfs_logged} PDFs, skipped {non_pdfs} other files")
        else:
            print(f"Processed {pdfs_logged} PDFs")

    print("Saving spreadsheet")
    save_xlsx(df, log_dir, page_count)

    run_time = datetime.now() - start_time
    log_usage(start_time, run_time, emls_logged, pdfs_logged, eml_dir, pdf_dir, log_dir)
    print("Complete")

if __name__ == "__main__":
    main()
