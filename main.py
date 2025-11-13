#!/usr/bin/env python3

"""Log email metadata into Excel"""

from datetime import datetime
from os import listdir, path

import pandas as pd
from pypdf import PdfReader
from tqdm import tqdm

from exception_logger import configure_logger

configure_logger()

def main():
    """Create email log, log usage info"""
    eml_dir = input("Enter the full path to the folder that contains the \033[1mEMLs\033[0m: ").strip('"')
    pdf_dir = input("Enter the full path to the folder that contains the \033[1mPDFs\033[0m (or press enter to skip): ").strip('"')
    log_dir = input("Enter the full path to the folder where the log should be saved: ").strip('"')
    start_time = datetime.now()

    print("Processing .eml files")
    df, emls_logged, non_emls = process_emls(eml_dir)
    if emls_logged == 0:
        raise ValueError(f"{eml_dir} contains no .eml files")
    if non_emls > 0:
        print(f"Processed {emls_logged} EMLs, skipped {non_emls} other files")
    else:
        print(f"Processed {emls_logged} EMLs")

    if pdf_dir == "":
        have_page_count = False
        pdfs_logged = 0
    else:
        have_page_count = True
        print("Processing PDF files")
        df, pdfs_logged, non_pdfs = process_pdfs(pdf_dir, df)
        if non_pdfs > 0:
            print(f"Processed {pdfs_logged} PDFs, skipped {non_pdfs} other files")
        else:
            print(f"Processed {pdfs_logged} PDFs")

    print("Saving spreadsheet")
    save_xlsx(df, log_dir, have_page_count)

    run_time = datetime.now() - start_time
    log_usage(start_time, run_time, emls_logged, pdfs_logged, eml_dir, pdf_dir, log_dir)
    print("Complete")

def process_emls(eml_dir):
    """Parse, log the .eml files from the given .eml directory"""
    df = pd.DataFrame()
    emls_logged = non_emls = 0
    for file_name in tqdm(listdir(eml_dir)):
        if file_name[-4:] != ".eml":
            non_emls += 1
            continue

        try:
            df_row = pd.DataFrame([file_name[:-4]], columns=["Message No."])
            file_path = path.join(eml_dir, file_name)
            recipients = combined_line = ""
            log_line = found_first_from_line = False

            with open(file_path, "r", encoding="utf-8") as file:
                # EMLs split items across multiple lines, so lines aren't processed in isolation
                for line in file.readlines():
                    # log or discard the combined prior line if this line isn't a continuation of it
                    if line[0] not in (" ", "\n", "\t"):
                        if log_line:
                            df_row, recipients = process_eml_line(combined_line, df_row, recipients)
                        combined_line = ""
                        log_line = False

                    # if we have all the info that we want, we can move on from the file
                    have_info = all(c in df_row.columns for c in ("Sender", "Subject", "Date and Time"))
                    if recipients != "" and have_info:
                        break

                    # identify information which is part of the log
                    if any(line.startswith(s) for s in ("To:", "CC:", "Subject:", "Date:")):
                        log_line = True
                    elif line.startswith("From:"):
                        log_line = True
                        # if there's already been a "from" line, this is now a different email
                        # this can happen in replies where the previous email is included
                        if found_first_from_line:
                            break
                        found_first_from_line = True

                    # combine multi-line items before processing them
                    combined_line += line
        except Exception as exc:
            raise Exception(f"error thorwn while parsing {file_name}") from exc

        df_row["Recipient(s)"] = recipients.strip(", ")
        df = pd.concat((df, df_row))
        emls_logged += 1
    return df, emls_logged, non_emls

def process_eml_line(line, df_row, recipients):
    """Get data from a reconstructed line from an .eml file"""
    line = line.replace("\n", "").replace("\t", "")
    # format is "alias, <email>"; log the alias unless there isn't one
    if line.startswith("From:"):
        alias = line[5:].split("<", maxsplit=1)[0].strip('" ')
        if alias != "":
            df_row["Sender"] = alias
        else:
            df_row["Sender"] = line[5:].strip('<>" ')

    elif line.startswith("To: ") or line.startswith("CC: "):
        for recipient in line[4:].split(">, "):
            alias = recipient.split("<", maxsplit=1)[0].strip('" ')
            if alias != "":
                recipient = alias
            # quotation marks added because aliases might include commas
            recipients += f'"{recipient.strip('<" ')}", '

    elif line.startswith("Subject:"):
        df_row["Subject"] = line[8:].strip()

    elif line.startswith("Date:"):
        # format should be "day, dd mmm yyyy hh:mm:ss +0000", though it could be shorter
        date_time = line.split(", ", maxsplit=1)[-1] # -1 to avoid IndexError if day isn't listed
        # remove UTC adjustments, if there are any
        date_time = date_time.split(" +", maxsplit=1)[0].split(" -", maxsplit=1)[0].strip()
        day, month, the_rest = date_time.split(" ", maxsplit=2)
        df_row["Date and Time"] = f"{month} {day} {the_rest}"
    return df_row, recipients

def process_pdfs(pdf_dir, df):
    """Log the PDF page counts from the given PDF directory"""
    pdfs_logged = non_pdfs = 0
    for file_name in tqdm(listdir(pdf_dir)):
        if file_name[-4:] != ".pdf":
            non_pdfs += 1
            continue

        file_path = path.join(pdf_dir, file_name)
        try:
            reader = PdfReader(file_path)
            page_count = len(reader.pages)
            df.loc[df["Message No."] == file_name[:-4], "Page Count"] = page_count
            pdfs_logged += 1
        except Exception as exc:
            raise Exception(f"error thorwn while parsing {file_name}") from exc
    return df, pdfs_logged, non_pdfs

def save_xlsx(df, log_dir, have_page_count):
    """Reindex the dataframe and save it to an Excel spreadsheet"""
    columns = [
        "Message No.", "Date and Time", "Page Count", "Sender", "Recipient(s)",
        "Subject", "Exemption", "Legal Authority"
    ]
    if not have_page_count:
        columns.remove("Page Count")
    df = df.reindex(columns=columns)

    spreadsheet_path = path.join(log_dir, "Exemption Log.xlsx")
    with pd.ExcelWriter(spreadsheet_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, header=True, index=False)

def log_usage(start_time, run_time, emls_logged, pdfs_logged, eml_dir, pdf_dir, log_dir):
    """Log info about the usage of the program"""
    file_name = "usage_log.csv"
    cols_line = "Start Time,Run Time,EMLs Logged,PDFs Logged,EML Directory,PDF Directory,Log Directory\n"
    log_line = f'{start_time},{run_time},{emls_logged},{pdfs_logged},"{eml_dir}","{pdf_dir}","{log_dir}"\n'

    try:
        if path.isfile(file_name):
            with open(file_name, "a", encoding="UTF-8") as file:
                file.write(log_line)
        else:
            with open(file_name, "x", encoding="UTF-8") as file:
                file.write(cols_line + log_line)
    except PermissionError:
        print(f"Usage not logged: {file_name} is open")

if __name__ == "__main__":
    main()
