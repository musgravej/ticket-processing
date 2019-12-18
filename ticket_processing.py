import sqlite3
import os
import configparser
import datetime
import re
import PyPDF2
import shutil
import win32api
import win32print
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib


class Email:
    def __init__(self):
        config = configparser.ConfigParser()
        config.read(os.path.join(os.curdir, 'config.ini'))

        self.email_to = config['EMAIL']['email_to']
        self.email_cc = config['EMAIL']['email_cc']
        self.email_alert = config['EMAIL']['email_alert']
        self.email_from = config['EMAIL']['email_from']
        self.email_user = config['EMAIL']['email_user']
        self.email_password = config['EMAIL']['email_password']
        self.email_server = config['EMAIL']['email_server']

    def send_alert_email(self):
        port = 587
        smtp_server = self.email_server
        sender_email = self.email_user
        email_from = self.email_from
        receiver_email = self.email_alert
        password = self.email_password

        pdt = datetime.datetime.strftime(gbl.process_dt, "%Y-%m-%d")

        html_head = ("<html> <head> <style> td, th { border: 1px solid #dddddd; "
                     "text-align: left; padding: 8px;}</style> </head> <body> ")
        html_foot = "</body> </html>"
        err_1_html = ""
        err_2_html = ""

        if fmv.error_messages:
            table_data = ""
            fle_format = "error has" if len(fmv.error_messages) == 1 else "errors have"

            for msg in fmv.error_messages:
                table_data += f"<tr><td>{msg}</td></tr>"

            err_1_html = ("<p>The following {fle_format} occurred while processing web orders:<br> "
                          "</p> <table width: 100%;> <tr> <th>Error</th> </tr> "
                          "{table_data} </table>".format(table_data=table_data, fle_format=fle_format))

        if gbl.duplicated_files:
            table_data = ""
            fle_format = "file has" if len(gbl.duplicated_files) == 1 else "files have"

            for file_name, file_date in gbl.duplicated_files:
                table_data += f"<tr><td>{file_name}</td><td>{file_date}</td></tr>"

            err_2_html = ("<p>The following {fle_format} been <em><b>updated</b></em> for web orders:<br> "
                          "</p> <table width: 100%;> <tr> <th>File Name</th> <th>File Date</th> </tr> "
                          "{table_data} </table>".format(table_data=table_data, fle_format=fle_format))

        html = f"{html_head}{err_1_html}{err_2_html}{html_foot}"
        subject = f"Web Order ERRORS for processing {pdt}"

        text = ""
        message = MIMEMultipart("alternative")

        message["Subject"] = subject
        message["From"] = sender_email
        message["To"] = receiver_email

        message.attach(MIMEText(text, "plain"))
        message.attach(MIMEText(html, "html"))

        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(email_from, message["To"].split(","),
                            message.as_string())

    def send_message_email(self):
        port = 587
        smtp_server = self.email_server
        sender_email = self.email_user
        email_from = self.email_from
        password = self.email_password
        receiver_email = self.email_to
        cc_email = self.email_cc

        pdt = datetime.datetime.strftime(gbl.process_dt, "%Y-%m-%d")

        table_data = ""
        fle_format = "file has" if len(rpt.portal_counts) == 1 else "files have"

        for filename, records in rpt.portal_counts.items():
            table_data += f"<tr><td>{filename}</td><td>{records}</td></tr>"

        html = ("<html> <head> <style> td, th {{ border: 1px solid #dddddd; text-align: left; padding: 8px;}}"
                "</style> </head> <body> <p>The following portals {fle_format} been processed for web orders:<br> "
                "</p> <table width: 100%;> <tr> <th>Portal</th> <th>Orders</th> </tr> "
                "{table_data} </table> </body> </html>".format(table_data=table_data, fle_format=fle_format))

        subject = f"Web Order Summary for processing {pdt}"

        text = ""
        message = MIMEMultipart("alternative")

        search = re.compile("DailyReportOfOrders_[\s\S]*.pdf")
        daily_reports = list(filter(search.match, gbl.report_files))
        for report in daily_reports:
            with open(os.path.join(gbl.processing_directory, report), "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)

            part.add_header("Content-Disposition", f"attachment; filename= {report}",)
            message.attach(part)

        message["Subject"] = subject
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Cc"] = cc_email

        message.attach(MIMEText(text, "plain"))
        message.attach(MIMEText(html, "html"))

        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(email_from, message["To"].split(",") + message["Cc"].split(","),
                            message.as_string())


class FilePrinter:
    def __init__(self):
        self.packing_slips = []
        self.pull_packing_slips = []
        self.work_orders = []
        self.pull_work_orders = []
        self.print_on_demand = []
        self.pull_print_on_demand = []
        self.daily_report = []
        self.pull_daily_report = []
        self.default_reports = 'Print Reports'
        self.default_printer = win32print.GetDefaultPrinter()

    def set_report_printer(self):
        win32print.SetDefaultPrinter(self.default_reports)

    def set_default_printer(self):
        win32print.SetDefaultPrinter(self.default_printer)

    def print_processing_reports(self):
        conn = sqlite3.connect(gbl.db)
        results = conn.execute("SELECT DATE(a.file_date) FROM `processing` a "
                               "WHERE NOT EXISTS(SELECT * FROM history b "
                               "WHERE b.file_name = a.file_name AND b.file_date = a.file_date) "
                               "GROUP BY DATE(a.file_date) ORDER BY a.file_date;")
        days = [r[0] for r in results.fetchall()]
        conn.close()

        for day in days:
            self.update_lists(day)
            self.print_reports()

    def update_lists(self, proc_date):
        conn = sqlite3.connect(gbl.db)
        conn.create_function('REGEXP', 2, lambda x, y: 1 if re.match(x, y) else 0)
        cursor = conn.cursor()

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name REGEXP "[0-9]{5}_PS_[0-9]{8}.pdf" ORDER BY a.file_date;')

        self.packing_slips = [r[1] for r in results.fetchall()]

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name REGEXP "[\d]{5}_PS_split_[\d]{8}.pdf" ORDER BY a.file_date;')

        self.pull_packing_slips = [r[1] for r in results.fetchall()]

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name REGEXP "[\d]{5}_WO_[\d]{8}.pdf" ORDER BY a.file_date;')

        self.work_orders = [r[1] for r in results.fetchall()]

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name REGEXP "[\d]{5}_WO_split_[\d]{8}.pdf" ORDER BY a.file_date;')

        self.pull_work_orders = [r[1] for r in results.fetchall()]

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name REGEXP "WM_POD_[\d]{8}.pdf" ORDER BY a.file_date;')

        self.print_on_demand = [r[1] for r in results.fetchall()]

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name REGEXP "WM_POD_split_[\d]{8}.pdf" ORDER BY a.file_date;')

        self.pull_print_on_demand = [r[1] for r in results.fetchall()]

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name REGEXP "DailyReportOfOrders_[\d]{8}.pdf" ORDER BY a.file_date;')

        self.daily_report = [r[1] for r in results.fetchall()]

        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 f"AND DATE(a.file_date) = '{proc_date}' AND "
                                 ' a.file_name '
                                 'REGEXP "DailyReportOfOrders_split_[\d]{8}.pdf" ORDER BY a.file_date;')

        self.pull_daily_report = [r[1] for r in results.fetchall()]

        conn.close()

    def print_reports(self):
        portal_codes = ['18241', '19404', '20403', '23005', '23396', '23798', '23640']

        # run through queue in order, iterate through codes
        for code in portal_codes:

            print_queue = list(filter(re.compile(f"{code}_[\s\S]*.pdf").match, self.packing_slips))
            if print_queue:
                for print_file in print_queue:
                    file_name = os.path.join(gbl.processing_directory, print_file)
                    print(file_name)
                    win32api.ShellExecute(0, "print", file_name, None, ".", 0)

            print_queue = list(filter(re.compile(f"{code}_[\s\S]*.pdf").match, self.pull_packing_slips))
            if print_queue:
                for print_file in print_queue:
                    file_name = os.path.join(gbl.processing_directory, print_file)
                    print(file_name)
                    win32api.ShellExecute(0, "print", file_name, None, ".", 0)

            print_queue = list(filter(re.compile(f"{code}_[\s\S]*.pdf").match, self.work_orders))
            if print_queue:
                for print_file in print_queue:
                    file_name = os.path.join(gbl.processing_directory, print_file)
                    print(file_name)
                    win32api.ShellExecute(0, "print", file_name, None, ".", 0)

            print_queue = list(filter(re.compile(f"{code}_[\s\S]*.pdf").match, self.pull_work_orders))
            if print_queue:
                for print_file in print_queue:
                    file_name = os.path.join(gbl.processing_directory, print_file)
                    print(file_name)
                    if code in ('19404', '20403'):
                        win32api.ShellExecute(0, "print", "slip.pdf", None, ".", 0)
                        win32api.ShellExecute(0, "print", file_name, None, ".", 0)
                    else:
                        win32api.ShellExecute(0, "print", file_name, None, ".", 0)

        print_queue = self.print_on_demand
        if print_queue:
            for print_file in print_queue:
                file_name = os.path.join(gbl.processing_directory, print_file)
                print(file_name)
                win32api.ShellExecute(0, "print", file_name, None, ".", 0)

        print_queue = self.pull_print_on_demand
        if print_queue:
            for print_file in print_queue:
                file_name = os.path.join(gbl.processing_directory, print_file)
                print(file_name)
                win32api.ShellExecute(0, "print", file_name, None, ".", 0)


class FileMover:
    def __init__(self):
        self.save_base_path = os.path.join("\\\\JTSRV4", "Data", "Customer Files",
                                           "In Progress", "01-Web Order Art")

        # self.save_base_path = os.path.join(os.path.curdir, 'art_path')
        self.error_messages = set()

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

    def move_farm_bureau_art(self):
        dt = datetime.datetime
        save_path = os.path.join(self.save_base_path, "FB Monthly Web Order")

        file_search = re.compile("FB[\s\S]*.pdf")
        file_match = (list(filter(file_search.match, gbl.art_files)))

        for fle in file_match:
            fle_create_date = dt.strftime(dt.fromtimestamp(os.path.getmtime(
                    os.path.join(gbl.processing_directory, fle))), "%Y%m%d")

            try:
                fle_job_no = self.config['farmbureau'][fle_create_date[:6]]
                job_dir = [f for f in os.listdir(save_path) if f[:5] == fle_job_no]
                if not job_dir:
                    self.error_messages.add(f"No Farm Bureau job folder for {fle_create_date[:6]}")
                    full_save_path = os.path.join(save_path, 'FB_No_Job_Folder')
                    os.makedirs(full_save_path, exist_ok=True)
                else:
                    full_save_path = os.path.join(save_path, job_dir[0])

                self.move_file_and_split(fle, full_save_path)

            except KeyError:
                self.error_messages.add(f"No Farm Bureau job number for {fle_create_date[:6]}")
                fle_job_no = "FB_No_Job_Number"
                full_save_path = os.path.join(save_path, fle_job_no)
                os.makedirs(full_save_path, exist_ok=True)
                self.move_file_and_split(fle, full_save_path)

    def move_willis_art(self):
        dt = datetime.datetime
        save_path = os.path.join(self.save_base_path, "Willis Auto Web Orders")

        file_search = re.compile("WAG[\s\S]*.pdf")
        file_match = (list(filter(file_search.match, gbl.art_files)))

        for fle in file_match:
            fle_create_date = dt.strftime(dt.fromtimestamp(os.path.getmtime(
                    os.path.join(gbl.processing_directory, fle))), "%Y%m%d")

            try:
                fle_job_no = self.config['willis'][fle_create_date[:6]]
                job_dir = [f for f in os.listdir(save_path) if f[:5] == fle_job_no]
                if not job_dir:
                    self.error_messages.add(f"No Willis job folder for {fle_create_date[:6]}")
                    full_save_path = os.path.join(save_path, 'WAG_No_Job_Folder')
                    os.makedirs(full_save_path, exist_ok=True)
                else:
                    full_save_path = os.path.join(save_path, job_dir[0])

                self.move_file_and_split(fle, full_save_path)

            except KeyError:
                self.error_messages.add(f"No Willis job number for {fle_create_date[:6]}")
                fle_job_no = "WAG_No_Job_Number"
                full_save_path = os.path.join(save_path, fle_job_no)
                os.makedirs(full_save_path, exist_ok=True)
                self.move_file_and_split(fle, full_save_path)

    def move_medica_art(self):
        dt = datetime.datetime
        save_path = os.path.join(self.save_base_path, "Medica Monthly Web Orders")

        file_search = re.compile("MMH[\s\S]*.pdf")
        file_match = (list(filter(file_search.match, gbl.art_files)))

        for fle in file_match:
            fle_create_date = dt.strftime(dt.fromtimestamp(os.path.getmtime(
                    os.path.join(gbl.processing_directory, fle))), "%Y%m%d")

            try:
                fle_job_no = self.config['medica'][fle_create_date[:6]]
                job_dir = [f for f in os.listdir(save_path) if f[:5] == fle_job_no]
                if not job_dir:
                    self.error_messages.add(f"No Medica job folder for {fle_create_date[:6]}")
                    full_save_path = os.path.join(save_path, 'MMH_No_Job_Folder')
                    os.makedirs(full_save_path, exist_ok=True)
                else:
                    full_save_path = os.path.join(save_path, job_dir[0])

                self.move_file_and_split(fle, full_save_path)

            except KeyError:
                self.error_messages.add(f"No Medica job number for {fle_create_date[:6]}")
                fle_job_no = "MMH_No_Job_Number"
                full_save_path = os.path.join(save_path, fle_job_no)
                os.makedirs(full_save_path, exist_ok=True)
                self.move_file_and_split(fle, full_save_path)

    def move_waukee_art(self):
        dt = datetime.datetime
        save_path = os.path.join(self.save_base_path, "City of Waukee Web Orders")

        file_search = re.compile("CW[\s\S]*.pdf")
        file_match = (list(filter(file_search.match, gbl.art_files)))

        for fle in file_match:
            fle_create_date = dt.strftime(dt.fromtimestamp(os.path.getmtime(
                    os.path.join(gbl.processing_directory, fle))), "%Y%m%d")

            try:
                fle_job_no = self.config['waukee'][fle_create_date[:6]]
                job_dir = [f for f in os.listdir(save_path) if f[:5] == fle_job_no]
                if not job_dir:
                    self.error_messages.add(f"No City of Waukee job folder for {fle_create_date[:6]}")
                    full_save_path = os.path.join(save_path, 'CW_No_Job_Folder')
                    os.makedirs(full_save_path, exist_ok=True)
                else:
                    full_save_path = os.path.join(save_path, job_dir[0])

                self.move_file_and_split(fle, full_save_path)

            except KeyError:
                self.error_messages.add(f"No City of Waukee job number for {fle_create_date[:6]}")
                fle_job_no = "CW_No_Job_Number"
                full_save_path = os.path.join(save_path, fle_job_no)
                os.makedirs(full_save_path, exist_ok=True)
                self.move_file_and_split(fle, full_save_path)

    def move_file_and_split(self, file, path):
        new_dir = file.split('_')[0]
        if not os.path.isdir(os.path.join(path, new_dir)):
            os.makedirs(os.path.join(path, new_dir))

        print(f"Copying file: {file}")
        # creates a temporary file as 'w-*'
        shutil.copy(os.path.join(gbl.processing_directory, file),
                    os.path.join(path, new_dir, "w-{0}".format(file)))

        w_pdf = PyPDF2.PdfFileReader(os.path.join(path, new_dir, "w-{0}".format(file)))
        s_pdf = PyPDF2.PdfFileWriter()
        data_sheet = PyPDF2.PdfFileWriter()

        for page in range(w_pdf.getNumPages()):
            if page == 0:
                data_sheet.addPage(w_pdf.getPage(page))
            else:
                s_pdf.addPage(w_pdf.getPage(page))

        with open(os.path.join(path, new_dir, file), 'wb') as s:
            s_pdf.write(s)

        with open(os.path.join(path, new_dir, "{0}-data sheet.pdf".format(file[:-4])), 'wb') as s:
            data_sheet.write(s)

        os.remove(os.path.join(path, new_dir, "w-{0}".format(file)))


class ReportCounts:
    def __init__(self):
        self.portal_counts = dict()

    def add_report_count(self, portal, count):
        self.portal_counts[portal] = self.portal_counts.get(portal, 0) + count

    def portal_count_message(self):
        count_string = ""
        for key, value in self.portal_counts.items():
            count_string += "{0}: {1}\r\n".format(key, value)

        return count_string

    def lite_portal_counts(self, pdf_path):
        """Reads pdf, gets portal counts, returns string of portal counts"""
        pdf = PyPDF2.PdfFileReader(pdf_path)
        for page in range(pdf.getNumPages()):
            portal = pdf.getPage(page).extractText().split('\n')[0]
            self.add_report_count(portal, 1)

    def get_report_counts(self):
        work_order_search = re.compile("[\d]*_WO_[\d]*.pdf")
        kit_search = re.compile("[\d]*_WO_split_[\d]*.pdf")

        work_orders = set((list(filter(work_order_search.match, gbl.report_files))))
        kit_orders = set((list(filter(kit_search.match, gbl.report_files))))

        for order in work_orders:
            if order[:5] == '20403':
                cnt = self.count_pdf_pages(os.path.join(gbl.processing_directory, order))
                self.add_report_count("Wellmark", cnt)

            if order[:5] == '19404':
                cnt = self.count_pdf_pages(os.path.join(gbl.processing_directory, order))
                self.add_report_count("Farm Bureau", cnt)

            if order[:5] == '23396':
                cnt = self.count_pdf_pages(os.path.join(gbl.processing_directory, order))
                self.add_report_count("Medica", cnt)

            if order[:5] == '18241':
                self.lite_portal_counts(os.path.join(gbl.processing_directory, order))

        for order in kit_orders:
            if order[:5] == '20403':
                cnt = self.count_pdf_pages(os.path.join(gbl.processing_directory, order))
                self.add_report_count("Wellmark Pull From Inventory", cnt)

            if order[:5] == '19404':
                cnt = self.count_pdf_pages(os.path.join(gbl.processing_directory, order))
                self.add_report_count("Farm Bureau Pull From Inventory", cnt)

            if order[:5] == '23396':
                cnt = self.count_pdf_pages(os.path.join(gbl.processing_directory, order))
                self.add_report_count("Medica Pull From Inventory", cnt)

    def count_pdf_pages(self, pdf_path):
        """Counts the number of pages in pdf_path (full pdf path and file name)"""
        pdf = PyPDF2.PdfFileReader(pdf_path)
        return pdf.getNumPages()


class GlobalVar:
    def __init__(self):
        self.db = 'ticket_processing.db'
        self.processing_directory = os.path.join("\\\\JTSRV3", "Print Facility", "Job Ticket Feed docs",
                                                 "WebToPrint")
        # self.processing_directory = os.path.join(os.curdir, "target")
        self.target_directory = None
        self.processing_files = None
        self.process_history = None
        self.new_files = None
        self.report_files = None
        self.art_files = None
        self.duplicated_files = []
        self.process_dt = datetime.datetime.now()

    def initialize_processing(self):
        self.connect_db()
        self.get_files_in_process()
        self.update_table_processing()
        self.get_process_history()
        self.get_new_files()
        self.get_report_files()
        self.get_art_files()
        self.get_duplicated_files()

    def set_target_directory(self):
        """Currently Unused, for moving to dated folders"""
        file_day = datetime.datetime.strftime(self.process_dt, "%d")
        # file_month = datetime.datetime.strftime(self.process_dt, "%m")
        file_month_name = datetime.datetime.strftime(self.process_dt, "%b")
        file_year = datetime.datetime.strftime(self.process_dt, "%Y")
        target_dir = os.path.dirname(os.path.realpath(__file__))

        self.target_directory = os.path.join(target_dir, f"{file_year}", f"{file_month_name}", f"{file_day}")
        os.makedirs(self.target_directory, exist_ok=True)

    def get_files_in_process(self):
        self.processing_files = set(f for f in os.listdir(self.processing_directory) if f.upper()[-3:] == 'PDF')

    def get_new_files(self):
        conn = sqlite3.connect(self.db)
        results = conn.execute("SELECT a.file_name FROM `processing` a "
                               "WHERE NOT EXISTS(SELECT * FROM `history` b "
                               "WHERE b.file_name = a.file_name "
                               "AND b.file_date = a.file_date);")
        self.new_files = set(r[0] for r in results.fetchall())
        conn.close()

    def get_process_history(self):
        conn = sqlite3.connect(self.db)
        results = conn.execute("SELECT `file_name` from `history`;")
        self.process_history = set(r[0] for r in results.fetchall())
        conn.close()

    def get_art_files(self):
        """Sets self.art_files to a set of artwork files for processing date"""
        conn = sqlite3.connect(self.db)
        conn.create_function('REGEXP', 2, lambda x, y: 1 if re.search(x, y) else 0)
        cursor = conn.cursor()
        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 'AND NOT a.file_name REGEXP "[\w]*[\d]{8}.pdf" ORDER BY a.file_date;')

        self.art_files = [r[1] for r in results.fetchall()]
        conn.close()

    def get_report_files(self):
        """Sets self.report_files to a set of packing slips, work orders, daily reports for processing date"""
        conn = sqlite3.connect(self.db)
        conn.create_function('REGEXP', 2, lambda x, y: 1 if re.search(x, y) else 0)
        cursor = conn.cursor()
        results = cursor.execute('SELECT a.* FROM `processing` a '
                                 'WHERE NOT EXISTS(SELECT * FROM history b '
                                 'WHERE b.file_name = a.file_name AND b.file_date = a.file_date) '
                                 'AND a.file_name REGEXP "[\w]*[\d]{8}.pdf" ORDER BY a.file_date;')

        self.report_files = [r[1] for r in results.fetchall()]
        conn.close()

    def connect_db(self):
        # if not os.path.exists(self.db):
        conn = sqlite3.connect(self.db)

        # # Only run when initializing history table
        # print("Building history table")
        # conn.execute("DROP TABLE IF EXISTS `history`;")
        # conn.execute("CREATE TABLE `history` (`process_datetime` DATETIME, "
        #              "`file_name` VARCHAR(100), `file_date` DATETIME);")

        print("Building processing table")
        conn.execute("DROP TABLE IF EXISTS `processing`;")
        conn.execute("VACUUM;")
        conn.execute("CREATE TABLE `processing` (`process_datetime` DATETIME, "
                     "`file_name` VARCHAR(100), `file_date` DATETIME);")
        conn.commit()
        conn.close()

    def update_table_history(self):
        print("Updating history table")
        conn = sqlite3.connect(self.db)

        sql = ("INSERT INTO `history` SELECT * FROM `processing` a " 
               "WHERE NOT EXISTS(SELECT * FROM history b WHERE "
               "b.file_name = a.file_name AND b.file_date = a.file_date);")

        conn.execute(sql)
        conn.commit()
        conn.close()

    def get_duplicated_files(self):
        conn = sqlite3.connect(self.db)
        sql = ("SELECT c.* FROM (SELECT a.* FROM `processing` a "
               "WHERE NOT EXISTS(SELECT * FROM history b "
               "WHERE b.file_name = a.file_name AND b.file_date = a.file_date) ) c "
               "WHERE EXISTS (SELECT * FROM history d WHERE c.file_name = d.file_name);")
        results = conn.execute(sql)
        self.duplicated_files = [(r[1], r[2]) for r in results.fetchall()]
        conn.close()

    def update_table_processing(self):
        print("Updating processing table")
        conn = sqlite3.connect(self.db)
        for f in self.processing_files:
            dt = datetime.datetime
            file_datetime = dt.strftime(dt.fromtimestamp(os.path.getmtime(
                    os.path.join(self.processing_directory, f))), "%Y-%m-%d %H:%M:%S")

            sql = "INSERT INTO `processing` VALUES (?, ?, ?);"
            conn.execute(sql, (gbl.process_dt, f, file_datetime,))
        conn.commit()
        conn.close()


def main():
    gbl.initialize_processing()
    fpr.set_report_printer()
    rpt.get_report_counts()
    fmv.move_farm_bureau_art()
    fmv.move_medica_art()
    fmv.move_willis_art()
    fmv.move_waukee_art()

    fpr.print_processing_reports()

    # send email if new files
    if len(gbl.report_files) > 0:
        eml.send_message_email()

    # send email if errors
    if (len(fmv.error_messages) > 0) or (len(gbl.duplicated_files) > 0):
        eml.send_alert_email()

    gbl.update_table_history()
    fpr.set_default_printer()


if __name__ == '__main__':
    gbl = GlobalVar()
    fpr = FilePrinter()
    rpt = ReportCounts()
    fmv = FileMover()
    eml = Email()

    main()
