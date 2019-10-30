# General Imports
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import Label
from threading import Thread
from datetime import datetime
from time import sleep
from email.mime.text import MIMEText
import base64
import sys
import io
import os

# Google API Imports
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
import pickle
import os.path

class Config():
    def __init__(self):
        # Scopes that Google API's use to auth account.
        self.scopes = [
            'https://www.googleapis.com/auth/gmail.send', 
            'https://www.googleapis.com/auth/spreadsheets', 
            'https://www.googleapis.com/auth/drive'
        ]

        # Document list to pull email templates from. For more information, see: https://developers.google.com/docs/api/how-tos/overview#document_id
        self.doc_id_lst = [
            '1B5ksSgydbwspBA5pHjOy54KB3HK1Mj9YxEF5r7Zrsy8', 
            '13sckG4rlBeURkcKhHhpXTfV_qHnGmFKdi9KeJiv8LSc', 
            '1boQxSwFRKY6MoFyq3E230KrMxsjcTIm5AdL2s_klUxU', 
            '17gULV1EtAW5Ix-cP4c8xvM8lBBWsKKepzifbw-Jvbo0', 
            '1gKgu7ALDOOa94tt4GA5Swp1xPewPm5_4D_FKJyNHwQc'
        ]

        # Spreadsheet ID to read and write from. For more information, see: https://developers.google.com/sheets/api/guides/concepts#spreadsheet_id
        self.spr_id = '1nmHf4NLFfl9-RVlSh872_mdNgGcGDOTZdWkvyJBV6sc'

        # Times at which students should be picked up. Each infraction's time is whatever pickup time the infraction's entry time is after.
        self.pickup_times = [
            "15:10",
            "16:40"
        ]

        # The student tab to read the student data from.
        self.spr_student_data_tab = 'All Students'

        # The logging tab to write to.
        self.spr_logging_tab = 'Active Log'

        # Whether an email should be sent to parents or not [True/False]
        self.send_parent_email_bol = True

        # Email adress override. If set, all emails will go to this account instead of parents. [None, string]
        self.email_adress_override = ["fkeller20@ssis.edu.vn"]

        # Aditional Emails are email adresses to which the email is always sent to, in addition to the parents or the override.
        self.additional_emails = ["forrestkeller531@gmail.com"]

        creds = self.check_credentials()

        while True: 
            try: 
                self.spr = build('sheets', 'v4', credentials=creds) 
                break
            except: 
                network_wait()
        
        while True: 
            try: 
                self.drv = build('drive', 'v3', credentials=creds)
                break
            except: 
                network_wait()
        
        while True: 
            try: 
                self.gml = build('gmail', 'v1', credentials=creds)
                break
            except: 
                network_wait()
    
    def check_credentials(self):
        """ Returns Google OAuth Credential Object """

        creds = None

        # Check to see if credentials(token.pickle file) exists
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.scopes)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        return creds

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.config = Config()
        self.gml = self.config.gml
        self.drv = self.config.drv
        self.spr = self.config.spr
        self.spr_id = self.config.spr_id
        self.spr_student_data_tab = self.config.spr_student_data_tab

        self.master = master
        self.pack()
        self.initialize_variables()
        self.create_widgets()
    
    def initialize_variables(self):
        self.active_infractions = []

    def create_widgets(self):
        self.uid_input = ttk.Entry(self, background='red')
        self.uid_input.bind("<Return>", (lambda event: self.add_student()))
        self.uid_input.focus()

        self.submit = ttk.Button(self)
        self.submit["text"] = "Submit"
        self.submit["command"] = self.add_student

        self.signed_in_label_text = tk.StringVar(value="Signed In: 0")
        self.signed_in_label = Label(self, textvariable=self.signed_in_label_text)
        self.signed_in_label.grid(row=1, column=0)
        
        self.processing_label_text = tk.StringVar(value="Processing: 0")
        self.processing_label = Label(self, textvariable=self.processing_label_text)
        self.processing_label.grid(row=1, column=1)
        
        self.signed_out_label_text = tk.StringVar(value="Signed Out: 0")
        self.signed_out_label = Label(self, textvariable=self.signed_out_label_text)
        self.signed_out_label.grid(row=1, column=2)

        self.signed_in = tk.Listbox(self, height=30)
        self.processing = tk.Listbox(self, height=30)
        self.signed_out = tk.Listbox(self, height=30)

        self.quit = ttk.Button(self, text="Quit", command=self.master.destroy)
        self.logout = ttk.Button(self, text="Log Out", command=self.log_out)

        # Pack Elements
        self.uid_input.grid(row=0, column=0)
        self.uid_input.grid(columnspan=3)
        self.uid_input.grid(sticky="ew")

        self.submit.grid(row=0, column=2)
        self.submit.grid(sticky="e")

        self.signed_in.grid(row=2, column=0, sticky="nsew")
        self.processing.grid(row=2, column=1, sticky="nsew")
        self.signed_out.grid(row=2, column=2, sticky="nsew")

        self.quit.grid(row=3, column=2, sticky="es")
        self.logout.grid(row=3, column=0, sticky="ws")

        self.grid(padx=20, pady=20)

    def log_out(self):
        os.remove('token.pickle')
        sys.exit()

    def get_student_name_from_spreadsheet(self, uid):
        """ Gets the student name from the specified spreadsheet based off uid. Returns string of name. """
        
        request = self.spr.spreadsheets().values().get(spreadsheetId=self.spr_id, range="{}!A2:C".format(self.spr_student_data_tab))
        try:
            response = request.execute()
        except:
            messagebox.showwarning("Error", "Network Error: Please Try Again.")
            return "network-error"
        
        student_name = [value[1] + ', ' + value[2] for value in response['values'] if value[0] == uid]

        if len(student_name) > 0:
            student_name = student_name[0]
        else:
            student_name = None
            messagebox.showerror("Error", "This student is not in the spreadsheet 'All Students' tab.")

        return student_name



    def add_student(self):
        """ Adds student to appropriate listbox, removes from others if neccesary, and clears the value in uid_input. """

        self.signed_in_uid_list = self.signed_in.get(0, tk.END)
        self.processing_uid_list = self.processing.get(0, tk.END)
        self.signed_out_uid_list = self.signed_out.get(0, tk.END)

        current_value = self.uid_input.get()
        student_name = self.get_student_name_from_spreadsheet(current_value)
        if student_name == "network-error":
            return
        else:
            display_text = str(student_name) + ' - ' + current_value

        if current_value != "" and current_value.isdigit():
            if display_text in self.signed_in_uid_list:
                # Move uid to processing & Delete from signed in
                self.processing.insert(0, display_text)
                self.signed_in.delete(self.signed_in_uid_list.index(display_text))

                # Reflect change with text update to Labels
                self.processing_label_text.set("Processing: {}".format(len(self.processing_uid_list) + 1))
                self.signed_in_label_text.set("Processing: {}".format(len(self.signed_in_uid_list) - 1))

                # Start multiprocess to process student
                Thread(target=self.process_student, args=(current_value, display_text)).start()
            elif display_text in self.processing_uid_list:
                messagebox.showwarning("Error", "Wait for this student to process before adding them again.")
            else:
                # Initialize infraction
                self.signed_in.insert(0, display_text)
                self.active_infractions.append([current_value, datetime.now().strftime('%H:%M')])

                # Reflect change with text update to labels
                self.signed_in_label_text.set("Signed In: {}".format(len(self.signed_in_uid_list) + 1))

            # Clear textbox
            self.uid_input.delete(0, tk.END)

    def process_student(self, uid, display_text):
        """ Process student by writing to spreadsheet, and sending email via document template """

        # get infraction information, then remove from active
        infration_log = [arr for arr in self.active_infractions if arr[0] == uid][0]
        self.active_infractions.remove(infration_log)
        
        # Create infraction object
        infraction = Infraction(self.config, infration_log)
        error = infraction.sign_out_student()
        del infraction

        # Reflect thread completion in parent tkinter window.
        self.processing_uid_list = self.processing.get(0, tk.END) # Refresh to keep conguency with main thread.
        self.signed_out.insert(0, display_text)

        try:
            self.processing.delete(self.processing_uid_list.index(display_text))
        except:
            print("ERR-Index Deleted on Different Thread.")
        
        # Reflect completion in text Labels
        self.processing_uid_list = self.processing.get(0, tk.END)
        self.signed_out_uid_list = self.signed_out.get(0, tk.END)

        self.processing_label_text.set("Processing: {}".format(len(self.processing_uid_list)))
        self.signed_out_label_text.set("Signed Out: {}".format(len(self.signed_out_uid_list)))

        return

class Infraction():
    def __init__(self, config, infraction_log):
        # Parse google api info from config object
        config = config
        self.gml = config.gml
        self.drv = config.drv
        self.spr = config.spr
        self.spr_id = config.spr_id
        self.doc_id_lst = config.doc_id_lst
        self.spr_student_data_tab = config.spr_student_data_tab
        self.spr_logging_tab = config.spr_logging_tab
        self.send_parent_email_bol = config.send_parent_email_bol
        self.email_adress_override = config.email_adress_override
        self.additional_emails = config.additional_emails
        self.pickup_times = config.pickup_times

        # Get variables from arguments
        self.infraction_log = infraction_log
        self.student_id = infraction_log[0]
        self.sign_in_time = infraction_log[1]

        print("{}: Initializing...".format(self.student_id))

        # Read spreadsheet and parse relevant information.
        spr_read = self.read_spreadsheet()

        # From whole spreadsheet read, extract the values from logging and data tabs.
        for response in spr_read:
            response_range = response['range']
            response_values = response['values']
            if self.spr_logging_tab in response_range:
                self.spr_logging_tab_read = response_values
            elif self.spr_student_data_tab in response_range:
                self.spr_student_data_tab_read = response_values
        
        # Get infraction information for student
        self.student_infractions = [infraction for infraction in self.spr_logging_tab_read if len(infraction) > 0 and infraction[0] == self.student_id]

        # From student data, extract current students information and header
        header_read = self.spr_student_data_tab_read[0]
        student_information_read = [row for row in self.spr_student_data_tab_read if len(row) !=0 and row[0] == self.student_id]

        # If the student is in the spreadsheet, collect data into dict.
        if len(student_information_read) > 0:
            student_information_read = student_information_read[0]
            self.student_in_spreadsheet = True
            self.student_information = {}

            for i in range(len(header_read)):
                key = header_read[i]
                value = student_information_read[i]
                self.student_information[key] = value
        else:
            self.student_in_spreadsheet = False
            self.student_information = {}

            for i in range(len(header_read)):
                key = header_read[i]
                value = ""
                self.student_information[key] = value


    def extend_infraction_log(self):
        """ Completes the information for the infraction log to be added to the spreadsheet """
        
        datetime_query = datetime.now()

        # This is where any further information requests should be added. ERROR will always be at the end.
        self.infraction_log.extend([
            datetime_query.strftime('%H:%M'), 
            datetime_query.strftime('%y-%m-%d'), 
            self.student_information["LAST_NAME"], 
            self.student_information["FIRST_NAME"]
        ])

        if not self.student_in_spreadsheet:
            self.infraction_log.append("Error: Student not in spreadsheet.")
    
    def log_student_infraction(self):
        """ Logs the student infraction in the spreadsheet. """
    
        # Get current active tab sheet_id
        request = self.spr.spreadsheets().get(spreadsheetId=self.spr_id)
        while True:
            try:
                response = request.execute()
                break
            except:
                self.network_wait()

        sheet_information = response['sheets']

        for sheet in sheet_information:
            if sheet['properties']['title'] == self.spr_logging_tab:
                spr_logging_tab_sheet_id = sheet['properties']['sheetId']
        
        # Build request body
        body = {
            "requests": [
                {
                "insertRange": {
                    "range": {
                    "sheetId": spr_logging_tab_sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2
                },
                    "shiftDimension": "ROWS"
            }
            },
            {
            "pasteData": {
                "data": ','.join(str(value) for value in self.infraction_log),
                "type": "PASTE_NORMAL",
                "delimiter": ",",
                "coordinate": {
                "sheetId": spr_logging_tab_sheet_id,
                "rowIndex": 1,
                }
            }
            }
            ]
        }
        
        # run batch update
        request = self.spr.spreadsheets().batchUpdate(spreadsheetId=self.spr_id, body=body)
        
        while True:
            try:
                response = request.execute()
                return
            except:
                self.network_wait()
    
    def download_email_template(self, file_id):
        """ Downloads specified email template, returns string of text in the downwloaded file. """

        while True: 
            try:
                # Download the file using Drive API into a file file_path
                file_path = "{}.txt".format(file_id)
                request = self.drv.files().export_media(fileId=file_id, mimeType='text/plain')
                fh = io.FileIO(file_path, "wb")
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()

                # Extract text from file_path file
                with open(file_path, 'r') as file:
                    data = file.read()
                
                os.remove(file_path)

                return data
            except:
                print("{}: Error downloading template, retrying...".format(self.student_id))
    
    def network_wait(self):
        """ Waits 10 seconds, then continues """

        print("{}: Network Error: Waiting 10 seconds to retry...".format(self.student_id))
        sleep(10)

    def read_spreadsheet(self):
        """ Returns JSON API response of all information in logging and student data tab. """

        ranges = [self.spr_logging_tab, self.spr_student_data_tab]
        value_render_option = 'FORMATTED_VALUE'
        request = self.spr.spreadsheets().values().batchGet(
            spreadsheetId=self.spr_id, 
            ranges=ranges, 
            valueRenderOption=value_render_option,
        )
        while True:
            try:
                response = request.execute()
                break
            except:
                self.network_wait()
        
        return response['valueRanges']

    def array_to_html_table(self, data):
        """ Function modified from: https://gist.github.com/aizquier/ef229826754626ffc4b8b6e49c599b68 """
        
        q = "<table border='1'>"
        for i in [(data[0:1], 'th'), (data[1:], 'td')]:
            q += "".join(
                [
                    "<tr>%s</tr>" % str(_mm) 
                    for _mm in [
                        "".join(
                            [
                                "<%s>%s</%s>" % (i[1], str(_q), i[1]) 
                                for _q in _m
                            ]
                        ) for _m in i[0]
                    ]
                ])
        q += "</table>"
        return q

    def calculate_wait_time(self, entry_time, exit_time):
        """ Returns the time the student waited. """

        infraction_entry_time_minutes = self.time_to_minutes(entry_time)
        infraction_exit_time_minutes = self.time_to_minutes(exit_time)
        
        return infraction_exit_time_minutes - infraction_entry_time_minutes

    def calculate_expected_time(self, entry_time):
        """ Calculates the expected time the student should have been picked up at. Returns time string """

        infraction_expected_time = "undef"
        for pickup_time in self.pickup_times:
            if self.time_to_minutes(entry_time) > self.time_to_minutes(pickup_time):
                infraction_expected_time = pickup_time
    
        return infraction_expected_time

    def time_to_minutes(self, time):
        """ Converts time to minutes ex: 1:15 => 75 """

        return int(time.split(":")[0]) * 60 + int(time.split(":")[1])
    
    def generate_table(self):
        """ Returns an HTML table with the correct information for the student infractions """

        table_headers = ["Incident", "Expected pick-up time", "Time brought to office", "Time picked up from office", "Minutes late/ child waiting", "School Action"]
        school_actions = [
            "Reminder email",
            "Email warning",
            "Meeting with Assistant Principal",
            "Meeting with Principal",
            "Meeting with Head of School"
        ]
        rowIDs = range(1, len(self.student_infractions) + 1)

        table_data = [table_headers]
        current_infractions = self.student_infractions
        current_infractions.insert(0, self.infraction_log)
        current_infractions.reverse()
        
        for i in range(len(current_infractions)):
            infraction = current_infractions[i]
            infraction_entry_time = infraction[1]
            infraction_exit_time = infraction[2]
            infraction_wait_time = self.calculate_wait_time(infraction_entry_time, infraction_exit_time)
            infraction_expected_time = self.calculate_expected_time(infraction_entry_time)

            # Find School Action
            if i >= 4:
                school_action = school_actions[4]
            else:
                school_action = school_actions[i]

            table_data.append([i + 1, infraction_expected_time, infraction_entry_time, infraction_exit_time, infraction_wait_time, school_action])

        table = self.array_to_html_table(table_data)

        return table

    def send_parent_email(self):
        """ Sends custom email taken from document template to parents. """

        student_infraction_count = len(self.student_infractions)

        # Get email template from google doc
        if student_infraction_count >= 5:
            doc_index = 4
        else:
            doc_index = student_infraction_count
        
        file_id = self.doc_id_lst[doc_index]

        email_template_text = self.download_email_template(file_id)
        
        while True:
            try:
                email_template_name = self.drv.files().get(fileId=file_id).execute()['name']
                break
            except:
                self.network_wait()
        
        # Extract student information
        student_name = self.student_information["FIRST_NAME"]
        if self.email_adress_override == None:
            email_adresses = [self.student_information["FATHER'S EMAIL ADDRESS"], self.student_information["MOTHER'S EMAIL ADDRESS"]]
            print("{}: Sending email to parent adresses: {}".format(self.student_id, email_adresses))
        else:
            email_adresses = self.email_adress_override
            print("{}: Override set... Overiding with: {}".format(self.student_id, self.email_adress_override))

        # Add additional_emails
        if self.additional_emails != None:
            email_adresses.extend(self.additional_emails)
            print("{}: Additional smails set... also sending email to: {}".format(self.student_id, self.additional_emails))

        if student_infraction_count >= 1:
            table = self.generate_table()
        else:
            entry_time = self.infraction_log[1]
            exit_time = self.infraction_log[2]
            wait_time = self.calculate_wait_time(entry_time, exit_time)
            expected_time = self.calculate_expected_time(entry_time)
        
        email_text = eval(f'f"""{email_template_text}"""')

        # Send email via Gmail API
        message = MIMEText(email_text.replace('\n', '<br />'), "html")
        message['to'] = ','.join(email_adresses)
        message['subject'] = email_template_name

        raw = base64.urlsafe_b64encode(message.as_bytes())
        raw = raw.decode()
        message =  {'raw': raw}

        return self.gml.users().messages().send(userId="me", body=message).execute()

    def sign_out_student(self):
        self.extend_infraction_log()
        print("{}: Logging infraction to spreadsheet...".format(self.student_id))
        self.log_student_infraction()

        if self.student_in_spreadsheet:
            error = False
            if self.send_parent_email_bol:
                print("{}: Sending email...".format(self.student_id))
                self.send_parent_email()
            else:
                print("{}: Sending email set to FALSE, no email sent...".format(self.student_id))
        else:
            error = True
        
        print("{}: Successfully completed.".format(self.student_id))
        
        return error
    
def network_wait():
    """ Waits 10 seconds, then returns True. """

    print("Network Error: Waiting 10 seconds to retry...")
    sleep(10)
    return True

if __name__ == "__main__":
    # Set up root tkinter window
    root = tk.Tk()
    root.title("Late Student Check Out")

    app = Application(master=root)
    app.mainloop()