import xlrd
from win32com.client import Dispatch
import win32com.client
import re
import time
import datetime
import calendar
import shutil
import os


class get_file(object):
    def __init__(self):
        self.raw_data_route = os.getcwd() + "\RawData\\"
        self.templates_route = os.getcwd() + "\Templates\\"
        self.raw_dirs = os.listdir(self.raw_data_route)
        self.template_dirs = os.listdir(self.templates_route)
        self.certificate_file = " "
        self.incident_file = " "
        self.account_raw_file = " "
        self.msr_report = " "
        self.incentive_report = " "

    def specify_file(self):
        for file in self.raw_dirs:
            if "New Certificates" in file:
                self.certificate_file = self.raw_data_route + file
            elif "security incident" in file:
                self.incident_file = self.raw_data_route + file
            elif "Create account" in file:
                self.account_raw_file = self.raw_data_route + file
        for file in self.template_dirs:
            if "MSR" in file:
                self.msr_report = self.templates_route + file
            elif "Incentive" in file:
                self.incentive_report = self.templates_route + file

    def logger(self):
        return "The files read this time are: \nCertificate_file:  %s \nIncident_file:  %s \nAccount_raw_file:  %s \n" \
               "MSR_report:  %s \nIncentive_report:  %s\n\n" % (self.certificate_file, self.incident_file,
                                                                self.account_raw_file,
                                                                self.msr_report, self.incentive_report)


class certificate(object):
    def __init__(self, data_file):
        self.certificate_file_route = data_file
        self.total_certificates = 0
        self.miss_sla = 0
        self.flag = "Miss SLA of Certificates' ID: \n"

    def read_certificate(self):
        book = xlrd.open_workbook(self.certificate_file_route)
        sheet = book.sheet_by_index(0)
        total_row = sheet.nrows
        for i in range(0, total_row - 1):
            self.total_certificates += int(sheet.cell(i + 1, 8).value)
            create_date = xlrd.xldate.xldate_as_datetime(sheet.cell(i + 1, 6).value, 0)
            finish_date = xlrd.xldate.xldate_as_datetime(sheet.cell(i + 1, 7).value, 0)
            if (finish_date - create_date).days > 2 and (sheet.cell(i + 1, 10).value == ""):
                self.miss_sla += 1
                self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"

    def logger(self):
        return "--------------------------------\nCertificates statistics: \nTotal_Certificates:  %s, Miss_SLA of Certificates:  %s\n\n" % (
        self.total_certificates, self.miss_sla)


class account(object):
    def __init__(self, raw_file):
        self.account_raw_data_route = raw_file
        self.account_data = {
            'user_account_21v_creation': 0,
            'miss_sla_user_account_21v_creation': 0,
            'user_account_21v_modification': 0,
            'miss_sla_user_account_21v_modification': 0,
            'user_account_21v_termination': 0,
            'miss_sla_user_account_21v_termination': 0,
            'user_account_ms_creation': 0,
            'miss_sla_user_account_ms_creation': 0,
            'user_account_ms_modification': 0,
            'miss_sla_user_account_ms_modification': 0,
            'user_account_ms_termination': 0,
            'miss_sla_user_account_ms_termination': 0,
            'security_group': 0,
            'miss_sla_security_group': 0
        }
        self.flag = "Miss SLA of CME Accounts' Ticket ID: \n"

    def read_account_raw(self):
        book = xlrd.open_workbook(self.account_raw_data_route)
        ex_year = time.localtime().tm_year
        ex_month_eng = calendar.month_abbr[time.localtime().tm_mon - 1]
        sheet = book.sheet_by_name(str(ex_month_eng) + " " + str(ex_year))
        total_row = sheet.nrows
        for i in range(0, total_row - 1):
            received_date = xlrd.xldate.xldate_as_datetime(sheet.cell(i + 1, 7).value, 0)
            closed_date = xlrd.xldate.xldate_as_datetime(sheet.cell(i + 1, 6).value, 0)
            if "Azure" in sheet.cell(i + 1, 12).value:
                if "CME add" in sheet.cell(i + 1, 10).value:
                    self.account_data['user_account_ms_creation'] += int(sheet.cell(i + 1, 2).value)
                    if (received_date - closed_date).days > 1 and (sheet.cell(i + 1, 13).value == ""):
                        self.account_data['miss_sla_user_account_ms_creation'] += 1
                        self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"
                elif "CME modify" in sheet.cell(i + 1, 10).value:
                    self.account_data['ser_account_ms_modification'] += int(sheet.cell(i + 1, 2).value)
                    if (received_date - closed_date).days > 1 and (sheet.cell(i + 1, 13).value == ""):
                        self.account_data['miss_sla_user_account_ms_modification'] += 1
                        self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"
                elif "CME delete" in sheet.cell(i + 1, 10).value:
                    self.account_data['user_account_ms_termination'] += int(sheet.cell(i + 1, 2).value)
                    if (received_date - closed_date).days > 1 and (sheet.cell(i + 1, 13).value == ""):
                        self.account_data['miss_sla_user_account_ms_termination'] += 1
                        self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"
                elif "SG" in sheet.cell(i + 1, 10).value:
                    self.account_data['security_group'] += int(sheet.cell(i + 1, 2).value)
                    if (received_date - closed_date).days > 1 and (sheet.cell(i + 1, 13).value == ""):
                        self.account_data['miss_sla_security_group'] += 1
                        self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"
            elif "21V" in sheet.cell(i + 1, 12).value:
                if "CME add" in sheet.cell(i + 1, 10).value:
                    self.account_data['user_account_21v_creation'] += int(sheet.cell(i + 1, 2).value)
                    if (received_date - closed_date).days > 1 and (sheet.cell(i + 1, 13).value == ""):
                        self.account_data['iss_sla_user_account_21v_creation'] += 1
                        self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"
                elif "CME modify" in sheet.cell(i + 1, 10).value:
                    self.account_data['user_account_21v_modification'] += int(sheet.cell(i + 1, 2).value)
                    if (received_date - closed_date).days > 1 and (sheet.cell(i + 1, 13).value == ""):
                        self.account_data['miss_sla_user_account_21v_modification'] += 1
                        self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"
                elif "CME delete" in sheet.cell(i + 1, 10).value:
                    self.account_data['user_account_21v_termination'] += int(sheet.cell(i + 1, 2).value)
                    if (received_date - closed_date).days > 1 and (sheet.cell(i + 1, 13).value == ""):
                        self.account_data['miss_sla_user_account_21v_termination'] += 1
                        self.flag = self.flag + "{:.0f}".format(sheet.cell(i + 1, 0).value) + " \n"

    def logger(self):
        return "\n--------------------------------\nCME Account Statistics: \n" \
               "User_account_21v_creation: %d,  Miss_SLA of User_account_21v_creation: %d \n" \
               "User_account_21v_modification: %d,  Miss_SLA of User_account_21v_modification: %d \n" \
               "User_account_21v_termination: %d,   Miss_SLA of User_account_21v_termination: %d \n" \
               "User_account_ms_creation: %d,   Miss_SLA of User_account_ms_creation: %d \n" \
               "User_account_ms_modification: %d,   Miss_SLA of User_account_ms_modification: %d \n" \
               "User_account_ms_termination: %d,    Miss_SLA of User_account_ms_termination: %d \n" \
               "Security_group: %d, Miss_SLA of Security_group: %d \n\n" \
               % (self.account_data["user_account_21v_creation"],
                  self.account_data["miss_sla_user_account_21v_creation"],
                  self.account_data["user_account_21v_modification"],
                  self.account_data["miss_sla_user_account_21v_modification"],
                  self.account_data["user_account_21v_termination"],
                  self.account_data["miss_sla_user_account_21v_termination"],
                  self.account_data["user_account_ms_creation"],
                  self.account_data["miss_sla_user_account_ms_creation"],
                  self.account_data["user_account_ms_modification"],
                  self.account_data["miss_sla_user_account_ms_modification"],
                  self.account_data["user_account_ms_termination"],
                  self.account_data["miss_sla_user_account_ms_termination"],
                  self.account_data["security_group"],
                  self.account_data["miss_sla_security_group"]
                  )


class incident(object):
    def __init__(self, data_file):
        self.incident_route = data_file
        self.incident_num = 0
        self.miss_sla = 0

    def count_incident(self):
        book = xlrd.open_workbook(self.incident_route)
        sheet = book.sheet_by_index(0)
        self.incident_num = sheet.nrows - 1
        for i in range(0, self.incident_num):
            if int(re.sub(" \D ", "", sheet.cell(i + 1, 4).value )) > 15:
                self.miss_sla += 1

    def logger(self):
        return "\n--------------------------------\nIncident statistics: \nTotal Incidents:  %s,    Miss_SLA of Incidents:  %s\n\n" % (
        self.incident_num, self.miss_sla)


def compute_percent(sla, miss_sla):
    if sla == 0:
        percent = {"sla_percent": "100%(0)",
                   "miss_sla_percent": "0%(0)"}
        return percent
    percent = {"sla_percent": "{:.0%}({:.0f})".format((1 - miss_sla / sla), sla - miss_sla),
               "miss_sla_percent": "{:.0%}({:.0f})".format((miss_sla / sla), miss_sla)}
    return percent


def compute_title():
    ex_month_eng = calendar.month_name[time.localtime().tm_mon - 1]
    ex_month_range = calendar.monthrange(time.localtime().tm_year, time.localtime().tm_mon - 1)
    month_title = ex_month_eng + " 1 - " + ex_month_eng + " " + str(ex_month_range[1])
    return month_title


class report_maker(object):
    def __init__(self, msr_report_route, incentive_report_route, certificate_ob, account_ob, incident_ob):
        self.msr_report = str.replace(msr_report_route, r'/', '\\')
        self.incentive_report = incentive_report_route
        self.certificate = certificate_ob
        self.account = account_ob
        self.incident = incident_ob

    def msr_report_make(self):
        msr_excel = win32com.client.Dispatch('Excel.Application')
        msr_excel.DisplayAlerts = False
        msr_excel.Visible = 0
        msr_book = msr_excel.Workbooks.Open(self.msr_report)
        msr_sheet = msr_book.Worksheets("MSR KPI Value")
        msr_sheet.Cells(4, 2).Value = self.account.account_data["user_account_21v_creation"]
        msr_sheet.Cells(4, 3).Value = compute_percent(self.account.account_data["user_account_21v_creation"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_21v_creation"])["sla_percent"]
        msr_sheet.Cells(4, 4).Value = compute_percent(self.account.account_data["user_account_21v_creation"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_21v_creation"])["miss_sla_percent"]

        msr_sheet.Cells(5, 2).Value = self.account.account_data["user_account_21v_modification"]
        msr_sheet.Cells(5, 3).Value = compute_percent(self.account.account_data["user_account_21v_modification"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_21v_modification"])["sla_percent"]
        msr_sheet.Cells(5, 4).Value = compute_percent(self.account.account_data["user_account_21v_modification"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_21v_modification"])["miss_sla_percent"]

        msr_sheet.Cells(6, 2).Value = self.account.account_data["user_account_21v_termination"]
        msr_sheet.Cells(6, 3).Value = compute_percent(self.account.account_data["user_account_21v_termination"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_21v_termination"])["sla_percent"]
        msr_sheet.Cells(6, 4).Value = compute_percent(self.account.account_data["user_account_21v_termination"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_21v_termination"])["miss_sla_percent"]

        msr_sheet.Cells(7, 2).Value = self.account.account_data["user_account_ms_creation"]
        msr_sheet.Cells(7, 3).Value = compute_percent(self.account.account_data["user_account_ms_creation"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_ms_creation"])["sla_percent"]
        msr_sheet.Cells(7, 4).Value = compute_percent(self.account.account_data["user_account_ms_creation"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_ms_creation"])["miss_sla_percent"]

        msr_sheet.Cells(8, 2).Value = self.account.account_data["user_account_ms_modification"]
        msr_sheet.Cells(8, 3).Value = compute_percent(self.account.account_data["user_account_ms_modification"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_ms_modification"])["sla_percent"]
        msr_sheet.Cells(8, 4).Value = compute_percent(self.account.account_data["user_account_ms_modification"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_ms_modification"])["miss_sla_percent"]

        msr_sheet.Cells(9, 2).Value = self.account.account_data["user_account_ms_termination"]
        msr_sheet.Cells(9, 3).Value = compute_percent(self.account.account_data["user_account_ms_termination"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_ms_termination"])["sla_percent"]
        msr_sheet.Cells(9, 4).Value = compute_percent(self.account.account_data["user_account_ms_termination"],
                                                      self.account.account_data[
                                                          "miss_sla_user_account_ms_termination"])["miss_sla_percent"]

        msr_sheet.Cells(11, 2).Value = self.account.account_data["security_group"]
        msr_sheet.Cells(11, 3).Value = compute_percent(self.account.account_data["security_group"],
                                                       self.account.account_data[
                                                           "miss_sla_security_group"])["sla_percent"]
        msr_sheet.Cells(11, 4).Value = compute_percent(self.account.account_data["security_group"],
                                                       self.account.account_data[
                                                           "miss_sla_security_group"])["miss_sla_percent"]

        # write certificate data
        msr_sheet.Cells(21, 2).Value = self.certificate.total_certificates
        msr_sheet.Cells(21, 3).Value = compute_percent(self.certificate.total_certificates,
                                                       self.certificate.miss_sla)["sla_percent"]
        msr_sheet.Cells(21, 4).Value = compute_percent(self.certificate.total_certificates,
                                                       self.certificate.miss_sla)["miss_sla_percent"]

        # write incident data
        msr_sheet.Cells(26, 2).Value = self.incident.incident_num
        msr_sheet.Cells(26, 3).Value = compute_percent(self.incident.incident_num,
                                                       self.incident.miss_sla)["sla_percent"]
        msr_sheet.Cells(26, 4).Value = compute_percent(self.incident.incident_num,
                                                       self.incident.miss_sla)["miss_sla_percent"]

        # write title
        msr_sheet.Cells(2, 2).Value = msr_sheet.Cells(13, 2).Value = msr_sheet.Cells(28, 2).Value = \
            msr_sheet.Cells(41, 2).Value = compute_title()

        msr_book.Close(SaveChanges=1)
        msr_excel.Quit()

    def incentive_report_make(self):
        ppt = win32com.client.Dispatch('PowerPoint.Application')
        msr_ppt = ppt.Presentations.Open(str(self.incentive_report).replace("/", "\\"))

        # write cme account data
        msr_ppt.Slides(1).Shapes(1).Table.Cell(4, 2).Shape.TextFrame.TextRange.Text = self.account.account_data[
            "user_account_21v_creation"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(4, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_21v_creation"],
                        self.account.account_data[
                            "miss_sla_user_account_21v_creation"])["sla_percent"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(4, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_21v_creation"],
                        self.account.account_data[
                            "miss_sla_user_account_21v_creation"])["miss_sla_percent"]

        msr_ppt.Slides(1).Shapes(1).Table.Cell(5, 2).Shape.TextFrame.TextRange.Text = self.account.account_data[
            "user_account_21v_modification"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(5, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_21v_modification"],
                        self.account.account_data[
                            "miss_sla_user_account_21v_modification"])["sla_percent"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(5, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_21v_modification"],
                        self.account.account_data[
                            "miss_sla_user_account_21v_modification"])["miss_sla_percent"]

        msr_ppt.Slides(1).Shapes(1).Table.Cell(6, 2).Shape.TextFrame.TextRange.Text = self.account.account_data[
            "user_account_21v_termination"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(6, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_21v_termination"],
                        self.account.account_data[
                            "miss_sla_user_account_21v_termination"])["sla_percent"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(6, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_21v_termination"],
                        self.account.account_data[
                            "miss_sla_user_account_21v_termination"])["miss_sla_percent"]

        msr_ppt.Slides(1).Shapes(1).Table.Cell(7, 2).Shape.TextFrame.TextRange.Text = self.account.account_data[
            "user_account_ms_creation"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(7, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_ms_creation"],
                        self.account.account_data[
                            "miss_sla_user_account_ms_creation"])["sla_percent"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(7, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_ms_creation"],
                        self.account.account_data[
                            "miss_sla_user_account_ms_creation"])["miss_sla_percent"]

        msr_ppt.Slides(1).Shapes(1).Table.Cell(8, 2).Shape.TextFrame.TextRange.Text = self.account.account_data[
            "user_account_ms_modification"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(8, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_ms_modification"],
                        self.account.account_data[
                            "miss_sla_user_account_ms_modification"])["sla_percent"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(8, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_ms_modification"],
                        self.account.account_data[
                            "miss_sla_user_account_ms_modification"])["miss_sla_percent"]

        msr_ppt.Slides(1).Shapes(1).Table.Cell(9, 2).Shape.TextFrame.TextRange.Text = self.account.account_data[
            "user_account_ms_termination"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(9, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_ms_termination"],
                        self.account.account_data[
                            "miss_sla_user_account_ms_termination"])["sla_percent"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(9, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["user_account_ms_termination"],
                        self.account.account_data[
                            "miss_sla_user_account_ms_termination"])["miss_sla_percent"]

        msr_ppt.Slides(1).Shapes(1).Table.Cell(11, 2).Shape.TextFrame.TextRange.Text = self.account.account_data[
            "security_group"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(11, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["security_group"],
                        self.account.account_data[
                            "miss_sla_security_group"])["sla_percent"]
        msr_ppt.Slides(1).Shapes(1).Table.Cell(11, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.account.account_data["security_group"],
                        self.account.account_data[
                            "miss_sla_security_group"])["miss_sla_percent"]

        # write certificate data
        msr_ppt.Slides(2).Shapes(1).Table.Cell(9,
                                               2).Shape.TextFrame.TextRange.Text = self.certificate.total_certificates
        msr_ppt.Slides(2).Shapes(1).Table.Cell(9, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.certificate.total_certificates,
                        self.certificate.miss_sla)["sla_percent"]
        msr_ppt.Slides(2).Shapes(1).Table.Cell(9, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.certificate.total_certificates,
                        self.certificate.miss_sla)["miss_sla_percent"]

        # write incident data
        msr_ppt.Slides(2).Shapes(1).Table.Cell(14, 2).Shape.TextFrame.TextRange.Text = self.incident.incident_num
        msr_ppt.Slides(2).Shapes(1).Table.Cell(14, 3).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.incident.incident_num,
                        self.incident.miss_sla)["sla_percent"]
        msr_ppt.Slides(2).Shapes(1).Table.Cell(14, 4).Shape.TextFrame.TextRange.Text = \
        compute_percent(self.incident.incident_num,
                        self.incident.miss_sla)["miss_sla_percent"]

        # write title
        msr_ppt.Slides(1).Shapes(1).Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = \
            msr_ppt.Slides(2).Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = \
            msr_ppt.Slides(3).Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = \
            msr_ppt.Slides(4).Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = compute_title()

        msr_ppt.Save()
        ppt.Quit()

    def logger(self):
        return "\n--------------------------------\n**MSR Report & Incentive Report 21V have been created.** \n" \
               "Please find these file here: \n" \
               "%s\n" \
               "%s\n\n******************************************************************************************\n\n\n\n" % (self.msr_report, self.incentive_report)


class report_mover(object):
    def __init__(self):
        self.sum_file = str(time.localtime().tm_year) + "{:0>2d}".format(time.localtime().tm_mon - 1) + "-Monthly"
        self.cer_file = "../" + self.sum_file + "/About Cert Application-From WADE"
        self.inc_file = "../" + self.sum_file + "/Security Incident Report From IM"
        self.acc_file = "../" + self.sum_file + "/About CME Account Management-From RM"

    def mkdir(self):
        try:
            os.mkdir("../" + self.sum_file)
            os.mkdir(self.cer_file)
            os.mkdir(self.inc_file)
            os.mkdir(self.acc_file)
        finally:
            return 0


class logger(object):
    def __init__(self):
        self.log_file = open(os.getcwd() + "/log.txt", "w")
        self.history_file = open(os.getcwd() + "/system_history", "a")

    def input_log(self, content):
        self.log_file.write(content)
        self.history_file.write(content)

    def close_log(self):
        self.log_file.close()
        self.history_file.close()


def main():
    log = logger()
    log.input_log(datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d %H:%M:%S") + " log:\n\n")

    raw_files = get_file()
    raw_files.specify_file()

    certificates = certificate(raw_files.certificate_file)
    certificates.read_certificate()

    accounts = account(raw_files.account_raw_file)
    accounts.read_account_raw()

    incidents = incident(raw_files.incident_file)
    incidents.count_incident()

    reports = report_maker(raw_files.msr_report, raw_files.incentive_report, certificates, accounts, incidents)
    reports.msr_report_make()
    reports.incentive_report_make()

    msr_report = report_mover()
    msr_report.mkdir()

    # file copy
    shutil.copy(raw_files.msr_report.replace("/", "\\"),
                "../" + msr_report.sum_file + "/" + str(time.localtime().tm_year) + "{:0>2d}".format(
                    time.localtime().tm_mon - 1) + " " + os.path.split(raw_files.msr_report)[1])
    shutil.copy(raw_files.incentive_report.replace("/", "\\"),
                "../" + msr_report.sum_file + "/" + os.path.split(raw_files.incentive_report)[1])
    shutil.copy(raw_files.incident_file.replace("/", "\\"),
                msr_report.inc_file + "/" + os.path.split(raw_files.incident_file)[1])
    shutil.copy(raw_files.account_raw_file.replace("/", "\\"),
                msr_report.acc_file + "/" + os.path.split(raw_files.account_raw_file)[1])
    shutil.copy(raw_files.certificate_file.replace("/", "\\"),
                msr_report.cer_file + "/" + os.path.split(raw_files.certificate_file)[1])

    log.input_log(raw_files.logger())
    log.input_log(certificates.logger())
    log.input_log(certificates.flag)
    log.input_log(accounts.logger())
    log.input_log(accounts.flag)
    log.input_log(incidents.logger())
    log.input_log(reports.logger())

    log.close_log()


if __name__ == '__main__':
    main()
    # print(__name__)

