import json
import logging
import os
from datetime import datetime

import ckanapi
from openpyxl.reader.excel import load_workbook

CONFIG_FILENAME = os.getenv('CONFIG_FILENAME', 'config.json')
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILENAME)
DATASET_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources/family-medicine-reports')
root_dir = os.path.dirname(os.path.abspath(__file__))
with open(CONFIG_PATH, 'r') as config_file:
    CONFIG = json.loads(config_file.read())['config']

MONTHS = {
    1: "Januart",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December"
}


def mutable_dataset_dict(title, name, month, resources, **kwargs):
    return {
        "title": title,
        "name": name,
        "type": "family-medicine",
        "notes": kwargs.get("notes", ""),
        "owner_org": "who-romania",
        "maintainer": kwargs.get("Admin", "Admin"),
        "maintainer_email": kwargs.get("maintainer_email", "admin@localhost"),
        "month": month,
        "groups": [{"name": "family-medicine"}],
        "tags": [{"name": "Data"}],
        "year": kwargs.get("year", 2023),
        "resources": resources
    }


def mutable_resource_dict(name, file_path, week, family_doctor):
    return {
        "name": name,
        "filename": file_path,
        "format": "XLSX",
        "week": week,
        "family_doctor": family_doctor
    }


def traverse_folder_tree():
    subfolders = [int(dir_name) for dir_name in os.listdir(CONFIG['resource_folder'])]
    subfolders.sort(reverse=True)
    files = []
    for subfolder in subfolders:
        for dirpath, dirnames, filenames in os.walk(CONFIG['resource_folder'] + '/' + str(subfolder)):
            # TODO edit folders names with full year
            year = '20' + (str(subfolder)[:2])
            # TODO fix this
            month = (str(subfolder).split('23')[1])
            folders_dict = {"year": year, "month": month, "files": []}
            for file in filenames:
                if "month" in file.lower(): # skip monthly files
                    continue
                if os.path.splitext(file)[1] in ['.xls', '.xlsx']:
                    folders_dict["files"].append(os.path.join(dirpath, file))
            files.append(folders_dict)
    return files


def read_resource_sheet(workbook, sheet):
    active = workbook[sheet]
    report_date = active['B2'].value
    fd_name = sheet.split('FD ')[1]

    if not isinstance(report_date, datetime):
        logging.warning(f"Report date {report_date} is not a date.")
        print("Failed to use report date, attempting to load from report period ...")
        report_date = active['B3'].value
        if not isinstance(report_date, datetime):
            logging.error(f"Report date {report_date} is not a date.")

    if report_date.weekday() != 4:
        logging.info(f"Report date {report_date} is not a Friday.")

    year = str(report_date.year % 100)
    month = "{:02}".format(report_date.month)
    day = "{:02}".format(report_date.day)
    week = year + '-' + month + '-' + day
    week_number = str(report_date.isocalendar()[1])

    new_folder_path = os.path.join(root_dir, "resources/family-medicine-reports", month, week_number, sheet)

    template_workbook = load_workbook(root_dir + "/template.xlsx")
    template_sheet = template_workbook.active

    start_row, end_row = 6, 41
    start_col, end_col = 1, 23  # A to W

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell_value = active.cell(row=row, column=col).value
            template_sheet.cell(row=row, column=col, value=cell_value)

    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)

    template_workbook.save(new_folder_path + '/report.xlsx')

    return mutable_resource_dict("Report from Family Doctor " + fd_name + " for week number " + week_number,
                                 'report.xlsx',
                                 week,
                                 fd_name)


def load_file_sheets(file):
    print("Reading file: ", file)
    workbook = load_workbook(file)
    sheets = [item for item in workbook.sheetnames if item.startswith('FD')]
    sheets.sort()
    all_resources = []
    for sheet in sheets:
        all_resources.append(read_resource_sheet(workbook, sheet))
    return all_resources


def generate_dataset_dict():
    dataset_dict = {
        "datasets": []
    }
    fd_files = traverse_folder_tree()
    for file_dict in fd_files:
        resources = []
        for file in file_dict["files"]:
            resources.extend(load_file_sheets(file))

        dataset_dict["datasets"].append(mutable_dataset_dict(
            title="Family Medicine Reports" + " for " + MONTHS[int(file_dict["month"])] + " " + file_dict["year"],
            name="family-medicine-reports" + "-" + (file_dict["month"]) + "-" + (file_dict["year"]),
            month=file_dict["year"] + "-" + file_dict["month"],
            year=file_dict["year"],
            notes="WHO ROMANIA - Family Medicine Reports for " + MONTHS[int(file_dict["month"])] + " " + file_dict["year"],
            resources=resources
        ))

    with open(root_dir + '/resources/datasets.json', 'w') as json_file:
        json.dump(dataset_dict, json_file, indent=4)


if __name__ == '__main__':
    generate_dataset_dict()
    # TODO sort files by week number
    # TODO create dataset for each month
    # TODO create resource for each week
    # TODO create resource for each family doctor
