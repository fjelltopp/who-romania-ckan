import logging
from resource_loader import *

log = logging.getLogger(__name__)
log.setLevel(logging.INFO)


def count_number_of_files_and_sheet():
    fd_files = traverse_folder_tree()
    num_files = 0
    num_sheets = 0
    family_doctors = []
    for file_dict in fd_files:
        num_files += 1
        num_sheets += len(file_dict["files"])
        for file in file_dict["files"]:
            workbook = load_workbook(file)
            sheets = [item for item in workbook.sheetnames if item.startswith('FD')]
            sheets.sort()
            for sheet in sheets:
                fd_name = sheet.split('FD ')[1]
                found = False
                for doctor in family_doctors:
                    if doctor["name"] == fd_name:
                        doctor["sheets"] += 1

                        active = workbook[sheet]
                        report_date = active['B2'].value
                        # add report date as week to dict (remove the 20 in year)
                        doctor["week"].append(str(report_date.year % 100) + "-" + "{:02}".format(
                            report_date.month) + "-" + "{:02}".format(report_date.day))
                        found = True
                        break
                if not found:
                    family_doctors.append({"name": fd_name, "sheets": 1, "week": []})
    return family_doctors


def traverse_output_resource_files():
    """
    Traverses the resources/family-medicine-reports folder tree and returns a list of all EXCEL files paths in the tree along
    """
    num_files = 0
    family_doctors = []
    for subdir, _, files in os.walk(RESOURCE_FOLDER):
        for file in files:
            if file.endswith('.xlsx'):
                num_files += 1
                filepath = os.path.join(subdir, file)

                family_doctor_name = os.path.basename(os.path.dirname(filepath)).split('FD ')[1]

                found = False
                for doctor in family_doctors:
                    if doctor["name"] == family_doctor_name:
                        doctor["files"] += 1
                        week = os.path.basename(os.path.dirname(os.path.dirname(filepath)))
                        doctor["week"].append(week)
                        found = True
                        break
                if not found:
                    family_doctors.append({"name": family_doctor_name, "files": 1,
                                           "week": [os.path.basename(os.path.dirname(os.path.dirname(filepath)))]})

    return family_doctors


def compare_family_doctors_lists_and_get_missing_weeks():
    fd_raw = count_number_of_files_and_sheet()
    fd_processed = traverse_output_resource_files()
    missing_week = []
    for doctor in fd_raw:
        found = False
        for doctor2 in fd_processed:
            if doctor["name"] == doctor2["name"]:
                found = True

                if doctor["sheets"] != doctor2["files"]:
                    log.warning(f"Doctor {doctor['name']} has {doctor['sheets']} sheets but {doctor2['files']} files")

                for week in doctor["week"]:
                    if week not in doctor2["week"]:
                        log.warning(f"Doctor {doctor['name']} is missing week {week}")
                        missing_week.append({"doctor": doctor["name"], "week": week})
                        with open(root_dir + '/logs/missing_weeks.csv', 'a') as file:
                            file.write(f"{doctor['name']},{week}\n")
                break
        if not found:
            log.error(f"Doctor {doctor['name']} is missing all weeks")
    return missing_week


def get_weeks_for_family_doctor(family_doctor_name):
    fd_raw = count_number_of_files_and_sheet()
    fd_processed = traverse_output_resource_files()
    # return a dict with the weeks for the family doctor from the raw data and the processed data
    weeks = {"raw": [], "processed": []}
    for doctor in fd_raw:
        if doctor["name"] == family_doctor_name:
            weeks["raw"] = doctor["week"]
            break
    for doctor in fd_processed:
        if doctor["name"] == family_doctor_name:
            weeks["processed"] = doctor["week"]
            break
    log.warning(f"Raw weeks {len(weeks['raw'])}:\n{json.dumps(weeks['raw'],indent=2)}")
    log.warning(f"Processed weeks {len(weeks['processed'])}:\n{json.dumps(weeks['processed'],indent=2)}")
    # log the missing weeks
    for week in weeks["raw"]:
        if week not in weeks["processed"]:
            log.warning(f"Doctor {family_doctor_name} is missing week {week}")


def init_log_files():
    if CONFIG["overwrite_logs"]:
        with open(root_dir + '/logs/report_date_error.csv', 'w') as file:
            file.write(f"file,sheet,report_date\n")
        with open(root_dir + '/logs/not_friday_error.csv', 'w') as file:
            file.write(f"file,sheet,report_date,weekday\n")
    with open(root_dir + '/logs/missing_weeks.csv', 'w') as file:
        file.write(f"fd,week\n")


def swap_report_dates_and_periods_cell():
    filename = root_dir + "/logs/report_date_error.csv"

    if not os.path.exists(filename):
        raise Exception("File not found")

    with open(filename, 'r') as file:
        csv_reader = csv.reader(file)

        if len(list(csv_reader)) <= 1 or not CONFIG['swap_report_dates_and_periods']:
            log.info("No report date errors found")
            return
        next(csv_reader)
        for row in csv_reader:
            current_file = row[0]

            workbook = load_workbook(os.path.join(root_dir, current_file))
            sheet = workbook[row[1]]
            report_date = sheet['B2'].value
            report_period = sheet['B3'].value
            sheet['B2'] = report_period
            sheet['B3'] = report_date
            workbook.save(current_file)
            log.info(f"Swapped report date and period for file {current_file} and sheet {row[1]}")


if __name__ == '__main__':
    # The swap will happen if CONFIG['swap_report_dates_and_periods'] is True
    swap_report_dates_and_periods_cell()
    # The logs will be overwritten if CONFIG['overwrite_logs'] is True
    init_log_files()
    compare_family_doctors_lists_and_get_missing_weeks()

