# Edit by Liam
# This script is for getting the store information base on storecode or ADOM name
# It's a better version of StoreInventory.py


from FMGRegisteredDevices import Getdeviceinfo
from concurrent.futures import ThreadPoolExecutor
from getpass import getpass
from pprint import pprint
from netmiko import ConnectHandler
from netmiko.exceptions import (NetmikoTimeoutException,NetmikoAuthenticationException)
from functools import partial
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path

# step 1
def get_target_devices(user_input):
    device_list = Getdeviceinfo()

    # Check whether the input is a storecode
    # list comprehension
    store_devices = [device for device in device_list if device["hostname"] == user_input]
    if store_devices:
        #pprint(store_devices)
        return store_devices

    # Check whether the input is an ADOM
    adom_devices = [device for device in device_list if device["adom"] == user_input]
    if adom_devices:
        #pprint(adom_devices)
        return adom_devices

    # nonthing match return an empty list
    return []


# Use netmiko to connect to FortiGate and execute some command to get information
def connect_fgt(ip,username,password):
    device = {
        "device_type": "fortinet",
        "host": ip,
        "username": username,
        "password": password,
        "port": 22,
    }

    try:
        with ConnectHandler(**device) as connection:
            output_sw = connection.send_command("execute switch-controller get-conn-status",read_timeout=30)
            output_ap = connection.send_command("show wireless-controller wtp",read_timeout=30)
            return output_sw, output_ap, None

    except NetmikoAuthenticationException:
        return None, None, "Authentication failed"
    except NetmikoTimeoutException:
        return None, None, "Connection timeout"
    except Exception as e:
        return None, None, f"Error: {e}"


def get_model_from_sn(serial):
    MODEL_MAP = {
        "FP231K": "FortiAP-231K",
        "PU431F": "FortiAP-U431F",
        "FP431F": "FortiAP-431F",
        "FP321E": "FortiAP-321E",
        "FP431G": "FortiAP-431G",
        "S108EF": "FortiSwitch-108E-FPOE",
        "S108FF": "FortiSwitch-108F-FPOE",
        "S124GF": "FortiSwitch-124G-FPOE",
        "S124FF": "FortiSwitch-124F-FPOE",
        "S124EF": "FortiSwitch-124E-FPOE",
        "S148FF": "FortiSwitch-148F-FPOE",

    }
    for prefix, model in MODEL_MAP.items():
        if serial.startswith(prefix):
            return model
    return "Unknown"


# This function is for getting infomation of FortiSwitch
def get_sw(output):
    swlist = []
    if not output:
        return swlist

    for line in output.splitlines():
        if "Authorized/Up" in line:
            parts = line.split()
            switch_name = parts[0]
            serial = parts[-1]

            swlist.append({
                "hostname": switch_name,
                "sn": serial,
                "platform": get_model_from_sn(serial)
            })
    return swlist


def get_ap(output):
    aplist = []
    serial = None
    if not output:
        return aplist

    for line in output.splitlines():
        line = line.strip()
        # check whether string is start with specific contents.
        if line.startswith("edit "):
            serial = line.split('"')[1]
        elif line.startswith("set name ") and serial:       # ensure already match the serial
            name = line.split('"')[1]
            aplist.append({
                "hostname": name,
                "sn": serial,
                "platform": get_model_from_sn(serial)
            })
            serial = None
    return aplist


def process_device(singledevice,username,password):
    # Get device IP address
    ip = singledevice["ip"]
    adom = singledevice["adom"]
    hostname = singledevice["hostname"]
    sn = singledevice["sn"]
    platform = singledevice["platform"]

    output_sw, output_ap, error = connect_fgt(ip,username, password)
    if error:
        storeinfo = {
            "storecode": hostname,
            "adom": adom,
            "FGT": {
                "hostname": hostname,
                "platform": platform,
                "sn": sn,
                "ip": ip,
            },
            "SW": [],
            "AP": [],
            "status": "Failed",
            "error": error
        }

    else:
        swlist = get_sw(output_sw)
        aplist = get_ap(output_ap)

        storeinfo = {
            "storecode": hostname,
            "adom": adom,
            "FGT": {
                "hostname": hostname,
                "platform": platform,
                "sn": sn,
                "ip": ip,
            },
            "SW": swlist,
            "AP": aplist,
            "status": "Success",
            "error": ""
    }
    #pprint(storeinfo)
    return storeinfo


# Step 2
def process_all_devices(target_devices, username, password):
    worker = partial(
        # call the Function "process_device"
        process_device,
        username=username,
        password=password
    )

    with ThreadPoolExecutor(max_workers=20) as executor:
        # list comprehension and map to create a list of device informations
        results = list(executor.map(worker, target_devices))

    return results


def write_to_excel(results,user_input,userfilename):
    # Create a new Excel workbook in memory.
    workbook = Workbook()
    # Get the active worksheet from the workbook.
    worksheet = workbook.active
    worksheet.title = "Store Inventory"

    # Write the storecode in the first row
    worksheet.append(["Storecode / ADOM", user_input])

    # Write the column headers in the second row
    headers = [
        "Storecode",
        "Device Type",
        "Hostname",
        "Platform",
        "SN",
        "Status",
        "Error",
    ]
    worksheet.append(headers)

    # Write FortiGate information
    for storeinfo in results:
        if storeinfo is None:
            continue

        fgt = storeinfo["FGT"]
        worksheet.append([
            storeinfo["storecode"],
            "FGT",
            fgt["hostname"],
            fgt["platform"],
            fgt["sn"],
            storeinfo["status"],
            storeinfo["error"],
        ])

        # Write FortiSwitch information
        for switch in storeinfo["SW"]:
            worksheet.append([
                storeinfo["storecode"],
                "SW",
                switch["hostname"],
                switch["platform"],
                switch["sn"],
                "",
                "",
            ])

        # Write FortiAP information
        for ap in storeinfo["AP"]:
            worksheet.append([
                storeinfo["storecode"],
                "AP",
                ap["hostname"],
                ap["platform"],
                ap["sn"],
                "",
                "",
            ])

    # Format the storecode row
    worksheet["A1"].font = Font(color="FFFFFF", bold=True)
    worksheet["A1"].fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    worksheet["B1"].font = Font(color="FFFFFF", bold=True)
    worksheet["B1"].fill = PatternFill(fill_type="solid", fgColor="1F4E78")

    # Format the header row
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")

    # 获取 Excel 的第 2 行所有单元格
    for cell in worksheet[2]:
        cell.font = Font(
            color="FFFFFF",
            bold=True
        )
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Set column widths
    worksheet.column_dimensions["A"].width = 16
    worksheet.column_dimensions["B"].width = 40
    worksheet.column_dimensions["C"].width = 28
    worksheet.column_dimensions["D"].width = 25
    worksheet.column_dimensions["E"].width = 25
    worksheet.column_dimensions["F"].width = 15
    worksheet.column_dimensions["G"].width = 40

    # Freeze the header rows，冻结A3上方的行。
    worksheet.freeze_panes = "A3"

    # Generate the full output path
    output_path = Path(userfilename).resolve()
    # Save the workbook
    workbook.save(output_path)

    return output_path


def main():
    user_input = input("Enter storecode or ADOM name: ")
    username = input("Enter your username: ")
    password = getpass("Enter your password: ")

    # Get target devices in a list
    target_devices = get_target_devices(user_input)
    # if target_devices is an empty list [], the condition is like if not false = if true
    if not target_devices:
        print("Storecode or ADOM not found.")
        return

    # Process target devices
    results = process_all_devices(
        target_devices,
        username,
        password
    )

    # write to Excel
    output_file = write_to_excel(results, user_input, userfilename=f"{user_input}_StoreInventory.xlsx")


if __name__ == "__main__":
    main()