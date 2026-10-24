# Edit by Liam Hong
# This script is used for Tunnel status check of the stores
import time
from FMGRegisteredDevices import Getdeviceinfo
from netmiko import ConnectHandler
from getpass import getpass
from netmiko.exceptions import (
    NetmikoTimeoutException,
    NetmikoAuthenticationException
)
from concurrent.futures import ThreadPoolExecutor
import re
from functools import partial
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path


# step 3
def parse_tunnel(tunneloutput, storecode):
    result = {
        "storecode": storecode
    }

    pattern = r"'([^']+)'.*?selectors\(total,up\):\s*\d+/(\d+)"
    for line in tunneloutput.splitlines():
        match = re.search(pattern, line)
        if match:
            tunnel_name = match.group(1)
            up_count = int(match.group(2))
            result[tunnel_name] = "up" if up_count > 0 else "down"

    return result


# Step 2
def process_device(singledevice,username,password):
    storecode = singledevice["hostname"]
    ip = singledevice["ip"]

    # use netmiko connect device
    devicepara = {
        "device_type": "fortinet",
        "host": ip,
        "username": username,
        "password": password,
        "port": 22,
    }

    try:
        with ConnectHandler(**devicepara) as connection:
            tunneloutput = connection.send_command("get vpn ipsec tunnel summary",read_timeout=30)

            # call the parse_tunnel function and return it's output
            return parse_tunnel(tunneloutput, storecode)

    except NetmikoAuthenticationException:
        return {
            "storecode": storecode,
            "connection": "Authentication failed"
        }
    except NetmikoTimeoutException:
        return {
            "storecode": storecode,
            "connection": "Connection timeout"
        }
    except Exception as e:
        return {
            "storecode": storecode,
            "connection": f"Error: {e}"
        }


# step 1
def process_all_devices(target_devices):
    worker = partial(
        process_device,
        username=username,
        password=password
    )
    with ThreadPoolExecutor(max_workers=20) as executor:
        # iterate the target_devices, pass each memeber to worker
        # use list comprehension and map
        results = list(executor.map(worker, target_devices))

    # If connection failed, it will return None which is hard to handle with for Excel.
    results = [result for result in results if result is not None]
    return results


# write result into a Excel
def write_to_excel(results, filename="Tunnel_Status.xlsx"):
    # Create a new Excel workbook in memory.
    workbook = Workbook()
    # Get the active worksheet from the workbook.
    worksheet = workbook.active
    # Rename the active worksheet to "Tunnel Status".
    worksheet.title = "Tunnel Status"

    # Excel column headers
    headers = [
        "Storecode",
        "Tunnel1",
        "Tunnel2",
        "Tunnel3",
        "Tunnel4"
    ]
    # Append the headers as a new row to the worksheet.
    worksheet.append(headers)

    # Write each device result
    for result in results:
        storecode = result["storecode"]

        if "connection" in result:
            # Connection failed: write the reason in column B
            row = [
                storecode,
                result["connection"],
                "",
                "",
                ""
            ]
        else:
            # Connection succeeded: write the tunnel status
            row = [
                storecode,
                result.get("Tunnel1", ""),
                result.get("Tunnel2", ""),
                result.get("Tunnel3", ""),
                result.get("Tunnel4", "")
            ]

        worksheet.append(row)

    # Header formatting
    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )
    # set the first row(header)
    for cell in worksheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Column widths
    worksheet.column_dimensions["A"].width = 22
    worksheet.column_dimensions["B"].width = 25
    worksheet.column_dimensions["C"].width = 15
    worksheet.column_dimensions["D"].width = 15
    worksheet.column_dimensions["E"].width = 15

    # Freeze header row
    worksheet.freeze_panes = "A2"

    # Save to the script's current working directory
    # resolve is used to Convert a relative path to a full absolute path.
    output_path = Path(filename).resolve()
    workbook.save(output_path)

    return output_path


if __name__ == "__main__":
    adom = input("Enter adomname: ")
    username = input("Enter your username: ")
    password = getpass("Enter your password: ")
    start = time.time()
    # get all devices informations from FMG
    device_list = Getdeviceinfo()

    # acquire information from specific ADOM
    target_devices = [device for device in device_list if device["adom"] == adom]
    #pprint(target_devices)
    results = process_all_devices(target_devices)

    #pprint(results)
    # write to a table and create a excel table
    output_file = write_to_excel(
        results,
        filename=f"{adom}_Tunnel_Status.xlsx"
    )

    cost = time.time() - start
    print(f"Cost: {cost:.2f} seconds")

