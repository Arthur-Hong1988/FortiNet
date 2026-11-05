# Edit by Liam 20261104
# Check the quality of WAN1 and WAN2 via LinkMonitor
# 1.create another Linkmonitor for WAN2 but "set update-static-route disable" so it only monitor the link quality
# it won't affect the route.
# 2.get packet loss, jitter, latency information for wan1 and wan2 separately
# 3.Write the statistics into a table.


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
        # pprint(store_devices)
        return store_devices

    # Check whether the input is an ADOM
    adom_devices = [device for device in device_list if device["adom"] == user_input]
    if adom_devices:
        # pprint(adom_devices)
        return adom_devices

    # nonthing match return an empty list
    return []


# Use netmiko to connect to FortiGate and execute some command to get information
def connect_fgt(ip, username, password):
    device = {
        "device_type": "fortinet",
        "host": ip,
        "username": username,
        "password": password,
        "port": 22,
    }

    try:
        with ConnectHandler(**device) as connection:
            output_LM = connection.send_command("diagnose sys link-monitor status", read_timeout=30)
            return output_LM, None

    except NetmikoAuthenticationException:
        return None, "Authentication failed"
    except NetmikoTimeoutException:
        return None, "Connection timeout"
    except Exception as e:
        return None, f"Error: {e}"


# This function is for getting infomation of FortiSwitch
def info_processor(output,adom,hostname):
    results = []
    current_interface = None
    current_peer = None

    for line in output.splitlines():   # splitlines splits a string into a list of lines
        line = line.strip()            # removes leading and trailing whitespace from a string.

        # 1. Source interface
        if line.startswith("Source interface:"):
            # Source interface: wan1 (5)
            # splits a string into a list, using whitespace as the default separator.
            current_interface = line.split()[2]

        # 2. Peer
        elif line.startswith("Peer:"):
            # 如果前一个 peer 已经采集完成，先保存
            if current_peer:
                results.append(current_peer)

            # Peer: 8.8.8.8(8.8.8.8)
            peer = line.split()[1].split("(")[0]

            current_peer = {
                "adom": adom,
                "storecode": hostname,
                "interface": current_interface,
                "peer": peer,
                "latency_min": None,
                "latency_max": None,
                "latency_avg": None,
                "jitter_min": None,
                "jitter_max": None,
                "jitter_avg": None,
                "packet_loss": None
            }

        # 3. Latency
        elif line.startswith("Latency(Min/Max/Avg):") and current_peer:
            # Latency(Min/Max/Avg): 7.813/8.089/7.946 ms
            value = line.split(":")[1].strip().split()[0]
            min_value, max_value, avg_value = value.split("/")
            current_peer["latency_min"] = float(min_value)
            current_peer["latency_max"] = float(max_value)
            current_peer["latency_avg"] = float(avg_value)

        # 4. Jitter
        elif line.startswith("Jitter(Min/Max/Avg):") and current_peer:
            # Jitter(Min/Max/Avg): 0.015/0.228/0.089 ms
            value = line.split(":")[1].strip().split()[0]
            min_value, max_value, avg_value = value.split("/")
            current_peer["jitter_min"] = float(min_value)
            current_peer["jitter_max"] = float(max_value)
            current_peer["jitter_avg"] = float(avg_value)

        # 5. Packet loss
        elif line.startswith("Packet lost:") and current_peer:
            # Packet lost: 0.000%
            value = line.split(":")[1].strip().replace("%", "")
            current_peer["packet_loss"] = float(value)

    # 最后一个 peer 后面不会再出现新的 Peer，所以手动保存
    if current_peer:
        results.append(current_peer)

    return results


def process_device(singledevice, username, password):
    # Get device IP address
    ip = singledevice["ip"]
    adom = singledevice["adom"]
    hostname = singledevice["hostname"]
    output_LM, error = connect_fgt(ip, username, password)
    if error:
        return [{
            "adom": adom,
            "storecode": hostname,
            "interface": None,
            "peer": None,
            "latency_min": None,
            "latency_max": None,
            "latency_avg": None,
            "jitter_min": None,
            "jitter_max": None,
            "jitter_avg": None,
            "packet_loss": None,
            "error": error
        }]

    quality_info = info_processor(output_LM, adom, hostname)

    #pprint(quality_info)
    return quality_info


# Step 2
def process_all_devices(target_devices, username, password):
    worker = partial(
        # call the Function "process_device"
        process_device,
        username=username,
        password=password
    )

    with ThreadPoolExecutor(max_workers=10) as executor:
        # list comprehension and map to create a list of device informations
        results = list(executor.map(worker, target_devices))

    return results


def write_to_excel(results, user_input, userfilename):
    # Creates a new Excel workbook
    workbook = Workbook()
    # Get the active worksheet from the workbook.
    worksheet = workbook.active
    worksheet.title = "Link Quality"

    # Write the column headers in the second row
    headers = [
        "adom",
        "storecode",
        "interface",
        "peer",
        "jitter_avg",
        "jitter_max",
        "jitter_min",
        "latency_avg",
        "latency_max",
        "latency_min",
        "packet_loss",
        "error"
    ]

    # Write headers
    worksheet.append(headers)

    # Header formatting
    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    # Apply formatting to all header cells
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",            # 水平居中
            vertical="center"               # 垂直居中
        )

    # Yellow fill for packet loss greater than 5%
    warning_fill = PatternFill(
        fill_type="solid",
        fgColor="FFFF00"
    )

    # results is a list of device result lists
    for device_results in results:

        # device_results is the list returned by one worker
        for record in device_results:
            row_data = [
                record.get("adom"),
                record.get("storecode"),
                record.get("interface"),
                record.get("peer"),
                record.get("jitter_avg"),
                record.get("jitter_max"),
                record.get("jitter_min"),
                record.get("latency_avg"),
                record.get("latency_max"),
                record.get("latency_min"),
                record.get("packet_loss"),
                record.get("error")
            ]

            worksheet.append(row_data)

            # Get the packet_loss cell in the current row
            packet_loss_cell = worksheet.cell(
                row=worksheet.max_row,
                # Python list index → 从 0 开始
                # Excel/OpenPyXL column → 从 1 开始
                column=headers.index("packet_loss") + 1
            )

            packet_loss = record.get("packet_loss")
            if packet_loss is not None and packet_loss > 5:
                packet_loss_cell.fill = warning_fill

    # Freeze the header row
    worksheet.freeze_panes = "A2"

    # Enable filters
    worksheet.auto_filter.ref = worksheet.dimensions

    # Set column widths
    column_widths = {
        "A": 18,  # adom
        "B": 18,  # storecode
        "C": 14,  # interface
        "D": 18,  # peer
        "E": 14,  # jitter_avg
        "F": 14,  # jitter_max
        "G": 14,  # jitter_min
        "H": 14,  # latency_avg
        "I": 14,  # latency_max
        "J": 14,  # latency_min
        "K": 14,  # packet_loss
        "L": 35   # error
    }

    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    # Center data except the error column
    for row in worksheet.iter_rows(
        min_row=2,
        max_row=worksheet.max_row,
        min_col=1,
        max_col=len(headers)
    ):
        for cell in row:
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

    # Align error messages to the left
    for cell in worksheet["L"][1:]:
        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

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
    #pprint(results)

    # write to Excel
    output_file = write_to_excel(
        results,
        user_input,
        userfilename=f"{user_input}_LinkQuality.xlsx"
    )

    print(f"Excel file created: {output_file}")


if __name__ == "__main__":
    main()
