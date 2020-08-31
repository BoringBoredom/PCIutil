import ctypes, sys

if ctypes.windll.shell32.IsUserAnAdmin() == False:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

import winreg, webbrowser, requests, os, json

current_version = 0.12
os.system('mode 300, 1000')
ctypes.windll.kernel32.SetConsoleTitleW("PCIutil " + str(current_version))
user32 = ctypes.WinDLL('user32')
user32.ShowWindow(user32.GetForegroundWindow(), 3)

try:
    r = requests.get("https://api.github.com/repos/BoringBoredom/PCIutil/releases/latest")
    new_version = float(r.json()["tag_name"])
    if new_version > current_version:
        webbrowser.open("https://github.com/BoringBoredom/PCIutil/releases/latest")
except:
    pass

path = r"SYSTEM\CurrentControlSet\Enum\PCI"
affinity_path = "\\Device Parameters\\Interrupt Management\\Affinity Policy"
msi_path = "\\Device Parameters\\Interrupt Management\\MessageSignaledInterruptProperties"
affinity_policies = {1: "IrqPolicyAllCloseProcessors", 2: "IrqPolicyOneCloseProcessor", 3: "IrqPolicyAllProcessorsInMachine", 4: "IrqPolicySpecifiedProcessors", 5: "IrqPolicySpreadMessagesAcrossAllProcessors", "-": "IrqPolicyMachineDefault"}
irq_priorities = {1: "Low", 2: "Normal", 3: "High", "-": "Undefined"}
msi = {1: "On", 0: "Off", "-": "-"}
value_types = {"REG_DWORD": 4, "REG_BINARY": 3}
message_content = ""
export_lock = False

def message(message):
    global message_content
    message_content = message

def create_registry_keys(path):
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path + "\\Device Parameters", 0, winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY) as key:
        winreg.CreateKeyEx(key, "Interrupt Management", 0, winreg.KEY_WRITE)
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path + "\\Device Parameters\\Interrupt Management", 0, winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY) as key2:
            winreg.CreateKeyEx(key2, "Affinity Policy", 0, winreg.KEY_WRITE)
            winreg.CreateKeyEx(key2, "MessagesignaledInterruptProperties", 0, winreg.KEY_WRITE)

def read_value(path, value_name):
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
        try:
            if winreg.QueryValueEx(key, value_name)[1] == 3:
                try:
                    return int.from_bytes(winreg.QueryValueEx(key, value_name)[0], "little")
                except TypeError:
                    return "?"
            return winreg.QueryValueEx(key, value_name)[0]
        except FileNotFoundError:
            return "-"

def write_value(path, value_name, value_type, value):
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as key:
        winreg.SetValueEx(key, value_name, 0, value_types[value_type], value)

def delete_value(path, value_name):
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as key:
        try:
            winreg.DeleteValue(key, value_name)
        except FileNotFoundError:
            pass

def fetch_devices():
    devices = []
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
        for i in range(winreg.QueryInfoKey(key)[0]):
            device = {}
            device["Hardware ID"] = winreg.EnumKey(key, i)
            path2 = path + "\\" + winreg.EnumKey(key, i)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path2, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key2:
                path3 = path2 + "\\" + winreg.EnumKey(key2, 0)
                if read_value(path3, "ConfigFlags") == 0:
                    device["Path"] = path3
                    create_registry_keys(device["Path"])
                    device["DeviceDesc"] = read_value(path3, "DeviceDesc").split(";")[1]
                    device["DevicePriority"] = read_value(path3 + affinity_path, "DevicePriority")
                    device["DevicePolicy"] = read_value(path3 + affinity_path, "DevicePolicy")
                    device["AssignmentSetOverride"] = read_value(path3 + affinity_path, "AssignmentSetOverride")
                    device["MessageNumberLimit"] = read_value(path3 + msi_path, "MessageNumberLimit")
                    device["MSISupported"] = read_value(path3 + msi_path, "MSISupported")
                    devices.append(device)
    return devices

def convert_affinities(value):
    if value == "-":
        return "-"
    value = bin(value)
    current_cpu = 0
    reversed_binary_string = value.split("b")[1][::-1]
    cpu_list = ""
    for char in reversed_binary_string:
        if char == "1":
            cpu_list += str(current_cpu)
            if current_cpu < len(reversed_binary_string) - 1:
                cpu_list += ", "
        current_cpu += 1
    return cpu_list

def print_device_information():
    global export_lock
    max_msi_length, max_devprio_length, max_messagelimit_length, max_affinitypolicy_length = 0, 0, 0, 0
    for device in devices:
        for option, value in device.items():
            if option == "MSISupported":
                try:
                    length = len(msi[value])
                except KeyError:
                    length = 1
                    export_lock = True
                if length > max_msi_length:
                    max_msi_length = length
            elif option == "DevicePriority":
                try:
                    length = len(irq_priorities[value])
                except KeyError:
                    length = 1
                    export_lock = True
                if length > max_devprio_length:
                    max_devprio_length = length
            elif option == "MessageNumberLimit":
                length = len(str(value))
                if length > max_messagelimit_length:
                    max_messagelimit_length = length
            elif option == "DevicePolicy":
                try:
                    length = len(affinity_policies[value])
                except KeyError:
                    length = 1
                    export_lock = True
                if length > max_affinitypolicy_length:
                    max_affinitypolicy_length = length
    max_index_length = len(str(len(devices) - 1))
    for device in devices:
        string = "\n" + (max_index_length - len(str(devices.index(device))))*" " + str(devices.index(device)) + ". " + device['DeviceDesc'] + "\n\n" + (max_index_length + 2)*" "
        string += "MSI: "
        try:
            string += msi[device["MSISupported"]] + (6 + max_msi_length - len(msi[device["MSISupported"]]))*" "
        except KeyError:
            string += "?" + (5 + max_msi_length)*" "
        string += "MSG Limit: " + str(device["MessageNumberLimit"]) + (6 + max_messagelimit_length - len(str(device["MessageNumberLimit"])))*" "
        string += "IRQ Priority: "
        try:
            string += irq_priorities[device["DevicePriority"]] + (6 + max_devprio_length - len(irq_priorities[device["DevicePriority"]]))*" "
        except KeyError:
            string += "?" + (5 + max_devprio_length)*" "
        string += "IRQ Policy: "
        try:
            string += affinity_policies[device["DevicePolicy"]] + (6 + max_affinitypolicy_length - len(affinity_policies[device["DevicePolicy"]]))*" "
        except KeyError:
            string += "?" + (5 + max_affinitypolicy_length)*" "
        string += "CPUs: "
        try:
            string += convert_affinities(device["AssignmentSetOverride"])
        except TypeError:
            string += "?"
        print(string)

def device_check(device_selection):
    try:
        for device in device_selection:
            device = int(device)
            if device > len(devices) - 1:
                message("Device doesn't exist.")
                return False
        return True
    except ValueError:
        message("Devices must be integers.")
        return False

def all_devices_selection():
    string = ""
    for i in range(len(devices)):
        string += str(i)
        if i < len(devices) - 1:
            string += " "
    return string

def change_msi():
    temp = {"0": "OFF", "1": "ON", "2": "DELETED"}
    option = input("Change MSI: 0 = OFF, 1 = ON, 2 = DELETED: ")
    if option not in temp:
        message("Invalid input. Only 0, 1 or 2 possible.")
        return
    device_selection = input(f"Change MSI to {temp[option]} for which devices?: ")
    if device_selection == "all":
        device_selection = all_devices_selection()
    device_selection = device_selection.split(" ")
    if device_check(device_selection) == False:
        return
    for device in device_selection:
        device = int(device)
        if option == "2":
            delete_value(devices[device]["Path"] + msi_path, "MSISupported")
        else:
            write_value(devices[device]["Path"] + msi_path, "MSISupported", "REG_DWORD", int(option))

def change_message_limit():
    try:
        limit = input("Limit (0 = no limit): ")
        limit = int(limit)
    except ValueError:
        message("Limit must be an integer.")
        return
    device_selection = input(f"Change the message limit to {str(limit)} for which devices?: ")
    if device_selection == "all":
        device_selection = all_devices_selection()
    device_selection = device_selection.split(" ")
    if device_check(device_selection) == False:
        return
    for device in device_selection:
        device = int(device)
        if limit == 0:
            delete_value(devices[device]["Path"] + msi_path, "MessageNumberLimit")
        else:
            write_value(devices[device]["Path"] + msi_path, "MessageNumberLimit", "REG_DWORD", limit)

def change_interrupt_priority():
    temp = {"0": "Undefined", "1": "Low", "2": "Normal", "3": "High"}
    option = input("Change Interrupt Priority to 0 = Undefined, 1 = Low, 2 = Normal, 3 = High: ")
    if option not in temp:
        message("Invalid input. Only 0, 1, 2 or 3 possible.")
        return
    device_selection = input(f"Change the Interrupt Priority to {temp[option]} for which devices?: ")
    if device_selection == "all":
        device_selection = all_devices_selection()
    device_selection = device_selection.split(" ")
    if device_check(device_selection) == False:
        return
    for device in device_selection:
        device = int(device)
        if option == "0":
            delete_value(devices[device]["Path"] + affinity_path, "DevicePriority")
        else:
            write_value(devices[device]["Path"] + affinity_path, "DevicePriority", "REG_DWORD", int(option))

def change_affinity_policy():
    temp = {"0": "IrqPolicyMachineDefault", "1": "IrqPolicyAllCloseProcessors", "2": "IrqPolicyOneCloseProcessor", "3": "IrqPolicyAllProcessorsInMachine", "5": "IrqPolicySpreadMessagesAcrossAllProcessors"}
    option = input("Change Affinity Policy to 0 = MachineDefault, 1 = AllCloseProcessors, 2 = OneCloseProcessor, 3 = AllProcessorsInMachine, 5 = SpreadMessagesAcrossAllProcessors: ")
    if option not in temp:
        message("Invalid input. Only 0, 1, 2, 3 or 5 possible.")
        return
    device_selection = input(f"Change the Affinity Policy to {temp[option]} for which devices?: ")
    if device_selection == "all":
        device_selection = all_devices_selection()
    device_selection = device_selection.split(" ")
    if device_check(device_selection) == False:
        return
    for device in device_selection:
        device = int(device)
        if option == "0":
            delete_value(devices[device]["Path"] + affinity_path, "DevicePolicy")
        else:
            write_value(devices[device]["Path"] + affinity_path, "DevicePolicy", "REG_DWORD", int(option))
        delete_value(devices[device]["Path"] + affinity_path, "AssignmentSetOverride")

def change_cpu_affinities():
    thread_count = os.cpu_count()
    option = input("Your last thread -> " + thread_count*"1" + " <- your first thread. 0 = no affinity, 1 = affinity.\nEnter the binary string corresponding to your thread count and desired affinities: ")
    for char in option:
        if char not in ["0", "1"]:
            message("Invalid format. Binary string can only consist of 0s and 1s.")
            return
    if len(option) != thread_count:
        message("Invalid format. Entered threads differ from system threads.")
        return
    device_selection = input(f"Change the affinity to CPUs {convert_affinities(int(option, 2))} for which devices?: ")
    if device_selection == "all":
        device_selection = all_devices_selection()
    device_selection = device_selection.split(" ")
    if device_check(device_selection) == False:
        return
    for device in device_selection:
        device = int(device)
        if option == thread_count*"0":
            delete_value(devices[device]["Path"] + affinity_path, "DevicePolicy")
            delete_value(devices[device]["Path"] + affinity_path, "AssignmentSetOverride")
        else:
            write_value(devices[device]["Path"] + affinity_path, "DevicePolicy", "REG_DWORD", 4)
            write_value(devices[device]["Path"] + affinity_path, "AssignmentSetOverride", "REG_BINARY", int(option, 2).to_bytes(8, "little").rstrip(b"\x00"))

def show_hardware_ids():
    ids = ""
    max_device_length = 0
    for device in devices:
        length = len(device["DeviceDesc"])
        if length > max_device_length:
            max_device_length = length
    max_index_length = len(str(len(devices) - 1))
    for device in devices:
        ids += "\n" + (max_index_length - len(str(devices.index(device))))*" " + str(devices.index(device)) + ". " + device['DeviceDesc'] + ": " + (max_device_length - len(device['DeviceDesc']))*" " + device['Hardware ID']
    message(ids)

def show_readme():
    message("Syntax of device selection: 0 1 2 3 4 5 etc. or all. Each device is separated by one space.\n0 3 5   - This executes the selected operation for devices 0, 3 and 5.\nall     - This executes the selected operation for all devices.\n\nBe careful when changing settings of storage drivers. Some settings might BSOD on specific drivers.\n\nexit    - This command exits the program.")

def apply_import(config):
    for device in config:
        for value_name, value in device.items():
            if value_name == "MSISupported":
                if value == "-":
                    delete_value(device["Path"] + msi_path, value_name)
                else:
                    write_value(device["Path"] + msi_path, value_name, "REG_DWORD", value)
            elif value_name == "MessageNumberLimit":
                if value == "-":
                    delete_value(device["Path"] + msi_path, value_name)
                else:
                    write_value(device["Path"] + msi_path, value_name, "REG_DWORD", value)
            elif value_name == "DevicePriority":
                if value == "-":
                    delete_value(device["Path"] + affinity_path, value_name)
                else:
                    write_value(device["Path"] + affinity_path, value_name, "REG_DWORD", value)
            elif value_name == "DevicePolicy":
                if value == "-":
                    delete_value(device["Path"] + affinity_path, value_name)
                else:
                    write_value(device["Path"] + affinity_path, value_name, "REG_DWORD", value)
            elif value_name == "AssignmentSetOverride":
                if value == "-":
                    delete_value(device["Path"] + affinity_path, value_name)
                else:
                    write_value(device["Path"] + affinity_path, "AssignmentSetOverride", "REG_BINARY", value.to_bytes(8, "little").rstrip(b"\x00"))

def import_config():
    file_list = []
    for file in os.listdir():
        if "json" in file:
            file_list.append(file.split(".")[0])
    i = 0
    files = {}
    for file in file_list:
        files[str(i)] = file
        print(f"{i}. {file}")
        i += 1
    if file_list == []:
        message("No configs found.")
        return
    option = input("Import which file?: ")
    if option not in files:
        message("File doesn't exist.")
        return
    with open(f"{files[option]}.json", "r") as f:
        config = json.load(f)
    j = 0
    for device in config:
        if device["Path"] != devices[j]["Path"]:
            message("Config incompatible with currently active devices.")
            return
        j += 1
    apply_import(config)

def config_menu():
    global export_lock
    option = input("Enter 1 to import a config, 2 to export a config: ")
    if option not in ["1", "2"]:
        message("Invalid input. Only 1 or 2 possible.")
        return
    if option == "1":
        import_config()
    elif option == "2":
        if export_lock == True:
            message("Export disallowed due to incompatible registry entries.")
            return
        config_name = input("Enter the config name: ")
        with open(f"{config_name}.json", 'w') as f:
            json.dump(devices, f, indent=4)

def show_suboptions(option_choice):
    if option_choice == "exit":
        sys.exit(0)
    elif option_choice == "0":
        show_readme()
    elif option_choice == "1":
        change_msi()
    elif option_choice == "2":
        change_message_limit()
    elif option_choice == "3":
        change_interrupt_priority()
    elif option_choice == "4":
        change_affinity_policy()
    elif option_choice == "5":
        change_cpu_affinities()
    elif option_choice == "6":
        show_hardware_ids()
    elif option_choice == "7":
        config_menu()
    else:
        message("Invalid input.")


while True:
    export_lock == False
    os.system('cls')
    devices = fetch_devices()
    print_device_information()
    print(f"\n0. Show README\n1. Change MSI\n2. Change Message Limit\n3. Change Interrupt Priority\n4. Change Affinity Policy\n5. Change CPU Affinities\n6. Show Hardware IDs\n7. Import/Export config")
    if message_content != "":
        print("\n" + message_content)
        message_content = ""
    option_choice = input("\n" + "Enter command number: ")
    show_suboptions(option_choice)