import os

user_home = os.path.expanduser("~")
desktop = os.path.join(user_home, "Desktop")

sw_patch_list = os.path.join(desktop, "sw_patch_list.xlsx")
ms_patch_list = os.path.join(desktop, "ms_patch_list.xlsx")

patch_folder = os.path.join(user_home, r"AppData\Local\Temp\package\standalone")


def find_file_in_folder(base_folder, target_name):
    for root, dirs, files in os.walk(base_folder):
        for file in files:
            if (
                file.lower() == target_name.lower()
                or file.lower() == f"{target_name.lower()}.exe"
            ):
                return os.path.join(root, file)
    return None
