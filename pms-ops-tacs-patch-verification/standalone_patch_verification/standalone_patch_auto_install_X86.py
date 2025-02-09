import os

user_home = os.path.expanduser("~")
desktop = os.path.join(user_home, "Desktop")

sw_patch_list = os.path.join(desktop, "sw_patch_list.xlsx")
ms_patch_list = os.path.join(desktop, "ms_patch_list.xlsx")

patch_folder = os.path.join(user_home, r"AppData\Local\Temp\package\standalone")
