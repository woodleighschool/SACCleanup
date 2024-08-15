import subprocess
import os
from pathlib import Path
import shutil
from tkinter import Tk, messagebox


def reset_microsoft_word():
    subprocess.run(['pkill', '-9', 'Microsoft Word'])

    word_paths = [
        "/Users/sac/Library/Preferences/com.microsoft.Word.plist",
        "/Users/sac/Library/Containers/com.microsoft.Word",
        "/Users/sac/Library/Application Scripts/com.microsoft.Word",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/*.dot",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/*.dotx",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/*.dotm",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/mip_policy",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/FontCache",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/ComRPC32",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/TemporaryItems",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/Microsoft Office ACL*",
        "/Users/sac/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB.reg",
    ]

    for path_string in word_paths:
        path = Path(path_string)
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        elif "*" in path_string:
            for file_path in path.parent.glob(path.name):
                file_path.unlink()
        elif path.exists():
            path.unlink()

    print("Office-Reset: Removed configuration data for Microsoft Word")


def erase_and_restrict_user_directories():
    directories_to_erase = ["Desktop", "Documents", "Downloads"]
    base_path = Path("/Users/sac")
    deleted_dir = base_path / "Deleted"

    if not deleted_dir.exists():
        deleted_dir.mkdir(parents=True, exist_ok=True)

    for directory in directories_to_erase:
        full_path = base_path / directory
        for item in full_path.glob('*'):
            target_path = deleted_dir / item.name

            counter = 1
            while target_path.exists():
                target_path = deleted_dir / f"{item.stem} {counter}{item.suffix}"
                counter += 1

            shutil.move(str(item), str(target_path))
            os.chmod(target_path, 0o000)


def dock_item(app_path):
    return f'''<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>{app_path}</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'''


def setup_dock():
    apps = [
        '/Applications/Microsoft Word.app',
        '/Applications/NAP Locked down browser.app'
    ]

    subprocess.run(['defaults', 'delete', 'com.apple.dock', 'persistent-apps'], check=False)
    subprocess.run(['defaults', 'delete', 'com.apple.dock', 'persistent-others'], check=False)

    dock_items = [dock_item(app) for app in apps]
    subprocess.run(["defaults", "write", "com.apple.dock", "persistent-apps", "-array"] + dock_items, check=True)
    subprocess.run(['defaults', 'write', 'com.apple.dock', 'minimize-to-application', '-bool', 'true'], check=False)
    subprocess.run(['defaults', 'write', 'com.apple.dock', 'show-recents', '-bool', 'false'], check=False)

    # Restart the Dock to apply changes
    subprocess.run(['killall', 'Dock'], check=False)


def main():
    root = Tk()
    root.withdraw()
    if messagebox.askokcancel("Reset This Mac", "Running this will reset Microsoft Word, and empty the Desktop, Documents and Downloads.\n*ALL WORK WILL BE LOST!*\nDo you want to continue?"):
        reset_microsoft_word()
        erase_and_restrict_user_directories()
        setup_dock()
    root.destroy()


if __name__ == '__main__':
    main()
