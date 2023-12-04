from pywinauto.application import Application
import openpyxl
import subprocess
import os
import time


def kill_cmd_process():
    """Kill the cmd.exe process if it exists."""
    try:
        # Check if cmd.exe process is running
        cmd_process = subprocess.run("tasklist | find \"cmd.exe\"", shell=True, text=True, capture_output=True)

        if "cmd.exe" in cmd_process.stdout:
            # If cmd.exe is running, attempt to kill it
            subprocess.run("taskkill /F /IM cmd.exe", shell=True, check=True)
            print("Cmd.exe process terminated successfully.")
        else:
            print("No cmd.exe process found.")
    except subprocess.CalledProcessError as e:
        print(f"Error terminating cmd.exe process: {e}")


def save_to_excel():
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Select the sheet name
    Script_Out = workbook.active

    # Configure your sheet
    Script_Out["A1"] = "Output"
    for i, line in enumerate(output_lines, start=2):
        Script_Out[f"A{i}"] = line

    # Save the Excel file
    workbook.save("Terminal Output.xlsx")

if __name__ == "__main__":
    # Start Notepad++ application
    app = Application(backend='uia').start(r"C:\Program Files (x86)\Notepad++\notepad++.exe")
    # Connect to the Notepad++ window
    app = Application(backend='uia').connect(title="new 1 - Notepad++", timeout=100)
    line1 = "Connected to Notepad++"
    print(line1)

    # Maximize the Notepad++ window
    try:
        app.New1Notepad.child_window(title="Maximize", control_type="Button").click_input()
        line2 = "Window maximized"
        print(line2)
    except:
        line2 = "Window already Maximized"
        print(line2)

    # Access the 'File' menu
    file = app.New1Notepad.child_window(title="File", control_type="MenuItem").wrapper_object()
    file.click_input()
    line3 = "File menu opened"
    print(line3)

    # Access the 'Rename...' option in the 'File' menu
    rename = app.New1Notepad.file.child_window(title="Rename...", control_type="MenuItem").wrapper_object()
    rename.click_input()
    line4 = "Opened 'Rename' dialog"
    print(line4)

    # Enter the new file name in the 'New Name:' text box
    name = app.New1Notepad.child_window(title="New Name: ", control_type="Edit").wrapper_object()
    name.type_keys("Automated_Rename", with_spaces=True)
    line5 = "Entered new file name"
    print(line5)

    # Confirm the rename operation
    app.New1Notepad.OK.click()
    line6 = "Renamed the file"
    print(line6)

    # Start a Comment Block
    app.AutomatedRenameNotepad.type_keys("\"\"\"{ENTER}")

    # Access the 'Edit' menu
    edit = app.AutomatedRenameNotepad.child_window(title="Edit", control_type="MenuItem").wrapper_object()
    edit.click_input()
    line7 = "Edit menu opened"
    print(line7)

    # Access the 'Insert' option in the 'Edit' menu
    insert = app.AutomatedRenameNotepad.edit.child_window(title="Insert", control_type="MenuItem").wrapper_object()
    insert.click_input()
    line8 = "Insert menu opened"
    print(line8)

    # Select the 'Date Time (short)' option from the 'Insert' menu
    date = app.AutomatedRenameNotepad.insert.child_window(title="Date Time (short)", control_type="MenuItem").wrapper_object()
    date.click_input()

    # End Comment Block
    app.AutomatedRenameNotepad.type_keys("{ENTER}\"\"\"{ENTER}")
    line9 = "'Date Time (short)' inserted As Comment"
    print(line9)

    # Write Python code
    app.AutomatedRenameNotepad.type_keys('print{(}"Hello{SPACE}From{SPACE}The{SPACE}Bot"{)}{ENTER}')

    # Save as py file named hello in current directory
    app.AutomatedRenameNotepad.type_keys('^%S')
    # Name the file
    file_name = app.AutomatedRenameNotepad.child_window(title="File name:", control_type="Edit").wrapper_object()
    file_name.type_keys("hello.py")
    # Choose the project directory
    file_path =  app.AutomatedRenameNotepad.child_window(title="Previous Locations", control_type="Button").wrapper_object()
    file_path.click_input()
    current_directory = os.getcwd()
    directory_except_last = os.path.dirname(current_directory)
    edit_file_path = app.AutomatedRenameNotepad.child_window(title="Address", control_type="Edit").wrapper_object()
    edit_file_path.type_keys(f"{directory_except_last}\\Notepad{{+}}{{+}}_Automation{{ENTER}}")

    app.AutomatedRenameNotepad.Save.click()
    try:
        app.AutomatedRenameNotepad.Yes.click()
        line10 = "File Replaced"
        print(line10)
    except:
        line10 = "File Created"
        print(line10)

    # Get the current working directory (project folder)
    project_folder = os.getcwd()

    # Command to open cmd in the project folder
    cmd_command = 'start cmd /K "cd /d {}"'.format(project_folder)

    # Open cmd in the project folder
    subprocess.Popen(cmd_command, shell=True)
    time.sleep(5)
    try:
        result = subprocess.run('hello.py', shell=True, text=True, capture_output=True)
        line11 = f"Command output: {result.stdout.strip()}"
        print(line11)

    except subprocess.CalledProcessError as e:
        line11 = f"Error running command: {e}"
        print(line11)

    # Close Notepad++ Application
    app.top_window().CloseButton.click()
    line12 = "Application Closed"
    print(line12)

    output_lines = [line1, line2, line3, line4, line5, line6, line7, line8, line9, line10, line11, line12]

    # Save Output in Excel Sheet
    save_to_excel()

    # Close Cmd
    kill_cmd_process()
