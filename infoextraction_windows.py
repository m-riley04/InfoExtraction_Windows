import os, sys, shutil, getpass, wmi, math, winreg, threading, datetime; import platform as pl; from win32 import win32api; from win32.lib import win32con; import tkinter as tk; from tkinter import ttk
def main():

    ###################################################
    ###########        CONFIGURATION       ############
    ###################################################
    
    # Username
    username = getpass.getuser()

    #-- File extraction?
    fileExtraction = False
    #-- Extraction target
    target = "findthisfile.txt"
    #-- Extraction Directories (samples)
    osDir = "C:/"
    usersDir = osDir + "Users/"
    mainUserDir = usersDir + f"{username}/"
    desktopDir = mainUserDir + "Desktop/"
    #-- Target Directory
    targetDir = mainUserDir

    ###################################################
    ###################################################
    ###################################################

    #--------------------------------------------------------------Functions
    def convert_size(size_bytes):
        if size_bytes == 0:
            return "0B"
        size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
        i = int(math.floor(math.log(size_bytes, 1024)))
        p = math.pow(1024, i)
        s = round(size_bytes / p, 2)
        return "%s %s" % (s, size_name[i])

    def windows_info():
        firstLine = "WINDOWS:\n"
        try:
            computer = wmi.WMI()
            computer_info = computer.Win32_ComputerSystem()[0]
            os_info = f"{computer.Win32_OperatingSystem()[0]}\n"
            proc_info = f"{computer.Win32_Processor()[0]}\n"
            gpu_info = f"{computer.Win32_VideoController()[0]}\n"
            ram_bytes = int(computer_info.TotalPhysicalMemory)
            ram_gigs = f"{convert_size(ram_bytes)}\n"
            processes = f"{computer.Win32_Process}"
        except:
            computer_info = ""
            os_info = ""
            proc_info = ""
            gpu_info = ""
            ram_bytes = ""
            ram_gigs = ""
            processes = ""
        return computer_info + os_info + proc_info + gpu_info + ram_gigs

    def other_info():
        username = getpass.getuser()
        arch = pl.architecture()
        machine = pl.machine()
        networkName = pl.node()
        plat = pl.platform()
        processor = pl.processor()
        pybuild = pl.python_build()
        system = pl.system()
        system_release = pl.release()
        system_version = pl.version()

    class App(threading.Thread):

        def __init__(self):
            threading.Thread.__init__(self)
            self.start()

        def callback(self):
            self.root.quit()

        def run(self):
            self.root = tk.Tk()
            self.root.title = "Processing..."
            self.root.geometry("300x125")
            self.root.grid_anchor("center")
            self.root.protocol("WM_DELETE_WINDOW", self.callback)

            self.pb = ttk.Progressbar(self.root, orient="horizontal", mode="determinate", length=250)
            self.pb.grid(column=0, row=0)
            self.pb.step()

            #self.root.overrideredirect(1)

            self.root.mainloop()
        
        def set_complete(self):
            self.pb.configure(value=100)

    def update_progress(value=1):
        try:
            progressBar.pb.step(value)
        except:
            pass

    # Start
    print("Grabbing info...")
    
    # Create Log File
    flashDir = os.getcwd()
    rootFolder = "info_extractions"
    userFolder = rootFolder + f"/{username}_extraction_data"
    fileFolder = userFolder + f"/{username}_extraction_files"
    logFolder = userFolder + f"/{username}_extraction_logs"
    logName = logFolder + f"/{username}_extraction_log_{datetime.date.today()}"
    print("Grabbed username")
    if not os.path.exists(rootFolder):
        os.mkdir(rootFolder)
    if not os.path.exists(userFolder):
        os.mkdir(userFolder)
    if not os.path.exists(fileFolder):
        os.mkdir(fileFolder)
    if not os.path.exists(logFolder):
        os.mkdir(logFolder)
    if not os.path.exists(logName + ".txt"):
        output = open(logName + ".txt", "a")
    else:
        i = 0
        origName = logName
        while os.path.exists(logName + ".txt"):
            logName = origName
            i += 1
            logName += f" ({i})"
        logName = origName + f" ({i})"
        output = open(logName + ".txt", "a")
    
    print("Created file")

    #--------------------------------------------------------------Log File Header
    print("========================================================", file=output)
    print(f"Extraction Date: {datetime.datetime.now()}", file=output)
    print(f"Name: (name here)", file=output)
    print(f"Username: {username}", file=output)
    print("========================================================", file=output)

    #--------------------------------------------------------------Make GUI
    progressBar = App()
    print("Progress bar created")

    #--------------------------------------------------------------WMI Setup
    c = wmi.WMI ()
    update_progress(8)
    print("WMI instance initiated")

    #--------------------------------------------------------------Hardware
    print("\n\n==============Hardware==============", file=output)
    computer_info = c.Win32_ComputerSystem()[0]
    os_info = f"{c.Win32_OperatingSystem()[0]}\n"
    proc_info = f"{c.Win32_Processor()[0]}\n"
    gpu_info = f"{c.Win32_VideoController()[0]}\n"
    ram_bytes = int(computer_info.TotalPhysicalMemory)
    ram_gigs = f"{convert_size(ram_bytes)}\n"
    processes = f"{c.Win32_Process}"
    # GPU
    print("\n---GPU---", file=output)
    print(gpu_info, file=output)
    update_progress()
    print("GPU found")
    # CPU
    print("\n---CPU---", file=output)
    print(proc_info, file=output)
    update_progress()
    print("CPU found")
    # Ram
    print("\n---RAM---", file=output)
    print(ram_gigs, file=output)
    update_progress()
    print("RAM found")
    # Operating System
    print("\n---Operating System---", file=output)
    print(os_info, file=output)
    update_progress()
    print("OS found")
    # Computer System
    print("\n---Computer System---", file=output)
    print(computer_info, file=output)
    update_progress()
    print("Computer System found")

    #--------------------------------------------------------------Drives
    print("\n\n==============Directories/Drives==============", file=output)
    # Find drive types
    print("\n---Drive Types---", file=output)
    DRIVE_TYPES = {
        0 : "Unknown",
        1 : "No Root Directory",
        2 : "Removable Disk",
        3 : "Local Disk",
        4 : "Network Drive",
        5 : "Compact Disc",
        6 : "RAM Disk"
        }
    for drive in c.Win32_LogicalDisk ():
        update_progress(.02)
        print (drive.Caption, DRIVE_TYPES[drive.DriveType], file=output)

    update_progress(5)
    print("Drives found")

    # Show partitions
    print("\n---Partitions---", file=output)
    for physical_disk in c.Win32_DiskDrive ():
        update_progress(.02)
        for partition in physical_disk.associators ("Win32_DiskDriveToDiskPartition"):
            update_progress(.02)
            for logical_disk in partition.associators ("Win32_LogicalDiskToPartition"):
                update_progress(.02)
                print( physical_disk.Caption, partition.Caption, logical_disk.Caption, file=output)
    
    update_progress(5)
    print("Partitions found")

    # Show Storage On Disks
    print("\n---Storage---", file=output)
    for disk in c.Win32_LogicalDisk (DriveType=3):
        update_progress(.02)
        print (disk.Caption, "%0.2f%% free" % (100.0 * int(disk.FreeSpace) / int(disk.Size)), file=output)

    update_progress(5)
    print("Storage amounts found")

    # Show shared drives
    print("\n---Shared Drives---", file=output)
    for share in c.Win32_Share ():
        update_progress(.02)
        print (share.Name, share.Path, file=output)

    update_progress(5)
    print("Shared drives found")

    #--------------------------------------------------------------Apps and Services
    print("\n\n==============Apps and Services==============", file=output)
    # Running Processes
    print("\n---Currently Running Processes---", file=output)
    for process in c.Win32_Process():
        update_progress(.07)
        print(process.ProcessId, process.Name, file=output)

    update_progress(5)
    print("Current running processes found")

    # See startup apps
    print("\n---Startup Applications---", file=output)
    for s in c.Win32_StartupCommand ():
        update_progress(.07)
        print("[%s] %s <%s>" % (s.Location, s.Caption, s.Command), file=output)
    
    update_progress(5)
    print("Startup apps found")

    # Stopped Auto-Services
    print("\n---Stopped AutoServices---", file=output)
    stopped_services = c.Win32_Service (StartMode="Auto", State="Stopped")
    if stopped_services:
        for s in stopped_services:
            update_progress(.07)
            print(s.Caption, "service is not running", file=output)
    else:
        print("No auto services stopped", file=output)

    update_progress(5)
    print("Disabled auto-services found")
    
    # Install a product
    '''
    c.Win32_Product.Install (
        PackageLocation="c:/temp/python-2.4.2.msi",
        AllUsers=False
        )
    '''

    update_progress(5)
    print("Product installed")
    
    #--------------------------------------------------------------Network
    print("\n\n==============Network==============", file=output)
    #IP Addresses and Mac Addresses
    print("\n---IP and Mac Addresses---", file=output)
    for interface in c.Win32_NetworkAdapterConfiguration (IPEnabled=1):
        print (interface.Description, interface.MACAddress, file=output)
        update_progress(.07)
        for ip_address in interface.IPAddress:
            print (ip_address, file=output)
            update_progress(.07)
        print(interface, file=output)
        
    update_progress(5)
    print("IP/Mac addresses found")


    #--------------------------------------------------------------Printing And External Devices
    print("\n\n==============Printing and Devices==============", file=output)
    # Current print jobs
    try:
        print("\n---Current Print Jobs---", file=output)
        for printer in c.Win32_Printer ():
            print (printer.Caption, file=output)
            update_progress(.07)
            for job in c.Win32_PrintJob (DriverName=printer.DriverName):
                print ("  ", job.Document, file=output)
                update_progress(.07)
            print(printer, file=output)
    except:
        print("Failed to find some or all printers.", file=output)

    update_progress(10)
    print("Current print jobs and printers found")
    
    # Watch for prints
    '''
    print_job_watcher = c.Win32_PrintJob.watch_for (notification_type="Creation", delay_secs=1 , file=output)

    while 1:
        pj = print_job_watcher ()
        print ("User %s has submitted %d pages to printer %s" % (pj.Owner, pj.TotalPages, pj.Name), file=output)
    '''

    #update_progress(10)
    #print("Finished watching for prints")

    #--------------------------------------------------------------Registry Entries
    print("\n\n==============Registry==============", file=output)
    # Print all registry keys
    try:
        print("\n---All Registry Keys---", file=output)
        r = wmi.WMI(namespace="DEFAULT").StdRegProv
        result, names = r.EnumKey(hDefKey=winreg.HKEY_LOCAL_MACHINE, sSubKeyName="Software")
        for key in names:
                update_progress(.01)
                print(key, file=output)
        print("Registry keys found")
        update_progress(15)
    except:
        print("Failed to find some or all registry keys.", file=output)

    # Add new registry value
    '''
    result, = r.SetStringValue (
        hDefKey=_winreg.HKEY_LOCAL_MACHINE,
        sSubKeyName=r"Software\TJG",
        sValueName="ApplicationName",
        sValue="TJG App"
        )
    update_progress(5)
    print("Added registry value")
    '''

    #--------------------------------------------------------------Misc
    print("\n\n==============Misc==============", file=output)
    # Get Background
    print("\n---Desktop Background(s)---", file=output)
    full_username = win32api.GetUserNameEx (win32con.NameSamCompatible)
    for desktop in c.Win32_Desktop (Name=full_username):
        print (desktop.Wallpaper or "[No Wallpaper]", desktop.WallpaperStretched, desktop.WallpaperTiled, file=output)
    
    update_progress(5)
    print("Desktop background(s) found")

    # Get Username
    print("\n----Username----", file=output)
    print(full_username, file=output)

    update_progress(5)
    print("Username found")

    
    #--------------------------------------------------------------Find Specific Files
    if fileExtraction:
        print("\n\n==============File Search==============", file=output)
        def find_files(filename, search_path):
            result = []
            # Walking top-down from the root
            for root, dir, files in os.walk(search_path):
                if filename in files:
                    result.append(os.path.join(root, filename))
            return result

        def copy_file(filepath):
            file = os.path.split(filepath)[1]
            fileTup = os.path.splitext(file)                                                    # Current file name and extension in tuple (name, extension)
            fileName = fileTup[0]                                                               # Current file name
            fileExt = fileTup[1]                                                                # Current file extension
            #fileCurrentPath = homeDir + f"/{fileName}{fileExt}"                                 # Current file path
            destination = flashDir + f"/{fileFolder}/{fileName}"
            destination = check_for_repeats(destination, fileExt)       # check for repeats and return correct directory
            destination += fileExt                                      # add extension
            shutil.copy2(filepath, destination)

        def check_for_repeats(destination, extension):
            '''Checks the destination path of a file to see if the a file of the same name already exists. If it does, " repeat(#)" is added to the end. Returns the modified directory.'''
            repeats = 1                                                                     # sets repeat value to base "1"
            if os.path.exists(destination + extension):                             # if the file already exists in the destination folder...
                    destination += f" repeat({repeats})"                            # add "repeat(1)" to the first dupe
                    if os.path.exists(destination + extension):                     # if it STILL exists...
                        while os.path.exists(destination + extension):              # and WHILE it still exists...
                            destination = (destination[::-1].replace(f"({repeats})"[::-1], "", 1))[::-1] # remove the number and parenthesis from the end of the file name
                            repeats += 1                                                    # increment the repeat value
                            destination += f"({repeats})"                           # add (#) with the incremented value
            return destination                                                      # return the end result

        print("\n\n---Files Found---", file=output)
        try:
            print(f"Trying to find '{target}'...")
            found = 0
            for file in find_files(target, targetDir):
                found += 1
                copy_file(file)
                print(f"'{target}' found at '{file}'.")
                print(f"File Copied: {file}", file=output)
            
        except:
            print(f"Failed to search for {target}.", file=output)
        else:
            print(f"Finished searching for files. Found {found} files with name '{target}' in {targetDir}")
            print(f"Found {found} files with name '{target}' in {targetDir}", file=output)


    #--------------------------------------------------------------Close GUI
    print("Process complete. You may now close the program.\n")
    progressBar.set_complete()
    progressBar.callback()
    exit()  
    
main()