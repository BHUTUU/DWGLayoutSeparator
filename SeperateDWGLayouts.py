import win32com.client,os, shutil, time
from tkinter import messagebox, filedialog
from tkinter import *

class SeperateDWGLayouts:
    numberOfLayouts = []
    @classmethod
    def getNumberOfLayouts(cls, path_to_source_dwg):
        layoutarray = []
        try:
            acad = win32com.client.Dispatch("AutoCAD.Application")
            try:
                acad.Visible = True
            except AttributeError:
                time.sleep(2)
                acad.Visible = True
            doc = acad.Documents.Open(path_to_source_dwg)
            time.sleep(5)
            layouts = doc.Layouts
            cls.numberOfLayouts = len(layouts) - 1 #model space ko v count kr raha hai to remove kr diya 1.
            for layout in layouts:
                layoutarray.append(layout.Name)
            layoutarray.remove("Model")
            doc.Close()
            return (cls.numberOfLayouts, layoutarray)
        except Exception:
            return [False, "Could not connect to AutoCAD."]
    @classmethod
    def deleteAllLayoutsExceptIndex(cls,filePath, fileName: str, index: str):
        acad = win32com.client.Dispatch("AutoCAD.Application")
        try:
            acad.Visible = True
        except AttributeError:
            time.sleep(2)
            acad.Visible = True
        doc = acad.Documents.Open(filePath)
        time.sleep(5)
        while True:
            try:
                layouts = doc.Layouts
                for layout in layouts:
                    layoutName = str(layout.Name)
                    if layoutName == fileName or layoutName == index or layoutName == "Model":
                        pass
                    else:
                        layout.Delete()
                        doc.Save()
                layouts = doc.Layouts
                for layout in layouts:
                    layoutName = str(layout.Name)
                    if layoutName != "Model":
                        layout.Name = fileName
                        doc.save()
                doc.Save()
                doc.Close()
                break
            except Exception:
                continue
        return [True, "Deleted uneccesary layout tab"]
    @staticmethod
    def doSeparate(path_to_source_dwg, path_to_dir_for_generated_dwg):
        if not os.path.isfile(path_to_source_dwg):
            return [False, "The path to the source dwg is not valid."]
        if not os.path.isdir(path_to_dir_for_generated_dwg):
            return [False, "The path to the dir for generated dwg is not valid."]
        layoutNumberAndName = SeperateDWGLayouts.getNumberOfLayouts(path_to_source_dwg)
        if layoutNumberAndName[0] == False:
            return layoutNumberAndName
        numberOfLayouts = layoutNumberAndName[0]
        layoutarray = layoutNumberAndName[1]
        baseNameOfDwg = str(os.path.basename(path_to_source_dwg))[:-5]
        # initialForRange=1
        errorInFiles=[]
        # if os.path.basename(path_to_source_dwg) in os.listdir(path_to_dir_for_generated_dwg):
        #     initialForRange=2
        for i in range(int(numberOfLayouts)):
            layoutName = layoutarray[i]
            tempfilehold  = os.path.join(path_to_dir_for_generated_dwg, f"{baseNameOfDwg}{layoutName}.dwg")
            if not os.path.isfile(tempfilehold):
                shutil.copy(path_to_source_dwg, tempfilehold)
                delResponse = SeperateDWGLayouts.deleteAllLayoutsExceptIndex(str(tempfilehold), f"{baseNameOfDwg}{layoutName}", str(layoutName))
                if delResponse[0] == False:
                    errorInFiles.append(tempfilehold+": error while deleting layouts for this file")
            else:
                errorInFiles.append(tempfilehold+": This file already exists at destination so i will not even delete layouts from it.")
        return [True, errorInFiles]
window = Tk()

window.title("Seperate Layouts")
folderPathtk=StringVar()
filePathtk=StringVar()
def getFilePath():
    global filePathtk
    filePath = filedialog.askopenfilename(title="Select Drawing", filetypes=(("Drawing files", "*.dwg"), ("All files", "*.*")))
    if filePath:
        filePathtk.set(filePath)
        fileLabel.configure(text=filePath)
def getDestinationFolder():
    global folderPathtk
    folderPath = filedialog.askdirectory()
    if folderPath:
        folderPathtk.set(folderPath)
        folderLabel.configure(text=folderPath)
def getDesktopPath():
    shell = win32com.client.Dispatch("WScript.Shell")
    desktopPathRaw = shell.SpecialFolders("Desktop")
    desktopPath=desktopPathRaw.replace("\\","/")
    return desktopPath
def getSystem32Path():
    system32path = os.path.join(os.getenv("systemroot"), "System32").replace("\\", "/")
    return system32path
def runner():
    global filePathtk, folderPathtk
    s_file = filePathtk.get()
    s_folder = folderPathtk.get()
    if os.path.exists(s_file):
        if os.path.exists(s_folder):
            s_ins = SeperateDWGLayouts()
            if s_file.split("\\")[:-1] == s_folder.split("\\"):
                messagebox.showerror("Error", "Please select different folder for the output")
                return False
            if s_folder == str(getDesktopPath()) or s_folder.lower() == str(getSystem32Path()).lower():
                messagebox.showerror("Error", "Output folder may contains system files. Please select different path for the output.")
                messagebox.showinfo("Precaution", "Desktop and system32 is not to be used as output folder!")
                return False
            try:
                resp = s_ins.doSeparate(s_file, s_folder)
                messagebox.showinfo("Success", f"Response: {resp}")
            except:
                messagebox.showerror("Error", "Failed! delete if any tab seperated and try again.")
                return False
        else:
            messagebox.showerror("Error", "Please select a valid folder then try RUN")
            return False
    else:
        messagebox.showerror("Error", "Please select a valid file then try RUN")
        return False
f1=Frame(window)
fileButton = Button(f1, text="Select DWG file", command=getFilePath, bg="pink").pack(side="left", padx=10)
fileLabel = Label(f1, text="Select the DWG file")
fileLabel.pack(side="right", padx=10)
fileLabel.bind('<Configure>', lambda e: fileLabel.config(wraplength=fileLabel.winfo_width()))
f1.pack(fill="x", padx=10, pady=5)
f2=Frame(window)
folderButton = Button(f2, text="Select Output Folder", command=getDestinationFolder, bg="pink").pack(side="left", padx=10)
folderLabel = Label(f2, text="Select the Output folder")
folderLabel.pack(side="right", padx=10)
f2.pack(fill="x", padx=10, pady=5)
submitButton = Button(window, text="Run",bg="blue", command=runner).pack(fill="both", padx=5,pady=5)
window.mainloop()