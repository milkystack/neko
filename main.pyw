import win32gui
import win32process
import time
import json
import os
import threading
import psutil
import tkinter as tk
import ctypes
from ctypes import wintypes
from collections import namedtuple
import ctypes
current_path = os.path.dirname(os.path.realpath(__file__))

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass

SelectBox = ""
AppNameLabel = ""
NewNameInput = ""
SetButton = ""
UnSetButton = ""

def main():
    global SelectBox
    global NewNameInput
    global AppNameLabel
    global SetButton
    global UnSetButton

    root = tk.Tk()
    root.title(u"Editer Host")
    root.geometry("900x600")

    root.columnconfigure(0, weight=1)
    root.columnconfigure(1, weight=2)
    root.rowconfigure(0, weight=1)

    SelectBox = tk.Listbox(
        root,
        bg="#161616",
        fg="white"
    )
    SelectBox.grid(
        column=0,
        row=0,
        ipadx=10,
        ipady=10,
        sticky="NSEW"
    )
    SelectBox.bind('<<ListboxSelect>>', select_action)

    UpdateButton = tk.Button(
        root,
        bg="#161616", 
        fg="white",
        text="Reload",
        command=Update_SelectBox
    )
    UpdateButton.grid(
        column=0,
        row=1,
        ipadx=10,
        ipady=10,
        sticky="NSEW",
    )


    EditFrame = tk.Frame(
        root,
        bg="#161616", 
    )
    EditFrame.grid(
        column=1,
        row=0,
        sticky="NSEW",
        rowspan=2
    )
    EditFrame.columnconfigure(0, weight=1)
    EditFrame.rowconfigure(1, weight=1)


    ## AppNameLabel =====================
    AppNameLabel = tk.Label(
        EditFrame,
        text="AppName",
        bg="#202020", 
        fg="white"
    )
    AppNameLabel.grid(
        column=0,
        row=0,
        ipadx=10,
        ipady=10,
        sticky="NEW"
    )

    ## NewNameInput =====================
    NewNameInput = tk.Entry(
        EditFrame,
        bg="#161616", 
        fg="white"
    )
    NewNameInput.grid(
        column=0,
        row=1,
        ipadx=10,
        ipady=10,
        sticky="NEW",
    )
    
    ## Button Frame =====================
    ButtonFrame = tk.Frame(
        EditFrame,
        bg="#161616", 
    )
    ButtonFrame.grid(
        column=0,
        row=2,
        sticky="NSEW",
    )
    ButtonFrame.columnconfigure(0, weight=1)
    ButtonFrame.columnconfigure(1, weight=1)

    ## Button ===========================
    SetButton = tk.Button(
        ButtonFrame,
        bg="#161616", 
        fg="white",
        text="Set",
        command=NameSet
    )
    SetButton.grid(
        column=1,
        row=0,
        ipadx=10,
        ipady=10,
        sticky="NSEW",
    )
    UnSetButton = tk.Button(
        ButtonFrame,
        bg="#161616", 
        fg="white",
        text="UnSet",
        command=NameUnSet
    )
    UnSetButton.grid(
        column=0,
        row=0,
        ipadx=10,
        ipady=10,
        sticky="NSEW",
    )

    def getLoop():
        try:
            for elem in list_windows():
                if elem[1] != "":
                    # reject =================
                    windowHandle = win32gui.FindWindow(None, elem[1]) 
                    rect = win32gui.GetWindowRect(windowHandle)
                    x = rect[0]  
                    y = rect[1]
                    if x < 2 and y < 2:
                        continue
                    if psutil.Process(elem[0]).name() == "ApplicationFrameHost.exe":
                        continue
                    # ========================
                    with open(current_path+'/data.json', 'r', encoding='utf-8') as f:
                        json_dict = json.load(f)
                        try:
                            # print("b = "+elem[1])
                            # print("a = "+json_dict[psutil.Process(elem[0]).name()])
                            if elem[1] != json_dict[psutil.Process(elem[0]).name()]:
                                # UpdateButton["text"] = "Reload - "+elem[1]
                                index = SelectBox.get(0, "end").index(psutil.Process(elem[0]).name())
                                SelectBox.itemconfig(index, {'bg':'red'})
                                windowHandle = win32gui.FindWindow(None, elem[1]) 
                                win32gui.SetWindowText(windowHandle, json_dict[psutil.Process(elem[0]).name()])
                                def markup_0(index):
                                    time.sleep(1)
                                    SelectBox.itemconfig(index, {'bg':'#161616'})
                                thread1 = threading.Thread(target=markup_0, args=(index,))
                                thread1.start()
                        except:
                            pass
        except:
            pass
        root.after(1000, getLoop)
    root.after(1000, getLoop)

    Update_SelectBox()
    root.mainloop()

def NameSet():
    global SetButton
    global UnSetButton
    if NewNameInput.get() == "未設定":
        return
    with open(current_path+'/data.json', 'r', encoding='utf-8') as f:
        json_dict = json.load(f)
        json_dict[str(AppNameLabel["text"])] = NewNameInput.get()
    with open(current_path+'/data.json', 'w', encoding='utf-8') as f:
        json.dump(json_dict, f, indent=4, ensure_ascii=False)
    SetButton["state"] = "disable"
    UnSetButton["state"] = "normal"

def NameUnSet():
    global SetButton
    global UnSetButton
    with open(current_path+'/data.json', 'r', encoding='utf-8') as f:
        json_dict = json.load(f)
        del json_dict[str(AppNameLabel["text"])]
    with open(current_path+'/data.json', 'w', encoding='utf-8') as f:
        json.dump(json_dict, f, indent=4, ensure_ascii=False)
    SetButton["state"] = "normal"
    UnSetButton["state"] = "disable"


def Update_SelectBox():
    SelectBox.delete(0, tk.END)
    for elem in list_windows():
        if elem[1] != "":
            # reject =================
            windowHandle = win32gui.FindWindow(None, elem[1]) 
            rect = win32gui.GetWindowRect(windowHandle)
            x = rect[0]  
            y = rect[1]
            if x < 2 and y < 2:
                continue
            if psutil.Process(elem[0]).name() == "ApplicationFrameHost.exe":
                continue
            # ========================
            SelectBox.insert(tk.END, psutil.Process(elem[0]).name())


def select_action(event):
    global NewNameInput
    global AppNameLabel
    # print(select_now(event))
    NewNameInput.delete(0, tk.END)
    with open(current_path+'/data.json', 'r', encoding='utf-8') as f:
        json_dict = json.load(f)
        try:
            NewNameInput.insert(tk.END, json_dict[select_now(event)])
            SetButton["state"] = "disable"
            UnSetButton["state"] = "normal"
        except:
            NewNameInput.insert(tk.END, "未設定")
            SetButton["state"] = "normal"
            UnSetButton["state"] = "disable"
    AppNameLabel["text"] = select_now(event)
    
## =========================================================================================
def list_windows():
    user32 = ctypes.windll.user32
    WindowInfo = namedtuple('WindowInfo', 'pid title')
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM,)
    result = []
    def enum_proc(hWnd, lParam):
        if user32.IsWindowVisible(hWnd):
            pid = wintypes.DWORD()
            tid = user32.GetWindowThreadProcessId(hWnd, ctypes.byref(pid))
            length = user32.GetWindowTextLengthW(hWnd) + 1
            title = ctypes.create_unicode_buffer(length)
            user32.GetWindowTextW(hWnd, title, length)
            result.append([pid.value, title.value])
        return True
    user32.EnumWindows(WNDENUMPROC(enum_proc), 0)
    return sorted(result, key=lambda x:x[1])

def select_now(event):
    global SelectBox
    selected_index = SelectBox.curselection()
    selected_module = SelectBox.get(selected_index)
    return selected_module

## =========================================================================================
main()
