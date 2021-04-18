from PIL import Image, ImageOps
from pathlib import Path
import win32com.client as win32
import pyautogui
import win32gui
import time
import numpy as np
import os

FOLDER_WITH_FILES = "Test"
doc_path = f"C:\\Users\\91635\\Desktop\\{FOLDER_WITH_FILES}"
directory = os.fsencode(doc_path)

# Create an images folder
image_path = f"{doc_path}\\Images"
Path(image_path).mkdir(parents=True, exist_ok=True)


def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


word = win32.gencache.EnsureDispatch('Word.Application')


def open_doc(file):
    global word
    docfile = f'C:\\Users\\91635\\Desktop\\Test\\{file}'
    word.Visible = True
    word.WindowState = 1  # maximize

    top_windows = []
    win32gui.EnumWindows(windowEnumerationHandler, top_windows)

    for i in top_windows:  # all open apps
        if "word" in i[1].lower():  # find word (assume only one)
            try:
                win32gui.ShowWindow(i[0], 5)
                win32gui.SetForegroundWindow(i[0])  # bring to front
                break
            except:
                pass

    doc = word.Documents.Add(docfile)  # open file

    time.sleep(1.1)  # wait for doc to load

    image_name = file.partition(".")
    image_save_path = f"{image_path}\\{image_name[0]}.png"

    my_screenshot = pyautogui.screenshot(region=(500, 230, 940, 700))  # take screenshot
    my_screenshot.save(image_save_path)  # save screenshot
    process_screenshot(image_save_path)
    print(image_save_path)
    doc.Close()
    # word.Application.Quit()


def take_screenshot(doc, word, file):
    image_save_path = f"{image_path}\\{file}"
    my_screenshot = pyautogui.screenshot(region=(500, 230, 940, 700))  # take screenshot
    my_screenshot.save(image_save_path)  # save screenshot

    # close doc and word app
    doc.Close()
    word.Application.Quit()


def process_screenshot(image_file):
    padding = 10
    padding = np.asarray([-1 * padding, -1 * padding, padding, padding]) # left, top, right, bottom
    image = Image.open(image_file)
    image.load()
    image_size = image.size

    # remove alpha channel
    invert_im = image.convert("RGB")

    # invert image (so that white is 0)
    invert_im = ImageOps.invert(invert_im)
    image_box = invert_im.getbbox()
    image_box = tuple(np.asarray(image_box) + padding)

    cropped = image.crop(image_box)
    cropped.save(image_file)


for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith(".docx"):
        open_doc(filename)

word.Application.Quit()
