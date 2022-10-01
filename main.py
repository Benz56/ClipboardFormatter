import time
import win32clipboard
import win32gui
from PIL import ImageGrab
from io import BytesIO

from clipboard_data import ClipboardData, ImageClipboardData, TextClipboardData

prev = ClipboardData('')
sources = ['Adobe Acrobat Reader DC (32-bit)', 'Screen Snipping']


def get_current_window():
    return win32gui.GetWindowText(win32gui.GetForegroundWindow())


def is_copied_from_selected_source():
    return get_current_window() in sources or 'ANY' in sources


def get_clipboard_data():
    try:
        win32clipboard.OpenClipboard()
        try:
            data = TextClipboardData(win32clipboard.GetClipboardData())  # Try to get text
            win32clipboard.CloseClipboard()
        except TypeError:
            win32clipboard.CloseClipboard()
            data = get_clipboard_image()  # Try to get image
    except BaseException as e:
        data = ClipboardData('')
    return data


def get_clipboard_image():
    im = ImageGrab.grabclipboard()
    if im is None:
        return ClipboardData('')
    max_width = 870
    if im.width > max_width:
        ratio = im.width / im.height
        height = int(max_width / ratio)
        im = im.resize((max_width, height))
    with BytesIO() as output:
        im.convert("RGB").save(output, "BMP")
        data = ImageClipboardData(output.getvalue()[14:])
    return data


def process_clipboard():
    global prev
    clipboard_data = get_clipboard_data()
    if clipboard_data.content == '' or prev.content == clipboard_data.content:
        return

    prev = clipboard_data
    if not is_copied_from_selected_source():
        return
    clipboard_data.format_content()
    if clipboard_data.replace_clipboard():
        prev = clipboard_data


if __name__ == '__main__':
    while True:
        process_clipboard()
        time.sleep(0.1)
