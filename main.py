import time
import win32clipboard
import win32gui
from PIL import ImageGrab

from clipboard_data import ClipboardData, ImageClipboardData, TextClipboardData

# Config options
sources = ['Screen Snipping']
source_endings = ['Adobe Acrobat Reader DC (32-bit)']

# Used for checking if the clipboard has been updated
sequence_number = -1


def get_current_window():
    return win32gui.GetWindowText(win32gui.GetForegroundWindow())


def is_copied_from_selected_source():
    current_window = get_current_window()
    return current_window in sources or any(current_window.endswith(source) for source in source_endings) or 'ANY' in sources


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
    return ImageClipboardData(im)


def process_clipboard():
    global sequence_number

    # Check if the clipboard has been updated
    new_sequence_number = win32clipboard.GetClipboardSequenceNumber()
    if sequence_number == new_sequence_number:
        if sequence_number != -1:
            return

    sequence_number = new_sequence_number

    if not is_copied_from_selected_source():
        return

    print('Processing clipboard')

    clipboard_data = get_clipboard_data()
    if clipboard_data.content == '' or clipboard_data.content is None:
        return

    clipboard_data.format_content()
    if clipboard_data.replace_clipboard():
        print('Clipboard formatted and replaced')
        sequence_number = win32clipboard.GetClipboardSequenceNumber()  # Update sequence number as it has changed during processing
    else:
        print('Failed to replace clipboard')


if __name__ == '__main__':
    while True:
        process_clipboard()
        time.sleep(0.1)

# TODO OCR to estimate text font in image and scale accordingly.
