import time
import win32clipboard
import win32gui
from PIL import ImageGrab
from pyxtension.streams import stream

from clipboard_data import ClipboardData, ImageClipboardData, TextClipboardData

# Config options
full_sources = ['AVPageView']
in_sources = ['Chrome']
ending_sources = ['Adobe Acrobat Reader DC (32-bit)']

# Used for checking if the clipboard has been updated
sequence_number = -1


def get_current_windows():
    """Returns both the active window and the window under the cursor"""
    return win32gui.GetWindowText(win32gui.GetForegroundWindow()), win32gui.GetWindowText(win32gui.WindowFromPoint(win32gui.GetCursorPos()))


def is_copied_from_selected_source():
    current_windows = get_current_windows()
    full_source_match = stream(current_windows).exists(lambda current_window: current_window in full_sources)
    in_source_match = stream(current_windows).exists(lambda current_window: stream(in_sources).exists(lambda source: source in current_window))
    ending_source_match = stream(current_windows).exists(lambda current_window: stream(ending_sources).exists(lambda source: current_window.endswith(source)))
    return full_source_match or in_source_match or ending_source_match or 'ANY' in full_sources


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


def wait_for_screen_snipping_to_finish():
    if 'Screen Snipping' in get_current_windows():
        time.sleep(0.01)
        wait_for_screen_snipping_to_finish()


def process_clipboard():
    global sequence_number

    # Check if the clipboard has been updated
    new_sequence_number = win32clipboard.GetClipboardSequenceNumber()
    if sequence_number == new_sequence_number:
        if sequence_number != -1:
            return

    sequence_number = new_sequence_number

    wait_for_screen_snipping_to_finish()

    if not is_copied_from_selected_source():
        print('Not copied from selected source. Source is {}'.format(get_current_windows()))
        return

    print('Processing clipboard from source: {}'.format(' || '.join(get_current_windows())))

    clipboard_data = None
    for tries in range(3):
        clipboard_data = get_clipboard_data()
        if clipboard_data.content != '' and clipboard_data.content is not None:
            break
        print("Clipboard is empty. Retrying...")
        time.sleep(0.1)
    else:
        print("Clipboard is empty!")
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
