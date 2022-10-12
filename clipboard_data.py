import time
import traceback
from io import BytesIO

import win32clipboard


class ClipboardData:
    def __init__(self, content):
        self.content = content

    def replace_clipboard(self):
        for tries in range(3):
            try:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                self.set_clipboard_content()  # Delegate to subclass
                win32clipboard.CloseClipboard()
                return True
            except BaseException as e:
                traceback.print_exc()
                print("Failed to replace clipboard. Retrying...")
                time.sleep(0.1)
        print("Failed to replace clipboard!")
        return False

    def format_content(self):
        raise NotImplementedError("Please Implement this method")

    def set_clipboard_content(self):
        raise NotImplementedError("Please Implement this method")

    def __eq__(self, other):
        return self.content == other.content


class ImageClipboardData(ClipboardData):
    image_max_width = 870

    def __init__(self, content):
        super().__init__(content)

    def format_content(self):
        if self.content.width > self.image_max_width:
            ratio = self.content.width / self.content.height
            height = int(self.image_max_width / ratio)
            self.content = self.content.resize((self.image_max_width, height))
        with BytesIO() as output:
            self.content.convert("RGB").save(output, "BMP")
            self.content = output.getvalue()[14:]

    def set_clipboard_content(self):
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, self.content)


class TextClipboardData(ClipboardData):
    def __init__(self, content):
        super().__init__(content)

    def format_content(self):
        def remove_newlines(text):
            return text.replace('\r', '').replace('\n', ' ')

        self.content = remove_newlines(self.content)

    def set_clipboard_content(self):
        win32clipboard.SetClipboardText(self.content)
