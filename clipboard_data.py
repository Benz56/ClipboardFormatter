from io import BytesIO

import win32clipboard


class ClipboardData:
    def __init__(self, content):
        self.content = content

    def format_content(self):
        raise NotImplementedError("Please Implement this method")

    def replace_clipboard(self):
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

    def replace_clipboard(self):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, self.content)
            win32clipboard.CloseClipboard()
            return True
        except BaseException as e:
            return False


class TextClipboardData(ClipboardData):
    def __init__(self, content):
        super().__init__(content)

    def format_content(self):
        def remove_newlines(text):
            return text.replace("\r", "").replace("\n", "")
        self.content = remove_newlines(self.content)

    def replace_clipboard(self):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(self.content, win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return True
        except BaseException as e:
            return False
