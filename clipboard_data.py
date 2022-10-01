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
    def __init__(self, content):
        super().__init__(content)

    def format_content(self):
        pass  # Preformatted when retrieved from clipboard.

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
