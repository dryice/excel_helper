import pyperclip


class ClipboardConverter:
    def __init__(self):
        self.clipboard_content = ""

    def read_clipboard(self):
        """Read the current clipboard content."""
        self.clipboard_content = pyperclip.paste()
        print("Clipboard content read successfully.")

    def process_content(self):
        """Process the clipboard content to format each cell as a separate paragraph."""
        # Split into lines and then into cells
        lines = self.clipboard_content.splitlines()
        cells = [cell for line in lines for cell in line.split('\t')]
        # Join cells into paragraphs
        formatted_content = '\n\n'.join(cells)
        return formatted_content

    def update_clipboard(self, content):
        """Update the clipboard with the formatted content."""
        pyperclip.copy(content)
        print("Clipboard updated with formatted content.")

    def convert_clipboard(self):
        """Read, process, and update the clipboard."""
        self.read_clipboard()
        formatted_content = self.process_content()
        self.update_clipboard(formatted_content)


if __name__ == "__main__":
    converter = ClipboardConverter()
    converter.convert_clipboard()
    print("Clipboard conversion completed. You can now paste into Word.")
