from win32com import __gen_path__
import win32com.client as win32
from os import walk
from shutil import rmtree
import ctypes
import py_win_keyboard_layout

def get_doc():
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.ActiveDocument
        return word
    except AttributeError:
        paths = list(walk(__gen_path__))
        rmtree(paths[0][0] + '\\' + paths[0][1][0])
        return get_doc()


def add_formul(doc):
    selection = doc.Selection
    cursor_position = selection.Start
    range_until_cursor = doc.ActiveDocument.Range(max(cursor_position-1, 0), cursor_position)
    if range_until_cursor.Text != '\r' and range_until_cursor.OMaths.Count > 0:
        selection.EndKey(Unit=5) 
        selection.TypeParagraph() 
    selection.Range.OMaths.Add(selection.Range)


def switch_keyboard_layout():
    u = ctypes.windll.LoadLibrary("user32.dll")
    pf = getattr(u, "GetKeyboardLayout")
    if hex(pf(0)) == '0x4190419':
        py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)

def main(): 
    doc = get_doc()
    add_formul(doc)
    switch_keyboard_layout()


if __name__ == '__main__':
    main()

