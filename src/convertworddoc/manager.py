import win32com.client as win32


class WordManager:
    def __init__(self):
        print('Initializing Word')
        self.word = win32.gencache.EnsureDispatch('Word.Application')

    def __enter__(self):
        print('Returning Word')
        return self.word

    def __exit__(self, exc_type, exc_val, exc_tb):
        print('Closing Word')
        self.word.Quit()


class DocManager:
    def __init__(self, word, path: str):
        print('Opening Document')
        self.doc = word.Documents.Open(path)

    def __enter__(self):
        print('Returning Document')
        return self.doc

    def __exit__(self, exc_type, exc_val, exc_tb):
        print('Closing Document')
        self.doc.Close()
