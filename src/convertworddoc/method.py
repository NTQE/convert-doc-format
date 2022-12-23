from src.convertworddoc.manager import WordManager
from src.convertworddoc.manager import DocManager
import os

text_set = ('txt', 'text', 'tx', 't', '7')
docx_set = ('docx', 'default', 'dox', 'd', '16')

file_format_dict = {
    7: text_set,
    16: docx_set
}

file_extension_dict = {
    7: '.txt',
    16: '.docx'
}


def create_doc_list(folder: str, file_type: str = '.doc', recurse: bool = False) -> list[str]:
    """Create a list of documents with extension of 'file_type' inside a directory. Can be a recursive or flat search.

    :param folder: Directory where the files to be found exist
    :param file_type: String annotating the file extension
    :param recurse: Recursively search the directory or not.
    :return: A list of absolute file paths
    """

    if recurse:
        pass
    else:
        path_list = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(file_type)]
        return path_list


def convert_document(path: str, convert_to: str) -> str:
    """Convert a single document into a new file format using Microsoft Word. This function uses pywin32 with the aid
    of context managers for the Word Application and Document.

    Note: Current supported conversion formats include:
    txt
    docx

    Developer Note: When using the word.ActiveDocument.SaveAs(path=, FileFormat=) method, the path must include
    an extension and FileFormat takes an integer as far as I can tell.

    :param path:
    :param convert_to:
    :return:
    """

    convert_to = convert_to.lower()
    print(f'Convert to: {convert_to}')
    for k, v in file_format_dict.items():
        print(f'Checking value {v} for {convert_to}')
        if convert_to in v:
            file_format = k

    new_ext = file_extension_dict.get(file_format)

    with WordManager() as word:
        print('Using Word')
        if not os.path.isfile(path):
            return f"This path doesn\'t lead to a real file. \n{path}\n"
        with DocManager() as doc:
            print('Using Document')
            new_path = f"{os.path.splitext(path)[0]}{new_ext}"
            print(f'Converting {path} to {new_ext}')
            word.ActiveDocument.SaveAs(new_path, FileFormat=file_format)

    return new_path


def convert_documents(path_list: list[str], convert_to: str) -> list[str]:
    """Convert a group of documents into a new file format using Microsoft Word. This function uses pywin32 with the
    aid of context managers for the Word Application and Document.

    Note: Current supported conversion formats include:
    txt
    docx

    Developer Note: When using the word.ActiveDocument.SaveAs(path=, FileFormat=) method, the path must include
    an extension and FileFormat takes an integer as far as I can tell.

    :param path_list: list of absolute paths to the documents to be converted
    :param convert_to: string representing what file type to convert to
    :return: absolute paths of new documents in the new file format
    """

    convert_to = convert_to.lower()
    print(f'Convert to: {convert_to}')
    for k, v in file_format_dict.items():
        print(f'Checking value {v} for {convert_to}')
        if convert_to in v:
            file_format = k

    new_ext = file_extension_dict.get(file_format)
    new_path_list = []

    with WordManager() as word:
        print('Using Word')
        for path in path_list:
            if not os.path.isfile(path):
                print(f'One of the paths doesn\'t lead to a real file. Skipping File.\n{path}\n')
                continue
            with DocManager(word=word, path=path) as doc:
                print('Using Document')
                new_path = f"{os.path.splitext(path)[0]}{new_ext}"
                print(f'Converting {path} to {new_ext}')
                word.ActiveDocument.SaveAs(new_path, FileFormat=file_format)
                new_path_list.append(new_path)

    return new_path_list
