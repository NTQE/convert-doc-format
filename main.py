from src.convertworddoc.method import convert_documents
from src.convertworddoc.method import create_doc_list


def main():
    """
    Example script to find and convert .doc files into .docx
    """

    path_list = create_doc_list('C:\\TEST\\TEST2', file_type='.doc')

    print('\nFiles found: ')
    for path in path_list:
        print(f"\t{path}")
    print()

    new_path_list = convert_documents(path_list, 'docx')

    print('\nFiles Converted: ')
    for new_path in new_path_list:
        print(f"\t{new_path}")
    print()


if __name__ == '__main__':
    main()
