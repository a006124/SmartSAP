import argparse
import os

from prompt_toolkit import prompt
from prompt_toolkit.completion import PathCompleter

from auto_doc_checker.excel_reader import query_and_fill_excel

os.environ["HTTP_PROXY"] = "http://cosmos2.mc2.renault.fr:3128"
os.environ["HTTPS_PROXY"] = "http://cosmos2.mc2.renault.fr:3128"


def parse_cli_args():
    parser = argparse.ArgumentParser(
        prog="auto_doc_checker",
        description="SAP Documentation checking automation",
    )
    parser.add_argument("file_path", nargs="?", help="Specify an Excel file to read and fill")

    return parser.parse_args()


def main():
    args = parse_cli_args()

    file_path = args.file_path
    if not file_path:
        file_path = prompt("Path to Excel file : ", completer=PathCompleter()).strip()

    query_and_fill_excel(file_path)


if __name__ == "__main__":
    main()
