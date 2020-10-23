#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "marcus"

import sys
import argparse
import os
import zipfile


def create_parser():
    """Creates an argument parser object."""
    """ to watch(dir) and magic text(magic)"""
    parser = argparse.ArgumentParser()
    parser.add_argument('dir',
                        help='destination directory to look for files')
    parser.add_argument('magic',
                        help='a search phrase')
    return parser


def print_dotm(dir):
    files = os.listdir(dir)
    matches = 0
    total_files = 0
    for file in files:
        if ".dotm" == os.path.splitext(file)[1]:
            total_files += 1
            with zipfile.ZipFile(dir+"/"+file) as myzip:
                # print(myzip)
                with myzip.open("word/document.xml") as myfile:
                    doc = myfile.read()
                    docm = str(doc)
                    if "$" in docm:
                        print(f"{dir}/{file}")
                        cash_index = docm.find("$")
                        print(docm[(cash_index-40):(cash_index+40)])
                        matches += 1
    print(f'exit log: matches - {matches} total_files - {total_files}')


def main(args):
    parser = create_parser()
    if not args:
        parser.print_usage()
        sys.exit(1)
    parsed_args = parser.parse_args(args)
    print_dotm(parsed_args.dir)


if __name__ == '__main__':
    main(sys.argv[1:])
