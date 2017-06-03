# -*- coding: utf-8 -*-
import fnmatch
import os


class QuestionType(enumerate):
    choice = 'choice'
    s_choice = 'single_choice'
    m_choice = 'multi_choice'
    blank = 'blank'
    yes_no = 'yes_no'


def find_all_files(path):
    files = []
    for root, dirnames, filenames in os.walk(path):
        for filename in fnmatch.filter(filenames, '*.docx'):
            files.append(os.path.join(root, filename))
    return files


def get_file_name(path):
    last_slash = path.rfind('/')
    last_dot = path.rfind('.')
    return path[last_slash + 1:last_dot]
