import pytest
import os

from modules.importWIP import listFiles, checkWipfile

def test_listFiles():
    path = './tests/test_listfiles'
    assert len(listFiles(path)) == 3
    assert listFiles(path) == ['tests\\test_listfiles\\test1.xls', 'tests\\test_listfiles\\test2.xlsx',
                               'tests\\test_listfiles\\WIP_file.xlsx']
    assert 'not_excel.txt' not in listFiles(path)


def test_checkWipfile():
    # test that only wip files are returned
    wipfile = ['tests\\test_listfiles\\WIP_file.xlsx', 'tests\\test_listfiles\\test1.xls']
    assert checkWipfile(wipfile) == ['tests\\test_listfiles\\WIP_file.xlsx']
    # test using a list that contains a non excel file
    wipfile = ['tests\\test_listfiles\\not_excel.txt']
    assert checkWipfile(wipfile) == []
