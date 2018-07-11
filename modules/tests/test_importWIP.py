import pytest
import os
import sqlite3

from modules.importWIP import listFiles, checkWipfile, importData, importWIPdata
from modules.multigraph import printgraphs


def test_listFiles():
    path = './test_listfiles'
    filelist = listFiles(path)
    print(filelist)
    assert len(filelist) == 4
    assert 'test_listfiles\\not_excel.txt' not in filelist
    assert 'test_listfiles\\test1.xls' in filelist
    assert 'test_listfiles\\test2.xlsx' in filelist
    assert 'test_listfiles\\WIP_file.xlsx' in filelist
    assert 'test_listfiles\\not_WIP_file.xlsx' in filelist
    # check temp files beginning with ~ are not included
    assert 'test_listfiles\\~WIP_file.xlsx' not in filelist


def test_checkWipfile():
    # test that only wip files are returned
    wipfile = ['test_listfiles\\WIP_file.xlsx', 'test_listfiles\\test1.xls']
    assert checkWipfile(wipfile) == ['test_listfiles\\WIP_file.xlsx']
    # test using a list that contains a non excel file
    wipfile = ['test_listfiles\\not_excel.txt']
    assert checkWipfile(wipfile) == []
    # test that a non wip file is rejected
    wipfile = ['test_listfiles\\not_WIP_file.xlsx']
    assert checkWipfile(wipfile) == []


def test_importData():
    wipfile = 'test_listfiles\\WIP_file.xlsx'
    data = importData(wipfile)
    assert data == {'projectName': 'test_project', 'projectNumber': 'test_project_number',
                       'wipDate': '1970-01-01', 'agreedVariationsNo': 1, 'budgetVariationsNo': 1,
                       'submittedVariationsNo': 1, 'variationsNoTotal': 3, 'orderValue': 6293357,
                       'agreedVariationsValue': 25450, 'budgetVariationsValue': 114,
                       'submittedVariationsValue': 56787, 'saleSubtotal': 6375708,
                       'reserveAgreed': -254, 'reserveBudget': -57, 'reserveSubmitted': -5678,
                       'contracharges': -9889, 'forecastSaleTotal': 6359828, 'currentCost': 54874,
                       'costToComplete': 87899, 'defectProvision': 68998, 'forecastCostTotal': 211771,
                       'contractContribution': 6148057, 'bettermentsRisks': 5587, 'managersView': 147,
                       'forecastMarginTotal': 6153791, 'latestApplication': 6598787, 'totalCertified': 1144}
    assert type(data['projectName']) is str
    assert type(data['projectNumber']) is str
    assert type(data['wipDate']) is str


class TestSqldata:

    def test_createdatabase(self):
        # import the data & create the database
        db_name = './testdata.sqlite'
        importWIPdata(db_name, './test_listfiles')
        assert os.path.isfile(db_name) is True


    def test_sqldata(self):
        db_name = './testdata.sqlite'
        conn = sqlite3.connect(os.path.normpath(db_name))
        cur = conn.cursor()
        cur.execute('''
            SELECT * FROM wipdata
            JOIN projectname
            ON wipdata.projectName = projectname.id
            ''')
        all_rows = cur.fetchall()
        print(all_rows)
        assert all_rows == [('test_project_number', 1, '1970-01-01', 1, 1, 1, 3, 6293357, 25450, 114,
                            56787, 6375708, -254, -57, -5678, -9889, 6359828, 54874, 87899, 68998,
                            211771, 6148057, 5587, 147, 6153791, 6598787, 1144, 1, 'test_project')]


    def test_deletedatabase(self):
        # delete the file when finished
        db_name = './testdata.sqlite'
        if os.path.isfile(db_name):
            os.remove('./testdata.sqlite')
        else:
            print(db_name, 'does not exist')
        assert os.path.isfile(db_name) is False
