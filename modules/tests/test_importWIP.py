# import pytest
import os
import sqlite3

global test_path
test_path = "modules\\tests\\"

from modules.importWIP import (
    listFiles,
    checkWipfile,
    importData,
    importWIPdata,
    exportDataSql,
)
from modules.multigraph import (
    mostRecentWip,
    recentProjectList,
    importDataSql,
    printgraphs,
)


class TestImportWIP:
    def test_listFiles(self):
        filelist = listFiles(test_path + "test_listfiles\\")
        print(filelist)
        assert len(filelist) == 4
        assert test_path + "test_listfiles\\not_excel.txt" not in filelist
        assert test_path + "test_listfiles\\test1.xls" in filelist
        assert test_path + "test_listfiles\\test2.xlsx" in filelist
        assert test_path + "test_listfiles\\WIP_file.xlsx" in filelist
        assert test_path + "test_listfiles\\not_WIP_file.xlsx" in filelist
        # check temp files beginning with ~ are not included
        assert test_path + "test_listfiles\\~WIP_file.xlsx" not in filelist

    def test_checkWipfile(self):
        # test that only wip files are returned
        wipfile = [
            test_path + "test_listfiles\\WIP_file.xlsx",
            test_path + "test_listfiles\\test1.xls",
        ]
        assert checkWipfile(wipfile) == [test_path + "test_listfiles\\WIP_file.xlsx"]
        # test using a list that contains a non excel file
        wipfile = [test_path + "test_listfiles\\not_excel.txt"]
        assert checkWipfile(wipfile) == []
        # test that a non wip file is rejected
        wipfile = [test_path + "test_listfiles\\not_WIP_file.xlsx"]
        assert checkWipfile(wipfile) == []

    def test_importData(self):
        wipfile = test_path + "test_listfiles\\WIP_file.xlsx"
        data = importData(wipfile)
        assert data == {
            "projectName": "test_project",
            "projectNumber": "test_project_number",
            "wipDate": "1970-01-01",
            "agreedVariationsNo": 1,
            "budgetVariationsNo": 1,
            "submittedVariationsNo": 1,
            "variationsNoTotal": 3,
            "orderValue": 6293357,
            "agreedVariationsValue": 25450,
            "budgetVariationsValue": 114,
            "submittedVariationsValue": 56787,
            "saleSubtotal": 6375708,
            "reserveAgreed": -254,
            "reserveBudget": -57,
            "reserveSubmitted": -5678,
            "contracharges": -9889,
            "forecastSaleTotal": 6359828,
            "currentCost": 54874,
            "costToComplete": 87899,
            "defectProvision": 68998,
            "forecastCostTotal": 211771,
            "contractContribution": 6148057,
            "bettermentsRisks": 5587,
            "managersView": 147,
            "forecastMarginTotal": 6153791,
            "latestApplication": 6598787,
            "totalCertified": 1144,
        }
        assert type(data["projectName"]) is str
        assert type(data["projectNumber"]) is str
        assert type(data["wipDate"]) is str

    def test_exportDataSql(self):
        db_name = test_path + "test_database\\testdata.sqlite"
        data = {
            "projectName": "test_project",
            "projectNumber": "test_project_number",
            "wipDate": "1970-01-01",
            "agreedVariationsNo": 1,
            "budgetVariationsNo": 1,
            "submittedVariationsNo": 1,
            "variationsNoTotal": 3,
            "orderValue": 6293357,
            "agreedVariationsValue": 25450,
            "budgetVariationsValue": 114,
            "submittedVariationsValue": 56787,
            "saleSubtotal": 6375708,
            "reserveAgreed": -254,
            "reserveBudget": -57,
            "reserveSubmitted": -5678,
            "contracharges": -9889,
            "forecastSaleTotal": 6359828,
            "currentCost": 54874,
            "costToComplete": 87899,
            "defectProvision": 68998,
            "forecastCostTotal": 211771,
            "contractContribution": 6148057,
            "bettermentsRisks": 5587,
            "managersView": 147,
            "forecastMarginTotal": 6153791,
            "latestApplication": 6598787,
            "totalCertified": 1144,
        }
        # create database
        exportDataSql(data, db_name)
        # check database is created
        assert os.path.isfile(db_name) is True
        conn = sqlite3.connect(os.path.normpath(db_name))
        cur = conn.cursor()
        cur.execute(
            """
            SELECT * FROM wipdata
            JOIN projectname
            ON wipdata.projectName = projectname.id
            """
        )
        all_rows = cur.fetchall()
        # check data returned is correct
        assert all_rows == [
            (
                "test_project_number",
                1,
                "1970-01-01",
                1,
                1,
                1,
                3,
                6293357,
                25450,
                114,
                56787,
                6375708,
                -254,
                -57,
                -5678,
                -9889,
                6359828,
                54874,
                87899,
                68998,
                211771,
                6148057,
                5587,
                147,
                6153791,
                6598787,
                1144,
                1,
                "test_project",
            )
        ]


class TestMultiGraph:
    def test_createdatabase(self):
        # import the data & create the database
        db_name = test_path + "\\test_database\\testdata.sqlite"
        importWIPdata(db_name, test_path + "\\test_listfiles")
        assert os.path.isfile(db_name) is True

    def test_sqldata(self):
        # check the data is correct
        db_name = test_path + "test_database\\testdata.sqlite"
        conn = sqlite3.connect(os.path.normpath(db_name))
        cur = conn.cursor()
        cur.execute(
            """
            SELECT * FROM wipdata
            JOIN projectname
            ON wipdata.projectName = projectname.id
            """
        )
        all_rows = cur.fetchall()
        assert all_rows == [
            (
                "test_project_number",
                1,
                "1970-01-01",
                1,
                1,
                1,
                3,
                6293357,
                25450,
                114,
                56787,
                6375708,
                -254,
                -57,
                -5678,
                -9889,
                6359828,
                54874,
                87899,
                68998,
                211771,
                6148057,
                5587,
                147,
                6153791,
                6598787,
                1144,
                1,
                "test_project",
            )
        ]

    def test_mostRecentWip(self):
        db_name = test_path + "test_database\\testdata.sqlite"
        assert mostRecentWip(db_name) == "1970-01-01"

    def test_recentProjectList(self):
        db_name = test_path + "test_database\\testdata.sqlite"
        assert recentProjectList("1970-02-01", db_name) == ["test_project_number"]

    def test_importDataSql(self):
        db_name = test_path + "test_database\\testdata.sqlite"
        months = 12
        searchData = ("test_project_number", months)
        request = "wipdate, projectNumber, projectname.name, \
                   forecastCostTotal, forecastSaleTotal, \
                   forecastMarginTotal, currentCost, totalCertified, \
                   agreedVariationsNo, budgetVariationsNo, \
                   submittedVariationsNo, agreedVariationsValue, \
                   budgetVariationsValue, submittedVariationsValue"
        assert importDataSql(searchData, request, db_name) == [
            (
                "1970-01-01",
                "test_project_number",
                "test_project",
                211771,
                6359828,
                6153791,
                54874,
                1144,
                1,
                1,
                1,
                25450,
                114,
                56787,
            )
        ]

    def test_deletedatabase(self):
        # delete the file when finished
        db_name = test_path + "test_database\\testdata.sqlite"
        if os.path.isfile(db_name):
            os.remove(test_path + "test_database\\testdata.sqlite")
        else:
            print(db_name, "does not exist")
        assert os.path.isfile(db_name) is False

    def test_printgraphs(self):
        printgraphs(
            test_path + "test_database\\test.sqlite", test_path + "\\test_graphs\\test_"
        )
        assert (
            os.path.isfile(test_path + "test_graphs\\test_12345 variations graph.png")
            is True
        )
        assert (
            os.path.isfile(
                test_path + "test_graphs\\test_12345 forecast totals graph.png"
            )
            is True
        )
