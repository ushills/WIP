# import pytest
import os
import sqlite3
from pathlib import WindowsPath

from modules.importWIP import (
    list_files,
    check_wipfile,
    import_data,
    import_wip_data,
    export_data_sql,
)
from modules.multigraph import (
    most_recent_wip,
    recent_project_list,
    import_data_sql,
    print_graphs,
)

test_path = "modules/tests/"


class TestImportWIP:
    def test_list_files(self):
        filelist = list_files(test_path + "test_listfiles/")
        print(filelist)
        assert len(filelist) == 4
        assert WindowsPath(test_path + "test_listfiles/not_excel.txt") not in filelist
        assert WindowsPath(test_path + "test_listfiles/test1.xls") in filelist
        assert WindowsPath(test_path + "test_listfiles/test2.xlsx") in filelist
        assert WindowsPath(test_path + "test_listfiles/WIP_file.xlsx") in filelist
        assert WindowsPath(test_path + "test_listfiles/not_WIP_file.xlsx") in filelist
        # check temp files beginning with ~ are not included
        assert WindowsPath(test_path + "test_listfiles/~WIP_file.xlsx") not in filelist

    def test_check_wipfile(self):
        # test that only wip files are returned
        wipfile = [
            test_path + "test_listfiles\\WIP_file.xlsx",
            test_path + "test_listfiles\\test1.xls",
        ]
        assert check_wipfile(wipfile) == [test_path + "test_listfiles\\WIP_file.xlsx"]
        # test using a list that contains a non excel file
        wipfile = [test_path + "test_listfiles\\not_excel.txt"]
        assert check_wipfile(wipfile) == []
        # test that a non wip file is rejected
        wipfile = [test_path + "test_listfiles\\not_WIP_file.xlsx"]
        assert check_wipfile(wipfile) == []

    def test_import_data(self):
        wipfile = test_path + "test_listfiles\\WIP_file.xlsx"
        data = import_data(wipfile)
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

    def test_export_data_sql(self):
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
        export_data_sql(data, db_name)
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
    def test_create_database(self):
        # import the data & create the database
        db_name = test_path + "\\test_database\\testdata.sqlite"
        import_wip_data(db_name, test_path + "\\test_listfiles")
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

    def test_most_recent_wip(self):
        db_name = test_path + "test_database\\testdata.sqlite"
        assert most_recent_wip(db_name) == "1970-01-01"

    def test_recent_project_list(self):
        db_name = test_path + "test_database\\testdata.sqlite"
        assert recent_project_list("1970-02-01", db_name) == ["test_project_number"]

    def test_import_data_sql(self):
        db_name = test_path + "test_database\\testdata.sqlite"
        months = 12
        search_data = ("test_project_number", months)
        request = "wipdate, projectNumber, projectname.name, \
                   forecastCostTotal, forecastSaleTotal, \
                   forecastMarginTotal, currentCost, totalCertified, \
                   agreedVariationsNo, budgetVariationsNo, \
                   submittedVariationsNo, agreedVariationsValue, \
                   budgetVariationsValue, submittedVariationsValue"
        assert import_data_sql(search_data, request, db_name) == [
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

    def test_delete_database(self):
        # delete the file when finished
        db_name = test_path + "test_database\\testdata.sqlite"
        if os.path.isfile(db_name):
            os.remove(test_path + "test_database\\testdata.sqlite")
        else:
            print(db_name, "does not exist")
        assert os.path.isfile(db_name) is False

    def test_print_graphs(self):
        print_graphs(
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
