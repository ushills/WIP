# produce graph from database


# import libraries
import sqlite3
import matplotlib.pyplot as plt
from matplotlib.dates import MonthLocator, DateFormatter
from colorama import init, Fore
init()
import datetime
import os


# main routine
def print_graphs(database_name, output_directory):
    plot_graphs(database_name, output_directory)


# function to plot graphs - primary function
def plot_graphs(database_name, output_directory):

    months_to_plot = "12"

    # extract the most recent wip date from the database
    recent_wip_date = most_recent_wip(database_name)

    # extract the project_list based on the most recent wip date
    project_list = recent_project_list(recent_wip_date, database_name)

    # iterate through the project_list and plot graphs to each project
    sucessful = 0
    failed = 0
    for project in project_list:
        try:
            search_data = (project, months_to_plot)
            # plot the Forecast Data graph
            request = "wipdate, projectNumber, projectname.name, \
                       forecastCostTotal, forecastSaleTotal, \
                       forecastMarginTotal, currentCost, totalCertified"
            graph_data = import_data_sql(search_data, request, database_name)
            plot_forecast_graph(graph_data, output_directory)

            # plot the variation histogram
            request = "wipdate, projectNumber, projectname.name, \
                       agreedVariationsNo, budgetVariationsNo, \
                       submittedVariationsNo, agreedVariationsValue, \
                       budgetVariationsValue, submittedVariationsValue"
            graph_data = import_data_sql(search_data, request, database_name)
            plot_variation_graph(graph_data, output_directory)
            sucessful += 1

        except Exception as e:
            print(Fore.RED)
            print("skipping", search_data[0], "data incorrect")
            print("Exception error", e)
            print("-" * 25)
            print(Fore.RESET)
            failed += 1
            continue

    print(Fore.WHITE)
    print("printed", sucessful, "project graphs", Fore.RESET)
    if failed > 0:
        print(Fore.RED)
        print(failed, "projects failed to print", Fore.RESET)


# function to extract date of most recent wip in database
def most_recent_wip(database):

    # connect to the datebase
    try:
        conn = sqlite3.connect(os.path.normpath(database))
    except NameError:
        print(Fore.RED)
        print("Database file", database, "does not exist")
        print("Failed in most_recent_wip")
        print(Fore.RESET)
    else:
        cur = conn.cursor()
        # extract the most recent wipdate
        cur.execute(
            """
        SELECT max(wipDate) as latestdate FROM wipdata
        """
        )

        latest_date = cur.fetchone()
        # extract the latest_date string from the list returned
        latest_date = str(latest_date[0])

        try:
            assert latest_date != "None", (
                database + "contains an incorrect date format of None"
            )
        except AssertionError as e:
            print(Fore.RED)
            print("FATAL", e)
            print(Fore.RESET)
            pass

        return latest_date


# function to extract list of most recent wips
def recent_project_list(search_date, database):
    projects = []

    # connect to the database
    try:
        conn = sqlite3.connect(os.path.normpath(database))
    except NameError:
        print(Fore.RED)
        print("Database file", database, "does not exist")
        print("Failed in recent_project_list")
        print(Fore.RESET)
    else:
        cur = conn.cursor()

        # extract the list of most recent wips using the search_date
        cur.execute(
            """
        SELECT projectNumber AS wipsThisMonth
        FROM wipdata
        WHERE (JulianDay(:search_date) - JulianDay(wipDate)) < 35
        GROUP BY projectNumber""",
            {"search_date": search_date},
        )

        project_list = cur.fetchall()

        for i in range(len(project_list)):
            projects.append(project_list[i][0])

        return projects


# function to read data from SQL database based on job number & months
def import_data_sql(search_data, request, database):
    project_number = search_data[0]
    months = search_data[1]
    conn = sqlite3.connect(os.path.normpath(database))
    cur = conn.cursor()

    # extract seachdata
    # reverse sort order newest to oldest
    cur.execute(
        """
    SELECT * FROM (
    SELECT """
        + request
        + """
    FROM wipdata
    JOIN projectName
    ON wipdata.projectName = projectname.id
    WHERE projectNumber = :projectNumber
    AND forecastSaleTotal > 0
    ORDER BY wipdate DESC
    LIMIT :months)
    ORDER BY wipdate ASC
    """,
        {"projectNumber": project_number, "months": months},
    )
    return cur.fetchall()


# function to produce and output forecast graph
def plot_forecast_graph(graph_data, output_directory):

    # Set the common variables

    # These are the "Tableau 20" colors as RGB.
    tableau20 = [
        (31, 119, 180),
        (174, 199, 232),
        (255, 127, 14),
        (255, 187, 120),
        (44, 160, 44),
        (152, 223, 138),
        (214, 39, 40),
        (255, 152, 150),
        (148, 103, 189),
        (197, 176, 213),
        (140, 86, 75),
        (196, 156, 148),
        (227, 119, 194),
        (247, 182, 210),
        (127, 127, 127),
        (199, 199, 199),
        (188, 189, 34),
        (219, 219, 141),
        (23, 190, 207),
        (158, 218, 229),
    ]

    # Scale the RGB values to the [0, 1] range,
    # which is the format matplotlib accepts.
    for i in range(len(tableau20)):
        r, g, b = tableau20[i]
        tableau20[i] = (r / 255.0, g / 255.0, b / 255.0)

    # You typically want your plot to be ~1.33x wider than tall.
    # This plot is a rare exception because of the number of lines
    # being plotted on it.
    # Common sizes: (10, 7.5) and (12, 9)
    plt.figure(figsize=(12, 9))
    # fig = plt.figure()

    # change to classic matplotlib styles
    plt.style.use("classic")

    fig, (ax, ax3) = plt.subplots(nrows=2)
    ax2 = ax.twinx()

    # Create the forcast totals graph
    # create the lists
    dates = []
    project_number = []
    project_name = []
    forecast_cost = []
    forecast_sale = []
    forecast_contribution = []

    # loop through the list graph_data to extract the x and y axis
    # data is produced in the folllowing format
    # [(u'2016-01-31', -1355036), (u'2015-12-31', -1354858),
    for data in graph_data:
        dates.append(data[0])
        project_number.append(data[1])
        project_name.append(data[2])
        forecast_cost.append(data[3] / 1000)
        forecast_sale.append(data[4] / 1000)
        forecast_contribution.append(data[5] / 1000)

    # format the dates in the correct format to show
    dates = [datetime.datetime.strptime(d, "%Y-%m-%d").date() for d in dates]

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk
    ax.get_xaxis().tick_bottom()
    ax.get_yaxis().tick_left()
    ax2.get_xaxis().tick_bottom()
    ax2.get_yaxis().tick_right()

    # plot the data
    forecast_sale_line = ax.plot(
        dates, forecast_sale, lw=2.5, color=tableau20[5], label="Forecast Sale"
    )

    forecast_cost_line = ax.plot(
        dates, forecast_cost, lw=2.5, color=tableau20[0], label="Forecast Cost"
    )

    forecast_contribution_line = ax2.plot(
        dates, forecast_contribution, lw=2.5, color=tableau20[2], label="Contribution"
    )

    # set the y axis label
    ylabel = "\xA3k"
    ax.set_ylabel(ylabel, fontsize=14, rotation="vertical")
    ax2.set_ylabel(ylabel, fontsize=14, rotation="vertical", color=tableau20[2])

    for tl in ax2.get_yticklabels():
        tl.set_color(tableau20[2])

    # set the title of the graph
    title = project_name[0] + " (" + project_number[0] + ")\n"
    ax.set_title(title, fontsize=16, ha="center")
    plt.gcf().autofmt_xdate()

    # create the to-date graph
    # create the lists
    dates = []
    project_number = []
    project_name = []
    current_cost = []
    total_certified = []

    # loop through the list graph_data to extract the x and y axis
    # data is produced in the folllowing format
    # [(u'2016-01-31', -1355036), (u'2015-12-31', -1354858),
    for data in graph_data:
        dates.append(data[0])
        project_number.append(data[1])
        project_name.append(data[2])
        current_cost.append(data[6] / 1000)
        total_certified.append(data[7] / 1000)

    # format the dates in the correct format to show
    dates = [datetime.datetime.strptime(d, "%Y-%m-%d").date() for d in dates]

    # format the x axis date format
    months = MonthLocator()
    month_format = DateFormatter("%b-%Y")
    ax3.fmt_xdata = DateFormatter("%b-%Y")
    ax3.xaxis.set_major_locator(months)
    ax3.xaxis.set_major_formatter(month_format)
    ax.xaxis.set_major_locator(months)
    ax.xaxis.set_major_formatter(month_format)

    # create the axis
    # ax3 = plt.subplot(212)

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk.
    ax3.get_xaxis().tick_bottom()
    ax3.get_yaxis().tick_left()

    # plot the data
    current_cost_line = ax3.plot(
        dates, current_cost, lw=2.5, color=tableau20[11], label="Cost"
    )
    total_certified_line = ax3.plot(
        dates, total_certified, lw=2.5, color=tableau20[6], label="Certified"
    )

    # set the y axis label
    ylabel = "\xA3k"
    ax3.set_ylabel(ylabel, fontsize=14, rotation="vertical")

    # add the legend
    # shrink the axis height by 10% at the bottom
    box = ax3.get_position()
    ax3.set_position([box.x0, box.y0 + box.height * 0.1, box.width, box.height * 0.9])

    # combine ax1 and ax2 labels
    lines, labels = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    lines3, labels3 = ax3.get_legend_handles_labels()

    # place the legend below the axis
    ax3.legend(
        lines + lines2 + lines3,
        labels + labels2 + labels3,
        loc="upper center",
        fontsize=10,
        bbox_to_anchor=(0.5, -0.35),
        fancybox=True,
        shadow=True,
        ncol=5,
    )

    # plot the graph and save
    # plt.show()
    plt.savefig(
        os.path.normpath(
            output_directory + project_number[0] + " forecast totals graph.png"
        ),
        bbox_inches="tight",
    )
    plt.close("all")


# function to produce and output variation graph
def plot_variation_graph(graph_data, output_directory):

    # Set the common variables

    # These are the "Tableau 20" colors as RGB.
    tableau20 = [
        (31, 119, 180),
        (174, 199, 232),
        (255, 127, 14),
        (255, 187, 120),
        (44, 160, 44),
        (152, 223, 138),
        (214, 39, 40),
        (255, 152, 150),
        (148, 103, 189),
        (197, 176, 213),
        (140, 86, 75),
        (196, 156, 148),
        (227, 119, 194),
        (247, 182, 210),
        (127, 127, 127),
        (199, 199, 199),
        (188, 189, 34),
        (219, 219, 141),
        (23, 190, 207),
        (158, 218, 229),
    ]

    # Scale the RGB values to the [0, 1] range,
    # which is the format matplotlib accepts.
    for i in range(len(tableau20)):
        r, g, b = tableau20[i]
        tableau20[i] = (r / 255.0, g / 255.0, b / 255.0)

    # You typically want your plot to be ~1.33x wider than tall.
    # This plot is a rare exception because of the number of lines
    # being plotted on it.
    # Common sizes: (10, 7.5) and (12, 9)

    plt.figure(figsize=(12, 9))
    fig, (ax, ax2) = plt.subplots(nrows=2)

    # change to classic matplotlib styles
    plt.style.use("classic")

    # Create the forcast totals graph
    # create the lists
    dates = []
    project_number = []
    project_name = []
    agreed_variation_no = []
    budget_variation_no = []
    submitted_variation_no = []
    agreed_variation_value = []
    budget_variation_value = []
    submitted_variation_value = []

    # loop through the list graph_data to extract the x and y axis
    # data is produced in the folllowing format
    # [(u'2016-01-31', -1355036), (u'2015-12-31', -1354858),
    for data in graph_data:
        dates.append(data[0])
        project_number.append(data[1])
        project_name.append(data[2])
        agreed_variation_no.append(data[3])
        budget_variation_no.append(data[4])
        submitted_variation_no.append(data[5])
        agreed_variation_value.append(data[6] / 1000)
        budget_variation_value.append(data[7] / 1000)
        submitted_variation_value.append(data[8] / 1000)

    # create the bins for the bar chart and set width of bar
    bins = list(range(len(dates)))
    width_of_date = 0.9

    # format the dates in the correct format to show
    dates = [datetime.datetime.strptime(d, "%Y-%m-%d").date() for d in dates]

    # convert the dates to MMM-YYY
    xdate = []
    for d in dates:
        d = d.strftime("%b-%Y")
        xdate.append(d)
    dates = xdate

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk.
    ax.get_xaxis().tick_bottom()
    ax.get_yaxis().tick_left()

    # plot the data
    # as this is a stacked graph first we need to zip the arrays for the
    # bottom statement
    cumulative_number_histogram = [
        a + b for a, b in zip(agreed_variation_no, submitted_variation_no)
    ]

    # plot the bars
    agreed_variation_no_histogram = ax.bar(
        bins,
        agreed_variation_no,
        width=width_of_date,
        color=tableau20[5],
        label="Agreed",
    )

    submitted_variation_no_histogram = ax.bar(
        bins,
        submitted_variation_no,
        bottom=agreed_variation_no,
        width=width_of_date,
        color=tableau20[15],
        label="Submitted",
    )

    budget_variation_no_histogram = ax.bar(
        bins,
        budget_variation_no,
        bottom=cumulative_number_histogram,
        width=width_of_date,
        color=tableau20[11],
        label="Budget",
    )

    # set the y axis label
    ylabel = "No of Variations"
    ax.set_ylabel(ylabel, fontsize=14, rotation="vertical")

    # set the title of the graph
    title = project_name[0] + " (" + project_number[0] + ")\n"
    ax.set_title(title, fontsize=16, ha="center")

    # Ensure that the axis ticks only show up on the bottom and
    # left of the plot.
    # Ticks on the right and top of the plot are generally
    # unnecessary chartjunk.
    ax2.get_xaxis().tick_bottom()
    ax2.get_yaxis().tick_left()

    # plot the data
    # first set the width of the bins for side by side bar chart
    width_of_date = 0.3
    cumulative_bin = [x + width_of_date for x in bins]
    cumulative_bin_2 = [x + (2 * width_of_date) for x in bins]

    # plot the bars
    agreed_variation_value_histogram = ax2.bar(
        bins,
        agreed_variation_value,
        width=width_of_date,
        color=tableau20[5],
        label="Agreed",
    )

    submitted_variation_value_histogram = ax2.bar(
        cumulative_bin,
        submitted_variation_value,
        width=width_of_date,
        color=tableau20[15],
        label="Submitted",
    )

    budget_variation_value_histogram = ax2.bar(
        cumulative_bin_2,
        budget_variation_value,
        width=width_of_date,
        color=tableau20[11],
        label="Budget",
    )

    # set the y axis label
    ylabel = "Value of Variations\n(\xA3k)"
    ax2.set_ylabel(ylabel, fontsize=14, rotation="vertical")

    # set the x axis label
    ax2.set_xticks(cumulative_bin)
    ax2.set_xticklabels(dates)
    plt.gcf().autofmt_xdate()

    # add the legend
    # shrink the axis height by 10% at the bottom
    box = ax2.get_position()
    ax2.set_position([box.x0, box.y0 + box.height * 0.1, box.width, box.height * 0.9])

    # combine ax1 and ax2 labels
    lines, labels = ax.get_legend_handles_labels()

    # place the legend below the axis
    ax2.legend(
        lines,
        labels,
        loc="upper center",
        fontsize=10,
        bbox_to_anchor=(0.5, -0.35),
        fancybox=True,
        shadow=True,
        ncol=5,
    )

    # plot the graph and save
    # plt.show()
    plt.savefig(
        os.path.normpath(
            output_directory + project_number[0] + " variations graph.png"
        ),
        bbox_inches="tight",
    )
    plt.close("all")


if __name__ == "__main__":
    plot_graphs("./../london/londonwipdata.sqlite", "./../london/graphs/")
