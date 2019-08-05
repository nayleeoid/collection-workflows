"""
Author Linnea Shieh <laiello@stanford.edu>
Copyright Stanford University 2019

Figures out remaining recurring purchases for a fiscal year, and outputs
them tabularly along with their cost estimates from the previous year.
"""

import argparse
import logging
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


RECURRING_TYPES = set(["4EXCH2", "APPROVAL2", "BLANKET", "EXCHANGE", "EXCHSER",
                       "GIFTSER", "REPLACE2", "SCORES2", "STANDING", "SUBSCRIPT",
                       "TERMSET", "UNITARY2", "CATREC2", "DATAB2", "DATAB2PUR",
                       "EBOOK2", "EBOOK2PUR", "ELEC2", "E-PACKAGE", "MAINTFEE",
                       "COMBO", "COMBO-CHG", "DATASET2", "MEMBERSHIP", "PACKAGE",
                       "SO-COMBO", "AUDIO2", "MAP2", "MICRO", "VISUAL2", "DEPNDT",
                       "DEPOSIT", "FISCAL2", "SHIPPING"])
OUTPUT_COLUMNS = set(["ORDER ID", "ORD LINE", "ORDER TYPE",
                      "VENDOR ID", "TITLE", "CATALOG KEY"])


def ParseCommandArgs():
  """
  Parses argument information passed on command line, which contains user options.
  
  Returns:
    A namespace called args populated with values for arguments.
  """
  parser = argparse.ArgumentParser(
      description="Fetch a table of upcoming recurring purchases for current FY.")
  parser.add_argument("orders_file", type=str, nargs="?",
                      default="EngEncumbrances20190723.xlsx",
                      help="Filename of current LIBSYS Open Orders report")
  parser.add_argument("expenditures_file", type=str, nargs="?",
                      default="EngExpenditures2018.xls",
                      help="Filename of last year's LIBSYS Expenditures report")
  parser.add_argument("--csv", dest="csv", action="store_true",
                      help="Output a csv file.")
  args = parser.parse_args()
  return args


def RunRecurringPurchases(orders_file, expenditures_file, csv=False):
  """
  Opens given files and outputs table of upcoming recurring purchases.

  Args:
    orders_file: (str) Path (relative or absolute) to the Open Orders report.
    expenditures_file: (str) Path to the year-ago Expenditures report.
    csv: (bool) Whether the user would like output as a csv (default is png).
  """
  df_ord = pd.read_excel(
      orders_file, sheet_name="enc_rpt1563914632", header=0)
  df_exp = pd.read_excel(
      expenditures_file, sheet_name="EngExpenditures2019", header=0)

  df_ord = AddManualOrders(df_ord)
  upcoming_orders = ParseOrdersFile(df_ord)
  upcoming_orders = FetchPreviousPrice(df_exp, upcoming_orders)
  OutputUpcomingOrders(upcoming_orders, csv)


def AddManualOrders(df_ord):
  """
  Appends "missing" orders to table of orders, due to their NOFUND status.

  Args:
    df_ord: (Pandas DataFrame) DataFrame with contents of Open Orders report.

  Returns:
    DataFrame df_ord appended with info on NOFUND orders.
  """
  nofund_ord = pd.DataFrame(
    [["965855F15", 1, "E-PACKAGE", "GRI", "SIMPLYANALYTICS", 11060871],
     ["454035F19", 1, "DATAB2", "PLANETLAB", 
         "PLANET LABS DAILY SATELLITE IMAGERY", 13157872],
     ["177114F07", 1, "E-PACKAGE", "NERL", "SAGE JOURNALS", 6847241],
     ["527176F11", 1, "E-PACKAGE", "ESRI",
         "ESRI EDUCATIONAL SITE LICENSE", 9652252]],
    columns=["ORDER ID", "ORD LINE", "ORDER TYPE",
             "VENDOR ID", "TITLE", "CATALOG KEY"])
  return df_ord.append(nofund_ord, ignore_index=True)


def ParseOrdersFile(df_ord):
  """
  Uses Open Orders report to find remaining yearly recurring purchases.

  Args:
    df_ord: (Pandas DataFrame) DataFrame with contents of Open Orders report.

  Returns:
    A DataFrame with selected info on upcoming recurring purchases.
  """
  # Select only needed columns from rows of recurring order types.
  df_ord = df_ord.loc[df_ord["ORDER TYPE"].isin(RECURRING_TYPES), OUTPUT_COLUMNS]
  
  # Flatten table into a dict, with unique orders as keys and values being lists
  # of orderlines.  We want to keep only those with a single orderline of value 1.
  order_dict = {}
  for row, order_series in df_ord.iterrows():
    if order_series["ORDER ID"] not in order_dict:
      order_dict[order_series["ORDER ID"]] = [order_series["ORD LINE"]]
    else:
      order_dict[order_series["ORDER ID"]].append(order_series["ORD LINE"])

  # Select only order rows with the single orderline 1.
  upcoming_list = []
  for order_id, ordline_list in order_dict.items():
    if len(ordline_list) == 1 and ordline_list[0] == 1:
      upcoming_list.append(order_id)
  
  df_ord = df_ord.loc[df_ord["ORDER ID"].isin(upcoming_list), :]
  logging.info("There are %d upcoming recurring orders.", df_ord.shape[0])
  return df_ord


def FetchPreviousPrice(df_exp, df_ord):
  """
  Uses Expenditures report to add last year's price to the table of upcoming purchases.
 
  Args:
    df_exp: (Pandas DataFrame) DataFrame with contents of Expenditures report.
    df_ord: (Pandas DataFrame) Selected info on upcoming recurring purchases.

  Returns:
    DataFrame of upcoming recurring purchases augmented with price column.
  """
  df_ord["PREV COST"] = pd.Series(np.zeros(df_ord.shape[0], np.float),
                                  index=df_ord.index)
  df_ord["PREV DATE PAID"] = pd.Series(["N/A"] * df_ord.shape[0],
                                       index=df_ord.index)
  
  for row, order_series in df_ord.iterrows():
    order_id = order_series["ORDER ID"]
    order_costs = df_exp.loc[df_exp["Order ID"] == order_id,
                             "Amt Paid on Fund (including tax)"]
    df_ord.loc[row, "PREV COST"] = order_costs.sum()
    order_dates = df_exp.loc[df_exp["Order ID"] == order_id, "Date to AP"]
    
    if not order_dates.empty:
      df_ord.loc[row, "PREV DATE PAID"] = order_dates.iat[0].strftime("%Y-%m-%d")
  
  return df_ord


def OutputUpcomingOrders(df_ord, csv):
  """
  Outputs complete table of recurring orders in png or csv.

  Args:
    df_ord: (Pandas DataFrame) Selected info on upcoming recurring purchases.
    csv: (bool) Whether the user would like output as a csv (default is png).
  """
  df_ord = df_ord.sort_values(
      by=["PREV DATE PAID", "TITLE"], ascending=[False, True])
  
  if csv:
    output_file = "recurring.csv"
    df_ord.to_csv(output_file, index=False, encoding="utf-8")
    logging.info("CSV file saved.")

  else:
    fig, ax = plt.subplots(figsize=(35, 20))
    ax.set_axis_off()
    table = ax.table(cellText=df_ord.values.tolist(), cellLoc="left",
                     loc="upper left",
                     colWidths=[.05, .3, .1, .1, .1, .1, .1, .1],
                     colLoc="left", colLabels=df_ord.columns, 
                     edges="BT")
    table.auto_set_font_size(False)
    table.set_fontsize(16)

    cells = table.get_celld()
    for r in range(0, df_ord.shape[0] + 1):
      for c in range(0, df_ord.shape[1]):
        cur_cell = cells[(r, c)]
        cur_cell.set_edgecolor("darkgray")
        cur_cell.set_height(0.02)
        if r == 0:
          cur_cell.set_text_props(weight="bold", color="navy")
    fig.tight_layout()
    plt.savefig("recurring.png", format="png")
    logging.info("Table saved.")
  

def main():
  logging.basicConfig(format="%(asctime)s %(message)s", level=logging.DEBUG)
  args = ParseCommandArgs()
  RunRecurringPurchases(args.orders_file, args.expenditures_file, args.csv)


if __name__ == "__main__":
  main()
