"""
Author Linnea Shieh <laiello@stanford.edu>
Copyright Stanford University 2019

Outputs various charts visualizing the state of library
collections, based on circulation reports.
"""

import argparse
import logging
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import re

from datetime import date


INPUT_COLUMNS = set(["CALL NUMBER", "PUB YR", "HOME LOC", "ITEM DATE",
                     "Last 10 yrs"])
UNUSED_LOCS = set(["TIMO-COLL", "TERM-COLL", "EQUIPMENT"])
LABEL_CALLS = set(["QC", "QA", "TA", "TK", "TN", "Z", "QD",
                  "BF", "HG", "P", "TJ", "T", "RC", "LB"])
CHART_CALLS = set(["QA", "TA", "TK", "TJ", "T", "TL"])


def ParseCommandArgs():
  """
  Parses argument information passed on command line, which contains user options.
  
  Returns:
    A namespace called args populated with values for arguments.
  """
  parser = argparse.ArgumentParser(
      description="Output charts visualizing library collections.")
  parser.add_argument("circ_file", type=str, nargs="?",
                      default="EngCircStats20190703.xlsx",
                      help="Filename for LIBSYS circ stats report.")
  args = parser.parse_args()
  return args


def RunCollectionsViz(circ_file):
  """
  Opens circ stats file and outputs collections vizualization charts.

  Args:
    circ_file: (str) Path (relative or absolute) to the Circ Stats report.
  """
  df_circ = pd.read_excel(
    circ_file, sheet_name="circ_rpt190702174508_copies_all", header=0)

  df_circ = ParseCircStatsFile(df_circ)
  OutputAccumulationChart(df_circ)


def ParseCircStatsFile(df_circ):
  """
  Performs cleaning and enhancing of circ stats data.

  Args:
    circ_file: (str) Path (relative or absolute) to the Circ Stats report.
 
  Returns:
    DataFrame with collections info needed for visualization.
  """
  # Select only needed colums for collection items. Exclude Terman,
  # Timoshenko collection items.
  df_circ = df_circ.loc[
      ~df_circ["HOME LOC"].isin(UNUSED_LOCS), INPUT_COLUMNS]
  
  # Add columns for beginning letter(s) of call number.
  df_circ["CALL LETTER"] = pd.Series(["Other"] * df_circ.shape[0],
                                     index=df_circ.index)
  for row, circ_series in df_circ.iterrows():
    letter = re.match(r"[A-Z]+", df_circ.loc[row, "CALL NUMBER"]).group(0)
    df_circ.loc[row, "CALL LETTER"] = letter

  # Add column for age.
  cur_year = date.today().year 
  df_circ["AGE"] = cur_year - df_circ["PUB YR"]
  logging.info("There are %d items in the collection.", df_circ.shape[0])
  return df_circ


def Output3dScatterplot(df_circ):
  """
  Groups stats by call number and create scatterplot of age, size, usage.
  
  Args:
    df_circ: (DataFrame) DF with stats per collection item.
  """
  grouped = df_circ.groupby("CALL LETTER").agg(
      {"AGE": "median", "CALL LETTER": "count",
       "Last 10 yrs": "sum"})
  grouped = grouped.rename(
      columns={"AGE": "median_age", "CALL LETTER": "size",
               "Last 10 yrs": "total_usage"})
  
  # Drop rows with <5 items. Use squared mean usage for point size.
  grouped = grouped.loc[grouped["size"] > 5, :]
  mean_usage = grouped["total_usage"].divide(grouped["size"]) ** 2
  
  fig, ax = plt.subplots()
  ax.scatter(grouped["median_age"], grouped["size"], c="navy", 
             s=mean_usage, alpha=0.5)
  for call in grouped.index:
    if call in LABEL_CALLS:
      std_x, std_y = -2, 12
      hg_x, hg_y = -10, 24
      p_x, p_y = 10, 22
      if call in ["HG", "BF"]:
        ax.annotate(
            call, (grouped.loc[call, "median_age"], grouped.loc[call, "size"]),
            xytext=(hg_x, hg_y), textcoords="offset points")
      elif call == "P":
        ax.annotate(
            call, (grouped.loc[call, "median_age"], grouped.loc[call, "size"]),
            xytext=(p_x, p_y), textcoords="offset points")
      else:
        ax.annotate(
            call, (grouped.loc[call, "median_age"], grouped.loc[call, "size"]),
            xytext=(std_x, std_y), textcoords="offset points")

  ax.yaxis.get_major_ticks()[0].label1.set_visible(False)
  ax.set_xlabel("Median age, years", fontsize=12),
  ax.set_ylabel("Items in collection", fontsize=12),
  ax.set_title("Terman Library age, size, and average usage")
  ax.grid(True)
  fig.tight_layout()

  plt.savefig("collections.png", format="png")
  logging.info("Scatterplot saved.")  


def OutputAccumulationChart(df_circ):
  """
  Outputs line chart depicting accumulation of various call letters over time.

  Args:
    df_circ: (DataFrame) DF with stats per collection item.
  """
  df_circ["ADDED YEAR"] = df_circ["ITEM DATE"].apply(lambda d: str(d)[0:4])
  df_circ["ADDED YEAR"] = pd.to_numeric(df_circ["ADDED YEAR"])

  fix, ax = plt.subplots()
  for call in CHART_CALLS:
    years = df_circ.loc[df_circ["CALL LETTER"] == call, "ADDED YEAR"].values
    sorted_years = np.sort(years)
    ax.step(sorted_years, np.arange(sorted_years.size), label=call)
  ax.grid(True)
  ax.legend(loc="upper left", shadow=True)
  ax.set_title("Accumulation of items over time")
  ax.set_ylabel("Items in collection")

  plt.savefig("accumulation.png", format="png")
  logging.info("Line chart saved.")


def main():
  logging.basicConfig(format="%(asctime)s %(message)s", level=logging.DEBUG)
  args = ParseCommandArgs()
  RunCollectionsViz(args.circ_file)


if __name__ == "__main__":
  main()
