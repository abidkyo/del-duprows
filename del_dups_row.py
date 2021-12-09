#! /usr/bin/env python3
"""
Script to delete duplicates rows by the same values in a column.

- Group the data and apply "first" function.

Usage:
    python3 del_dups_row.py -v -f data.xlsx -g "First Name" -o output.xlsx
"""


import argparse

import pandas as pd


# Create argument parser
parser = argparse.ArgumentParser(description="Delete duplicates rows by the same values in a column.")
required_argument = parser.add_argument_group("required named arguments")
parser.add_argument("-v", "--verbose", action="store_true", help="verbosity")
required_argument.add_argument("-f", "--file", help="excel filename", required=True)
parser.add_argument("-s", "--sheet", default=0, help="sheet name or number")
parser.add_argument("-n", "--nrows", default=None, type=int, help="number of rows")
required_argument.add_argument("-g", "--groupby", help="groupby column", required=True)
required_argument.add_argument("-o", "--output", help="output filename", required=True)

args = parser.parse_args()

# Store arguments in variables.
verbose = args.verbose
fileName = args.file

try:
    sheet = int(args.sheet)
except ValueError:
    sheet = args.sheet

nRows = args.nrows
groupbyCol = args.groupby
outFileName = args.output


# Read excel file, assuming data on the first sheet.
df = pd.read_excel(fileName, sheet_name=sheet, nrows=nRows)
if verbose:
    print("Original Data")
    print(df)
    print()


# Apply set function to all columns except the groupbyCol column.
# Set prevent duplicates values!! :)
df_output = df.groupby(groupbyCol).agg("first").reset_index()
if verbose:
    print("Data without duplicate rows")
    print(df_output)
    print()


# Export to an excel file.
df_output.to_excel(outFileName, index=False)
