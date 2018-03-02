# -*- coding: utf-8 -*-
"""
Created on Tues Nov 14 17::00 2017

@author: agatab
"""
import json
import os
import numpy as np
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta, MO
from dateutil import parser
import jem_funcs

"""-------------------------FUNCTION DEFINITIONS--------------------------------"""

def attempt_count(series):
    """Return number of items in a pandas group
    Parameters
    ----------
    series : pandas series
    
    Returns
    -------
    int : length of series
    """
    return len(series)

def success_count(series):
    """Return number of items in a pandas group that are successful recordings.
    Parameters
    ----------
    series : pandas series
    
    Returns
    -------
    int : number of successes in series
    """
    return sum(series.isin(["SUCCESS", "SUCCESS (high confidence)", "SUCCESS (low confidence)"]))

def fail_count(series):
    """Return number of items in a pandas group that are failed recordings.
    Parameters
    ----------
    series : pandas series
    
    Returns
    -------
    int : number of failures in series
    """
    return sum(series == "FAILURE")     

def issue_counter(prep):
    """Return fraction of slices in prep that were marked with various slice issues.
    
    Parameters
    ----------
    prep : pandas dataframe row with metadata for a given recording prep
    
    Returns
    -------
    wov_frac : float
        Fraction of slices in the prep that had a "Wave of Death" tag.        
    uneven_frac : float
        Fraction of slices in the prep that had an "Uneven Thickness" tag.       
    damaged_frac : float
        Fraction of slices in the prep that had a "Damaged" tag.
    """
    
    num_attempts = len(prep["slice_name"])
    comments = prep["slice_quality"]
    wov_frac = len([c for c in comments if "Wave of Death" in c]) / float(num_attempts)
    uneven_frac = len([c for c in comments if "Uneven Thickness" in c]) / float(num_attempts)
    damaged_frac = len([c for c in comments if "Damaged" in c]) / float(num_attempts)
    return wov_frac, uneven_frac, damaged_frac


def save_xlsx(prep_df, slice_df, dirname, spreadname, norm_d, head_d1, head_d2, issue_d):
    """Save an excel spreadsheet from dataframe
    
    Parameters
    ----------
    prep_df : pandas dataframe
    slice_df : pandas dataframe
    dirname : string
    spreadname : string
    norm_d, head_d1, head_d2, issue_d: dictionaries with formatting
    
    Returns
    -------
    Saved .xlsx file with name spreadname in directory dirname.
    """
        
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(os.path.join(dirname, spreadname), engine='xlsxwriter', date_format='mm/dd/yy')
    
    # Convert the dataframe to an XlsxWriter Excel object.
    prep_df.to_excel(writer, sheet_name='prep_summary', index=False) 
    slice_df.to_excel(writer, sheet_name='slice_summary', index=False)    
    
    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet1 = writer.sheets['prep_summary']
    worksheet2 = writer.sheets['slice_summary']
    
    norm_fmt = workbook.add_format(norm_d)
    head_fmt1 = workbook.add_format(head_d1)
    head_fmt2 = workbook.add_format(head_d2)
    issue_fmt = workbook.add_format(issue_d)

    worksheet1.set_column('A:N', 18, norm_fmt)
    worksheet2.set_column('A:N', 18, norm_fmt)

    # Write the column headers with the defined format.
    for col_num, value in enumerate(prep_df.columns.values):
        worksheet1.write(0, col_num, value, head_fmt1)
    for col_num, value in enumerate(slice_df.columns.values):
        worksheet2.write(0, col_num, value, head_fmt2)

    worksheet1.conditional_format("G2:I{}".format(len(prep_df)+1), {'type':'cell','criteria':'>','value':0.0,'format':issue_fmt})
    try:
        writer.save()
    except IOError:
        print "\nOh no! Unable to save spreadsheet :(\nMake sure you don't already have a file with the same name opened."


"""-----------------------CONSTANTS-----------------------------"""

output_name = "ps_slicehealth_report.csv" 
DEFAULT_DIR = "//allen/programs/celltypes/workgroups/279/Patch-Seq/all-metadata-files"
REPO_DIR = os.path.abspath(os.path.join(os.getcwd() ,"../"))
OUTPUT_DIR =  os.path.join(REPO_DIR, "reports")
CONSTANTS_DIR = os.path.join(REPO_DIR, "jem-constants")

col_dict = {"mean":"mean_health",
            "sliceQuality":"slice_quality",
            "acsfProductionDate":"acsf_date",
            "limsSpecName":"slice_name", 
            "limsPrepName":"prep_name",
            "rigOperator":"user", 
            "rigNumber":"rig"}

slice_output_cols = ["day", "slice_name", "time", "mean_health", "slice_success_rate", "slice_quality", "attempt_count", "user", "rig", "acsf_date"]

prep_output_cols = ["day", "prep_name", "first_time", "last_time", "mean_health", "slice_success_rate", "wov_fraction", "uneven_fraction", "damaged_fraction", "attempt_count",  "user", "rig", "acsf_date", "slice_name"]

# xlsx formatting constants
norm_d = {"font_name":"Arial",
          "font_size":10,
          "align":"left",
          "bold": False,
          "num_format":"0.00"}
head_d1 = norm_d.copy(); head_d2 = norm_d.copy(); issue_d = norm_d.copy()
head_d1["bg_color"] = "#998ec3"
head_d2["bg_color"] = "#f1a340"
issue_d["bg_color"] = "fee0d2"

end_day = datetime.today().date()
start_day = end_day - relativedelta(weekday=MO(-1))
start_day_str, end_day_str = [x.strftime("%y%m%d") for x in [start_day, end_day]]

"""--------------Import PatchSeq user and roi info from csv files-----------------"""

roi_df = pd.read_csv(os.path.join(CONSTANTS_DIR, "roi_info.csv"))
roi_df.set_index(keys="acronym", drop=True, inplace=True)

u = pd.read_csv(os.path.join(CONSTANTS_DIR, "ps_user_info.csv"))
login_to_user = u.set_index("login").to_dict()["p_user"]
name_to_login = u.set_index("name").to_dict()["login"]


"""------------------------Ask for user input----------------"""
str_prompt1 = "\nWould you like to report on samples collected between %s and %s? (y / n): "  %(start_day_str, end_day_str)
valid_vals = ["y", "n"]
str_prompt2 = "Please enter report start date (YYMMDD on or after 171110): "
str_prompt3 = "Please enter report end date (YYMMDD): "
response1 = "\nPlease try again..."
response2 = "\nPlease try again... date should be YYMMDD"

default_dates_state = jem_funcs.validated_input(str_prompt1, response1, valid_vals)
if default_dates_state == "n":
    start_day_str = jem_funcs.validated_date_input(str_prompt2, response2, valid_options=None)
    end_day_str = jem_funcs.validated_date_input(str_prompt3, response2, valid_options=None)
    start_day, end_day = [datetime.strptime(x, "%y%m%d").date() for x in [start_day_str, end_day_str]]
    
print("Generating report for samples collected between %s and %s..." %(start_day_str, end_day_str))
dated_output_name = "%s-%s_%s.csv" %(start_day_str, end_day_str, output_name[0:-4])
dated_output_xlsx = "%s_%s_%s.xlsx" %(start_day_str, end_day_str, output_name[0:-4])

"""-------------------------------------------------------------------------------"""

"""Get Patch-Seq JSON pathnames of files that have been created since the report start date
(with a 3 day buffer)"""

delta_mod_date = (datetime.today().date() - start_day).days + 3
json_paths = jem_funcs.get_jsons(dirname=DEFAULT_DIR, expt="PS", delta_days=delta_mod_date)


"""Flatten data in recent JSON files and output successful experiments (with tube IDs) in a dataframe. """

json_df = pd.DataFrame()
for json_path in json_paths:
    with open(json_path) as data_file:
        slice_info = json.load(data_file)
        if jem_funcs.is_field(slice_info, "formVersion"):
            jem_version = slice_info["formVersion"]
        else:
            jem_version = "1.0.0"
        flat_df = jem_funcs.flatten_attempts(slice_info, jem_version)
        json_df = pd.concat([json_df, flat_df], axis=0)

"""Remove tissue touches and other non-data."""
json_df = json_df[json_df["approach.pilotName"]!="Tissue_Touch"]
json_df = json_df[json_df["rigOperator"]!="davidre"]

"""Ping LIMS for proper prep name. Clean up some columns."""
json_df.loc[:,"limsPrepName"] = json_df["limsSpecName"].apply(jem_funcs.get_prep_from_specimen_name)

json_df.loc[:, "date_dt"] = json_df["date"].apply(lambda x: parser.parse(x))
json_df.loc[:, "acsfProductionDate"] = json_df["acsfProductionDate"].apply(lambda x: parser.parse(x).strftime("%Y-%m-%d"))
json_df.loc[:, "day"] = json_df["date_dt"].apply(lambda x: x.strftime("%Y-%m-%d"))
json_df.loc[:, "time"] = json_df["date_dt"].apply(lambda x: x.time().strftime("%H:%M"))
json_df.replace({"rigOperator": name_to_login}, inplace=True)
json_df.loc[:,"approach.sliceHealth"] = json_df["approach.sliceHealth"].apply(lambda x: np.float(x))
json_df = jem_funcs.select_report_date_attempts(json_df, report_dt=[start_day, end_day])
json_df.reset_index(drop=True, inplace=True)

"""---------GENERATE PER SLICE SUMMARY------------"""
slice_summary = json_df.groupby(by=["limsSpecName"]).agg(
        {"date":attempt_count,
         "status":[success_count, fail_count],
         "limsSpecName":lambda x: list(set(x))[0],
         "limsPrepName":lambda x: list(set(x))[0],
         "day":lambda x: list(set(x))[0],
         "time":lambda x: list(set(x))[0],
         "rigOperator":lambda x: list(set(x))[0],
         "rigNumber":lambda x: list(set(x))[0],
         "acsfProductionDate":lambda x: list(set(x))[0],
         "sliceQuality":lambda x: list(set(x))[0], 
         "approach.sliceHealth":np.mean
         }, asindex=False)

slice_summary.columns =  [c[0] if c[1] == "<lambda>" else c[1] for c in slice_summary.columns]
slice_summary.loc[:,"slice_success_rate"] = slice_summary["success_count"] / slice_summary["attempt_count"]
slice_summary.rename(columns=col_dict, inplace=True)
slice_summary.drop(labels=["success_count", "fail_count"], axis=1, inplace=True)


"""---------GENERATE PER PREP SUMMARY------------"""

prep_summary = slice_summary.groupby(by=["prep_name"]).agg(
        {"slice_success_rate": lambda x: np.mean(list(x)),
         "prep_name":lambda x: list(set(x))[0],
         "slice_name":lambda x: list(set(x)),
         "day":lambda x: ", ".join(set(x)),
         "time":lambda x: list(set(x)),
         "user":lambda x: ", ".join(set(x)),
         "rig":lambda x: ", ".join(set(x)),
         "attempt_count":lambda x: sum(list(x)),
         "acsf_date":lambda x: ", ".join(list(set(x))),
         "slice_quality":lambda x: list(x[~x.isnull()]),
         "mean_health":np.mean
         }, asindex=False)

prep_summary.loc[:,"first_time"], prep_summary.loc[:,"last_time"]  = zip(*prep_summary["time"].apply(lambda x: (min(x), max(x))))
prep_summary.loc[:,"wov_fraction"], prep_summary.loc[:,"uneven_fraction"], prep_summary.loc[:,"damaged_fraction"] = zip(*prep_summary.apply(lambda x: issue_counter(x), axis=1))
prep_summary.loc[:,"slice_name"] = prep_summary["slice_name"].apply(lambda x: ", ".join(x))


"""-----------SAVE SPREADSHEET OUTPUT-----------------"""
prep_fmt =  prep_summary[prep_output_cols].sort_values(by=["day", "first_time"])
slice_fmt = slice_summary[slice_output_cols].sort_values(by=["day", "slice_name"])

try:
    save_xlsx(prep_fmt, slice_fmt, OUTPUT_DIR, dated_output_xlsx, norm_d, head_d1, head_d2, issue_d)
except IOError:
        print "\nOh no! Unable to save spreadsheet :(\nMake sure you don't already have a file with the same name opened."