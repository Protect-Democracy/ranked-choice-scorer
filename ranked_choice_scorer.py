import argparse
import json
import logging
import os.path
import random
import re

import gspread
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly
import plotly.graph_objects as go

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logging.getLogger().setLevel(logging.INFO)

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
# TODO: Make this configurable
RANGE_NAME = "Form Responses 1"


def get_spreadsheet(google_id):
    """
    Gets the spreadsheet from a Google Sheet via the supplied google_id
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=google_id, range=RANGE_NAME).execute()
        values = result.get("values", [])

        if not values:
            logging.warning("No data found.")
            return

        return pd.DataFrame.from_records(values)
    except HttpError as err:
        logging.error(err)


def clean_column_name(col):
    """Standardizes column names based on Google Forms format"""
    match = re.search(r"(?<=\[).+?(?=\])", col)
    return match.group() if match else None


def clean_data(df, questions):
    """Cleans up Google Form data so it can be processed"""
    # Grab the first row for the header
    new_header = df.iloc[0]
    # Take the data less the header row
    df = df[1:]
    # Set the header row as the df header
    df.columns = new_header
    # Remove extraneous columns
    df = df.loc[:, df.columns.str.contains("|".join(questions), regex=True)]
    # Set any rows that are blank to na
    df = df.replace("", np.NaN)
    df = df.replace("None", np.NaN)
    # Drop rows with missing data
    df = df.dropna(how="all")

    # Convert votes to int
    for col in df.columns:
        try:
            df[col] = df[col].replace("", np.nan, regex=True)
            df[col] = df[col].fillna(99)
            df[col] = pd.to_numeric(df[col])
        except ValueError:
            raise

    result = {}
    for index in questions:
        result[index] = df.filter(regex=index)
        cols = {col: clean_column_name(col) for col in result[index].columns}
        result[index].rename(columns=cols, inplace=True)
        result[index] = result[index].reset_index().iloc[:, 1:]
    return result


def vote_by_ranking(df, verbose=False):
    """Process the vote data into ranking data structures"""
    results = []
    vote_rounds = pd.DataFrame()

    # Change rows and columns to have voters as columns
    df_t = df.transpose()
    for col in df_t.columns:
        top_choice = df_t[col].min()
        top_candidate = df_t[df_t[col] == top_choice].index.tolist()[0]
        results.append(top_candidate)

    vote_rounds[0] = results

    left_voters = []
    losers = []
    for r in range(1, df.shape[1] - 1):
        # Stop loop when there are already two candidates
        if vote_rounds[r - 1].nunique() == 2:
            break

        # Start the new voting round
        vote_rounds[r] = vote_rounds[r - 1]

        # FInd out who are the potential losers
        aggre = pd.DataFrame(vote_rounds[r - 1].value_counts())
        min_vote = aggre[r - 1].min()
        potential_losers = aggre[aggre[r - 1] == min_vote].index.tolist()
        least_votes = df[potential_losers].sum().max()
        potential_losers_df = df[potential_losers]

        # Choose loser based on worse overall ranking (sum of ranking):
        sum_votes = pd.DataFrame(potential_losers_df.sum())
        least_ranking = sum_votes[0].max()
        loser = sum_votes[sum_votes[0] == least_ranking].index.tolist()[0]
        if verbose:
            print(f"Loser of round {r} is {loser}")
        losers.append(loser)

        # Determining who their votes go to
        voters_non_selected = df[df[loser] == 1].index.tolist()
        for voter in voters_non_selected:
            left_voters.append(voter)

        votes_to_distribute = df.iloc[list(set(left_voters)), :]
        votes_to_distribute = votes_to_distribute.loc[
            :, ~votes_to_distribute.columns.isin(losers)
        ]
        votes_to_distribute_t = votes_to_distribute.transpose()
        for votr in votes_to_distribute_t.columns:
            nxt_choice = votes_to_distribute_t[votr].min()
            vote_goes_to = votes_to_distribute_t[
                votes_to_distribute_t[votr] == nxt_choice
            ].index.tolist()[0]
            if verbose:
                print(f"Vote goes to {vote_goes_to}")

            # Changing their votes
            vote_rounds.loc[votr, r] = vote_goes_to
        if verbose:
            print("\n")

    col_rounds = vote_rounds.columns.tolist()
    vote_rounds["value"] = [1 for x in range(vote_rounds.shape[0])]
    return (vote_rounds, col_rounds)


def select_winner(vote_rounds):
    """Select and output the winner to text"""
    final_count = pd.DataFrame(vote_rounds.iloc[:, -2:-1].value_counts()).reset_index()
    final_count.columns = ["candidate", "final_votes"]
    winner = final_count[final_count.final_votes == final_count.final_votes.max()][
        "candidate"
    ].tolist()
    if len(winner) > 1:
        print("There is a draw")
    else:
        print(f"The final winner is… {winner[0]}!\n")
    print(final_count)
    print("\n")


def get_sankey_dataframe(vote_rounds, col_rounds):
    """Generate a Plotly-compatible dataframe for a Sankey chart"""
    df_sankey = vote_rounds.groupby(col_rounds).count().reset_index()
    for col in col_rounds:
        df_sankey[col] = df_sankey[col] + str(col)
    return df_sankey


def generate_sankey(df, cat_cols=[], value_cols="", title="Sankey Diagram"):
    """
    Configure the options for the Sankey chart
    Adapted from https://gist.github.com/ken333135/09f8793fff5a6df28558b17e516f91ab
    """
    # Color palettes from https://www.schemecolor.com
    with open("color_palettes.json", "r") as fp:
        color_palettes = json.load(fp)
    label_list = []
    for cat_col in cat_cols:
        label_list_temp = list(set(df[cat_col].values))
        label_list = label_list + label_list_temp

    # remove duplicates from label_list
    label_list = list(dict.fromkeys(label_list))

    color_list = []
    color_dict = {}
    color_palette = color_palettes[random.randint(0, len(color_palettes) - 1)]

    for item in label_list:
        if item[:-1] not in color_dict.keys():
            color_dict[item[:-1]] = color_palette.pop(
                random.randint(0, len(color_palette) - 1)
            )
        color_list.append(color_dict[item[:-1]])

    # transform df into a source-target pair
    for i in range(len(cat_cols) - 1):
        if i == 0:
            source_target_df = df[[cat_cols[i], cat_cols[i + 1], value_cols]]
            source_target_df.columns = ["source", "target", "count"]
        else:
            temp_df = df[[cat_cols[i], cat_cols[i + 1], value_cols]]
            temp_df.columns = ["source", "target", "count"]
            source_target_df = pd.concat([source_target_df, temp_df])
        source_target_df = (
            source_target_df.groupby(["source", "target"])
            .agg({"count": "sum"})
            .reset_index()
        )

    # add index for source-target pair
    source_target_df["source_id"] = [
        label_list.index(x) for x in source_target_df["source"]
    ]
    source_target_df["target_id"] = [
        label_list.index(x) for x in source_target_df["target"]
    ]

    link_color_list = []
    for item in source_target_df["target"]:
        link_color_list.append(color_dict[item[:-1]].replace("0.8", "0.3"))

    # Remove the appended numbers from the label list
    label_list = [item[:-1] for item in label_list]

    # creating the sankey diagram
    data = dict(
        type="sankey",
        node=dict(
            pad=50,
            thickness=20,
            line=dict(color="black", width=0.5),
            label=label_list,
            color=color_list,
        ),
        link=dict(
            source=source_target_df["source_id"],
            target=source_target_df["target_id"],
            value=source_target_df["count"],
            color=link_color_list,
        ),
    )

    layout = dict(title=title, font=dict(size=10))

    fig = dict(data=[data], layout=layout)
    return fig


def chart_votes(df_sankey, col_rounds, question):
    """Generate and display the charts"""
    sankey_title = f"{question} – Vote by Ranking"
    sankey_fig = generate_sankey(
        df_sankey, cat_cols=col_rounds, value_cols="value", title=sankey_title
    )
    # plotly.offline.plot(fig, validate=False)
    fig = go.Figure(sankey_fig)
    fig.update_layout(width=int(1200))
    fig.add_annotation(x=0, y=1.05, showarrow=False, text="First round")
    fig.add_annotation(x=1, y=1.05, showarrow=False, text="Final round")
    # TODO: Save charts to files?
    fig.show()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--questions",
        nargs="+",
        type=str,
        required=True,
        help="List of questions to parse",
    )
    parser.add_argument(
        "--googleid",
        type=str,
        required=True,
        help="ID of the source results spreadsheet",
    )
    parser.add_argument(
        "--chart",
        dest="chart",
        action="store_true",
        default=False,
    )
    parser.add_argument(
        "--verbose",
        dest="verbose",
        action="store_true",
        default=False,
    )
    args = parser.parse_args()

    df = get_spreadsheet(args.googleid)
    result = clean_data(df, args.questions)

    for key, value in result.items():
        (vote_rounds, col_rounds) = vote_by_ranking(value, verbose=args.verbose)
        select_winner(vote_rounds)
        df_sankey = get_sankey_dataframe(vote_rounds, col_rounds)
        if args.chart:
            chart_votes(df_sankey, col_rounds, key)
