#!/usr/bin/env python

import argparse

import sys, os
import sys
import os
from datetime import datetime, timedelta
import pandas as pd
from alpha_vantage.timeseries import TimeSeries
@@ -33,8 +34,9 @@ def main():
    while I saw my GME shares tanking. But hey, I like the stock.
    """

    # Enable VT100 Escape Sequence for WINDOWS 10 Ver. 1607
    if sys.platform == "win32":
        os.system("")  # enable VT100 Escape Sequence for WINDOWS 10 Ver. 1607
        os.system("")

    s_ticker = ""
    s_start = ""
    df_stock = pd.DataFrame()
    s_interval = "1440min"
    # Set stock by default to speed up testing
    # s_ticker = "BB"
    # ts = TimeSeries(key=cfg.API_KEY_ALPHAVANTAGE, output_format='pandas')
    # df_stock, d_stock_metadata = ts.get_daily_adjusted(symbol=s_ticker, outputsize='full')
    # df_stock.sort_index(ascending=True, inplace=True)
    # s_start = datetime.strptime("2020-06-04", "%Y-%m-%d")
    # df_stock = df_stock[s_start:]
    # Add list of arguments that the main parser accepts
    menu_parser = argparse.ArgumentParser(add_help=False, prog="gamestonk_terminal")
    choices = [
        "help",
        "quit",
        "q",
        "clear",
        "load",
        "candle",
        "view",
        "export",
        "disc",
        "mill",
        "sen",
        "res",
        "fa",
        "ta",
        "dd",
        "pred",
        "ca",
    ]
    menu_parser.add_argument("opt", choices=choices)
    completer = NestedCompleter.from_nested_dict({c: None for c in choices})
    # Print first welcome message and help
    print("\nWelcome to Gamestonk Terminal 🚀\n")
    should_print_help = True
    # Loop forever and ever
    while True:
        main_cmd = False
        if should_print_help:
            print_help(s_ticker, s_start, s_interval, b_is_stock_market_open())
            should_print_help = False
        # Get input command from user
        if session and gtff.USE_PROMPT_TOOLKIT:
            as_input = session.prompt(f"{get_flair()}> ", completer=completer)
        else:
            as_input = input(f"{get_flair()}> ")
        # Is command empty
        if not as_input:
            print("")
            continue
        # Parse main command of the list of possible commands
        try:
            (ns_known_args, l_args) = menu_parser.parse_known_args(as_input.split())
        except SystemExit:
            print("The command selected doesn't exist\n")
            continue
        b_quit = False
        if ns_known_args.opt == "help":
            should_print_help = True
        elif (ns_known_args.opt == "quit") or (ns_known_args.opt == "q"):
            break
        elif ns_known_args.opt == "clear":
            s_ticker, s_start, s_interval, df_stock = clear(
                l_args, s_ticker, s_start, s_interval, df_stock
            )
            main_cmd = True
        elif ns_known_args.opt == "load":
            s_ticker, s_start, s_interval, df_stock = load(
                l_args, s_ticker, s_start, s_interval, df_stock
            )
            main_cmd = True
        elif ns_known_args.opt == "candle":
            if s_ticker:
                candle(
                    s_ticker,
                    (datetime.now() - timedelta(days=180)).strftime("%Y-%m-%d"),
                )
            else:
                print(
                    "No ticker selected. Use 'load ticker' to load the ticker you want to look at."
                )
            main_cmd = True
        elif ns_known_args.opt == "view":
            if s_ticker:
                view(l_args, s_ticker, s_start, s_interval, df_stock)
            else:
                print(
                    "No ticker selected. Use 'load ticker' to load the ticker you want to look at."
                )
            main_cmd = True
        elif ns_known_args.opt == "export":
            export(l_args, df_stock)
            main_cmd = True
        elif ns_known_args.opt == "disc":
            b_quit = dm.disc_menu()
        elif ns_known_args.opt == "mill":
            b_quit = mill.papermill_menu()
        elif ns_known_args.opt == "sen":
            b_quit = sm.sen_menu(s_ticker, s_start)
        elif ns_known_args.opt == "res":
            b_quit = rm.res_menu(s_ticker, s_start, s_interval)
        elif ns_known_args.opt == "ca":
            b_quit = cam.ca_menu(df_stock, s_ticker, s_start, s_interval)
        elif ns_known_args.opt == "fa":
            b_quit = fam.fa_menu(s_ticker, s_start, s_interval)
        elif ns_known_args.opt == "ta":
            b_quit = tam.ta_menu(df_stock, s_ticker, s_start, s_interval)
        elif ns_known_args.opt == "dd":
            b_quit = ddm.dd_menu(df_stock, s_ticker, s_start, s_interval)
        elif ns_known_args.opt == "pred":
            if not gtff.ENABLE_PREDICT:
                print("Predict is not enabled in feature_flags.py")
                print("Prediction menu is disabled")
                print("")
                continue
            try:
                # pylint: disable=import-outside-toplevel
                from gamestonk_terminal.prediction_techniques import pred_menu as pm
            except ModuleNotFoundError as e:
                print("One of the optional packages seems to be missing")
                print("Optional packages need to be installed")
                print(e)
                print("")
                continue
            except Exception as e:
                print(e)
                print("")
                continue
            if s_interval == "1440min":
                b_quit = pm.pred_menu(df_stock, s_ticker, s_start, s_interval)
            # If stock data is intradaily, we need to get data again as prediction
            # techniques work on daily adjusted data. By default we load data from
            # Alpha Vantage because the historical data loaded gives a larger
            # dataset than the one provided by quandl
            else:
                try:
                    ts = TimeSeries(
                        key=cfg.API_KEY_ALPHAVANTAGE, output_format="pandas"
                    )
                    # pylint: disable=unbalanced-tuple-unpacking
                    df_stock_pred, _ = ts.get_daily_adjusted(
                        symbol=s_ticker, outputsize="full"
                    )
                    # pylint: disable=no-member
                    df_stock_pred = df_stock_pred.sort_index(ascending=True)
                    df_stock_pred = df_stock_pred[s_start:]
                    b_quit = pm.pred_menu(
                        df_stock_pred, s_ticker, s_start, s_interval="1440min"
                    )
                except Exception as e:
                    print(e)
                    print("Either the ticker or the API_KEY are invalids. Try again!")
                    b_quit = False
                    return
        else:
            print("Shouldn't see this command!")
            continue
        if b_quit:
            break
        else:
            if not main_cmd:
                should_print_help = True
    print(
        "Hope you enjoyed the terminal. Remember that stonks only go up. Diamond hands.\n"
    )
if __name__ == "__main__":
    main()