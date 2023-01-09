import string
import warnings
import xlsxwriter

import numpy as np
import pandas as pd
import seaborn as sns
import datetime as dt

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

warnings.filterwarnings(action='ignore')

sns.set_style('whitegrid')

##########################################################################################################
# ------------------------------------------------------------------------------------------------------ #

class attribution():
    
    def __init__(self, filenames, dates, fx_filenames, mjt_rates, us_rates, alm_hedging, fx_hedging):
        self.filenames = filenames
        self.dates = dates
        self.fx_filenames = fx_filenames
        self.mjt_rates = mjt_rates
        self.us_rates = us_rates
        self.alm_hedging = alm_hedging
        self.fx_hedging = fx_hedging
# ------------------------------------------------------------------------------------------------------ #
# Millions Notation
    mm = 10**6
# ------------------------------------------------------------------------------------------------------ #
# The USDJPY CURNCY Value
    bbg_value = 131.121
# ------------------------------------------------------------------------------------------------------ #
# SUMPRODUCT / SUM
    sumprod_sum = lambda self, x, y: sum([float(i) * \
                                          float(j) for i, j in zip(x.replace('---', '0'),
                                                                   y.replace('---', '0'))]) / sum(y)
# ------------------------------------------------------------------------------------------------------ #
# Data sheets - obtains data from Clearwater Analytics        

    def data_sheet(self, filename, date):
        
        """
        Parameters
        ----------
        filename:   the name of the spreadsheet in XYZ filepath
                        - filepath to be altered at a later stage.
        date:       the date of the spreadsheet that has been extracted.
        ----------

        Return
        ------
        d: a dataframe of the spreadsheet
        ------
        """
        
        d = pd.read_excel(filename, engine="openpyxl")
        # b = pd.read_csv(bbg_data)
        # BBG data would match the data presented in Kiku
        acc = d.iloc[1, 1]
        # Looks at the account type
        filedate = d.iloc[2, 1]
        # Looks at the date of when the file was created

        if date == filedate:
            d.columns = [*[i for i in d.iloc[5, :]]]
            d = d.iloc[6:-8, :]
            if acc == 'DL-Kiku-CA-JPM-FI (294204)':
            # The Kiku Fixed Income Account

            # Have to bring in MMFund at some point

            # Get rid of any data points that are:
                # Not in the currency which we want
                # Missing ISIN labels
                # That are money market funds
                d = d[(d['ISIN'] != '---') & (d['Security Type'] != 'MMFUND')]
                # MMFUND will need to be eventually incorporated

            elif acc == 'DL-Kiku-Surplus-JPM-FX Forward (308715)':
            # The Kiku FX Forward Account
                d = d
                
            else:
                pass

            d = d.reset_index(drop=True)
            d.index.names = [filedate]

            return d

        else:
            return print('Date specified is not the same as the date in the file.\n'\
                         'Please make sure that the date is in MM/DD/YYYY format.\n'\
                         f'The date specified in the entered workbook is {filedate}.')
# ------------------------------------------------------------------------------------------------------ #
# FX Spot Attribution Sheet #
        
    def fx_spot_sheet(self):

        """
        Parameters
        ----------
        filenames:    a list of the file names required
        dates:        a list of the dates relative to the file names
        fx_filenames: a list of file names for the FX forwards
        ----------

        Return
        ------
        df1: a dataframe of USDJPY CURNCY, MtM USD Bond Portfolio, and MtM JPY
        s1:  a series of FX P&L Bonds, FX Relative Performance
        ------
        """

        MtM_usd_bp, MtM_jpy = [], []

        usdjpy_ccy = [144.74, 131.121]

        for filename, date, i in zip(self.filenames, self.dates, np.arange(0, 2)):

            df = self.data_sheet(filename, date)
            usd_df = df[df['Currency'] == 'USD']
            # Dataframe of values with respect to the file name and the date
            bmv_a = usd_df['Base Market Value + Accrued']
            # Base Market Value + Accrued
            mv = usd_df['Market Value']
            # Market Value
            ab = usd_df['Accrued Balance']
            # Accrued Balance

            MtM_usd_bp.append(sum(mv) + sum(ab))
            MtM_jpy.append(usdjpy_ccy[i] * MtM_usd_bp[i])

        fx_rl_perf = usdjpy_ccy[1]/usdjpy_ccy[0] -1
        # FX Relative Performance

        fx_pl_b = [sum(MtM_jpy)/len(MtM_jpy) * fx_rl_perf]
        # FX P&L Bond

        fx_pl_b.append(fx_pl_b[0]/usdjpy_ccy[-1])
        #  The row below the FX P&L Bond value — Not really certain what this value is called # 

        df1 = pd.DataFrame({
            'USDJPY CURNCY': usdjpy_ccy,
            'MtM USD Bond Portfolio': MtM_usd_bp,
            'MtM JPY': MtM_jpy
        }, index = dates)

        s1 = pd.DataFrame({
            'FX Relative Performance': [fx_rl_perf, '-'],
            'FX P&L Bond': fx_pl_b

        }, index = dates)
# ------------------------------------------------------------------------------------------------------ #
# FX Forwards Calculations

        def fx_forwards(self):

            """
            Parameters
            ----------
            fx_filenames: a list of file names for the FX forwards
            ----------

            Return
            ------
            l: a list of FX Forward values based off the "Monthly 
               Pending FX by Maturity" worksheet received from 
               M. Murray.
            ------
            """

            l = []

            for i in self.fx_filenames:

                d = pd.read_excel(i, engine='openpyxl')
                date = d.iloc[0:, 0][0][11:].replace(' ', '')
                d.columns = d.iloc[3].values
                d = d[4:].reset_index(drop=True)
                l.append(sum(d['Unrealised G/L Total Base']))

            return l

        fx_fwd = sum(fx_forwards(self))

        s1['FX Forward'] = [fx_fwd, fx_fwd/usdjpy_ccy[-1]]

        return df1, s1
# ------------------------------------------------------------------------------------------------------ #
# Curve Attribution Sheet # 
    
    def curve_attribution(self):

        """
        Parameters
        ----------
        mjt_rates:   a list of MoF, JGB, and TONAR rates
        us_rates:    a list of UST and SOFR rates
        alm_hedging: name of the ALM Hedging file
        fx_hedging:  name of the FX Hedging file
        ----------

        Return
        ------
        alm_hedge_dataframe: the ALM Hedging dataframe containing
                                 DV01 BEL
                                 DV01 JGB
                                 DV01 Swap

        fx_hedge_dataframe: the FX Hedging dataframe containing
                                DV01 Bond
                                DV01 Swaps

        ust_sofr_diff_df:   the UST and SOFR rates and their differences in a dataframe
                                UST Rates
                                SOFR Rates
                                the difference in UST rates between time 0 and time 1
                                the difference in SOFR rates between time 0 and time 1

        mof_jgb_tnr_diff_df: the MoF, JGB, and TONAR rates and their differences in a dataframe
                                MoF Rates
                                JGB Rates
                                TONAR Rates
                                the difference in MoF rates between time 0 and time 1
                                the difference in JGB rates between time 0 and time 1
                                the difference in TONAR rates between time 0 and time 1

        alm_hedging_pnl_df: the ALM Hedging Dataframe with P&L, variables listed below
                                DV01 BEL
                                DV01 JGB
                                DV01 Swaps
                                P&L BEL
                                P&L JGB
                                P&L Swap

        fx_hedging_pnl_df:  the FX Hedging Dataframe with P&L, variables listed below
                                DV01 Bond
                                DV01 Swaps
                                P&L Bond
                                P&L Swaps

        bs_ts:               a dataframe containing
                                P&L USD Bond Portfolio
                                P&L USD Swap Portfolio
                                Treasury Change
                                Swap Change

        ljj_mjs:             a dataframe containing
                                P&L Liabilities
                                P&L JGB + JPY Swap Portfolio
                                MoF Change
                                JGB Change
                                Swap Change

        te_hedge_alm:        the ALM Hedging dataframe containing
                                P&L BEL
                                P&L JGB
                                P&L Swap

        te_hedge_fx:         the FX Hedging dataframe containing
                                P&L Bond
                                P&L Swap
        ------
        """
# ------------------------------------------------------------------------------------------------------ #
# Reads the Rates Data obtained from Bloomberg
    # UST/SOFR and MoF/JGB/TONAR Data
        
        def read_bbg_data_rates(self, filename):

            """
            Parameters
            ----------
            filename: the name of the spreadsheet in XYZ filepath
                      - filepath to be altered at a later stage.
            ----------

            Return
            ------
            d: a dataframe of the spreadsheet - the values in the dataframe
               are all floats
            ------
            """

            d = pd.read_excel(f'{filename}')
            # Reads the excel file
            cols = ['Dates', 
                    'Rates', 
                    *[i[7:10].replace(' ', '') for i in d.iloc[1, 2:].values[1:]]]
            # Creates the columns 

            d = d[3:].dropna(axis=1)
            # Drop NaN columns (should be only 1 at the start)

            d.columns = cols
            # Sets the columns equal to our created columns variable

            d['Dates'] = d['Dates'].replace([i for i in d['Dates']], 
                                            [dt.datetime.date(i) for i in d['Dates']]) 
            # Replaces the datetime objects in dates column to just dates, without the
            # specific times

            return d.set_index('Dates')
            # Sets the index of the dataframe to dates and returns the dataframe

        mjt_rates_df_1 = read_bbg_data_rates(self, self.mjt_rates[0])
        # This dataframe should contain all the data from the first MJT rates variable
        # from the mjt_rates variable - the earliest date - which would include:
            # MoF Rates
            # JGB Rates
            # TONAR Rates
        mjt_rates_df_2 = read_bbg_data_rates(self, self.mjt_rates[1])
        # This dataframe should contain all the data from the second MJT rates variable
        # from the mjt_rates variable - the later date - which would include:
            # MoF Rates
            # JGB Rates
            # TONAR Rates

        us_rates_df_1 = read_bbg_data_rates(self, self.us_rates[0])
        # This dataframe should contain all the data from the first US rates variable
        # from the us_rates variable - the earliest date - which would include:
            # UST Rates
            # SOFR Rates
        us_rates_df_2 = read_bbg_data_rates(self, self.us_rates[1])
        # This dataframe should contain all the data from the second US rates variable
        # from the us_rates variable - the later date - which would include:
            # UST Rates
            # SOFR Rates
# ------------------------------------------------------------------------------------------------------ #
# Reads ALM and FX Hedging Data obtained from M. Murray
    # ALM and FX DV01s 
    
        def read_data_hedging(self, filename):

            """
            Parameters
            ----------
            filename:   the name of the spreadsheet in XYZ filepath
                            - filepath to be altered at a later stage.
            ----------

            Return
            ------
            d: a dataframe of the spreadsheet - the values in the dataframe
               are all integers
            ------
            """

            d = pd.read_excel(f'{filename}').dropna(axis=1)
            cols = [d.loc[0].values[0], *[f'Tenor {int(i)}Y' for i in d.loc[0].values[1:]]]
            # This line is required because, for some odd reason, year 10 is read as a float
            # and all other numbers are read as integers.
            # To change this and to add a 'Y' to the end of this, the loop was required.
            d.columns = cols
            # Changes the columns of the dataframe to the define cols variable above
            d = d[1:]
            # Shifts the dataframe down and ignores the first row

            return d.set_index(d.columns[0])
            # Returns the dataframe and sets the index as the first item of the 
            # dataframe's column.

        alm_hedge_dataframe = read_data_hedging(self, self.alm_hedging)
        # Creates the ALM Hedging data with:
            # DV01 BEL
            # DV01 JGB
            # DV01 Swap
        fx_hedge_dataframe = read_data_hedging(self, self.fx_hedging)
        # Creates the FX Hedging data with:
            # DV01 Bond
            # DV01 Swaps
# ------------------------------------------------------------------------------------------------------ #
# Differences DataFrames
    # UST/SOFR Differences
    # MoF/JGB/TONAR Differences

        def differences_in_rates(self,
                                 df1: pd.DataFrame,
                                 df2: pd.DataFrame):

            """
            Parameters
            ----------
            df1: The first rates dataframe - at the first date
            df2: The second rates dataframe - at the final date

            Return
            ------
            dataframe: a dataframe of the differences in rates
            diff:      the difference in days is returned to be referred to 
            """

            fv = lambda x, i: x.values[i][1:].astype(float)
            # Gets values from the x variable (list/series) and takes all values apart
            # form the first value and changes them to floats.

            l = []
            [l.append([*[(y-x)*100 for x, y in zip(fv(df1, i), 
                                                   fv(df2, i))], 
                       0.0]) for i in np.arange(0, len(df1.index))]
            # Creates a list of values that calculates the difference between the two files
            # and gives a percentage output (multiplied by 100)

            diff_days = lambda x, y: y.index - x.index
            # Calculates the difference in the number of days between the two files

            diff = f'{(diff_days(df1, df2)).days.to_list()[0]} Days Difference'
            # Ensures that only the number of days are returned, and not the hours and minutes

            dt_mjt_names = ['MoF rates', 'JGB rates', 'TONAR rates']
            # MoF, JGB, and TONAR rates
            dt_us_names = ['UST rates', 'SOFR rates']
            # UST and SOFR rates

            d = pd.DataFrame(l, columns=(*[i for i in df1.columns[1:]],
                                     '50Y')).set_index([[diff] * len(df1.index)])
            # Creates a dataframe with an index, and the additional column of 50Y.

            if df1['Rates'].values.all() in dt_mjt_names:
            # If the first dataframe has all the dt_mjt_names variables names:
                # ['MoF rates', 'JGB rates', 'TONAR rates']
            # then the values of the Rates column are named after the values
                d['Rates'] = ['MoF Diff', 'JGB Diff', 'TONAR Diff']

            elif df1['Rates'].values.all() in dt_us_names:
            # If the first dataframe has all the dt_mjt_names variables names:
                # ['UST rates', 'SOFR rates']
            # then the values of the Rates column are named after the values
                d['Rates'] = ['UST Diff', 'SOFR Diff']

            return df1.append(df2).append(d).fillna(0.0), diff
            # Appends all the different files on a single dataframe, and fills the missing
            # values as float zeros
                # The missing values come from when we add in an extra column of 50Y
                # Instead of filling it in with a for loop for each of the rows, the 
                # .fillna() function is called, and 0.0 replaces NaN.

        ust_sofr_diff_df, ust_sofr_td = differences_in_rates(self,
                                                             us_rates_df_1, 
                                                             us_rates_df_2)
        # Returns the UST and SOFR rates and their differences in a single dataframe,
        # and also the difference in the number of days, in the: ust_sofr_df and 
        # ust_sofr_td, respectively, containing:
            # UST Rates
            # SOFR Rates
            # the difference in UST rates between time 0 and time 1
            # the difference in SOFR rates between time 0 and time 1

        mof_jgb_tnr_diff_df, mof_jgb_tnr_td = differences_in_rates(self,
                                                                   mjt_rates_df_1, 
                                                                   mjt_rates_df_2)
        # Returns the MoF, JGB, and TONAR rates and their differences in a single dataframe,
        # and also the difference in the number of days, in the: mof_jgb_tnr_df and 
        # mof_jgb_tnr_df, respectively, containing:
            # MoF Rates
            # JGB Rates
            # TONAR Rates
            # the difference in MoF rates between time 0 and time 1
            # the difference in JGB rates between time 0 and time 1
            # the difference in TONAR rates between time 0 and time 1
# ------------------------------------------------------------------------------------------------------ #
# Hedging P&L DataFrames
    # ALM Hedging DataFrames
    # FX Hedgind DataFrames
        
        def hedging_p_and_l(self,
                            differences_dataframe: pd.DataFrame,
                            hedging_dataframe: pd.DataFrame):

            """
            Parameters
            ----------
            differences_dataframe: a dataframe with the differences in rates data
            hedging_dataframe:     a dataframe with the hedging data  
            ----------

            Return
            ------
            h: a dataframe of the DV01 data (depending on whether ALM or FX Hedging
               data is used) and the P&L data.
            ------
            """

            rng = lambda x: np.arange(1, len(x.index) + 1)
            # The range of the data shifted 1 up and accounting for multi-level indexing
            p_and_l = lambda x, y: [[i * j for i, j in zip(x.iloc[-k].values,
                                                           y.iloc[-k, 1:].values)] for k in np.arange(1, 
                                                                                                       len(h)+1)]
            # Profit and Loss double for-loop: x is the treasury rate and y is the DV01
            to_dt = lambda x, y, name: pd.DataFrame({f'{name}': [*[i for i in x]]}, index=y.columns).T
            # Transforms the data into a useable dataframe that can be read together with the rates data
            # x is the resulting list from the p_and_l lambda function
            # y is the DV01 data which we want to compare it to
            # name is the what we want to name the data - this will have to be manually entered

            rs = differences_dataframe
            # Rates Data

            h = hedging_dataframe
            # Reads the Hedging Data

            if h.index.name == 'ALM Hedging':

                pl_swap, pl_jgb, pl_bel = p_and_l(h, rs)[::-1]
                # Returns the P&L Swap, P&L JGB, and P&L BEL data, in that order.


                pl = [pl_swap, pl_jgb, pl_bel]
                # Put all the data inside a list - so we have a list inside of a list

                names = ['P&L BEL', 'P&L JGB', 'P&L Swap']
                # Name all the data fields manually

            elif h.index.name == 'FX Hedging':

                pl_swap, pl_bond = p_and_l(h, rs)
                # Returns the P&L Swap and the P&L bond data, in that order.

                pl = [pl_bond, pl_swap]
                # Put all the data inside a list - so we have a list inside of a list

                names = ['P&L Bond', 'P&L Swap']
                # Name all the data fields manually

            for i, name in zip(pl, names):
                 h = h.append(to_dt(i, h, name))
            # A two line for loop is used here since we need to use the equal sign
            # We simply append our data onto the hedging data
            return h

        alm_hedging_pnl_df = hedging_p_and_l(self,
                                             mof_jgb_tnr_diff_df,
                                             alm_hedge_dataframe)
        # The ALM Hedging Dataframe with P&L, variables listed below:
            # DV01 BEL
            # DV01 JGB
            # DV01 Swaps
            # P&L BEL
            # P&L JGB
            # P&L Swap

        fx_hedging_pnl_df = hedging_p_and_l(self,
                                            ust_sofr_diff_df, 
                                            fx_hedge_dataframe)
        # The FX Hedging Dataframe with P&L, variables listed below:
            # DV01 Bond
            # DV01 Swaps
            # P&L Bond
            # P&L Swaps
# ------------------------------------------------------------------------------------------------------ #
# P&L DataFrames #
    # Bond, Swap, Treasury Change, and Swap Change DataFrame
    # P&L Liabilities, P&L JGB + JPY Swap Portfolio, MoF Change, Treasury Change, Swap Change
            
        def p_and_l(self,
                    fx_hedging_pnl_df: pd.DataFrame, 
                    ust_sofr_diff_df: pd.DataFrame,
                    ust_sofr_td: str,
                    alm_hedging_pnl_df: pd.DataFrame,
                    mof_jgb_tnr_diff_df: pd.DataFrame,
                    mof_jgb_tnr_td: str):

            """
            Parameters
            ----------
            fx_hedging_pnl_df:   the FX Hedging Dataframe with P&L
            ust_sofr_diff_df:    the UST and SOFR rates differences dataframe
            ust_sofr_td:         the UST and SOFR rates time differences
            alm_hedging_pnl_df:  the ALM Hedging Dataframe with P&L
            mof_jgb_tnr_diff_df: the MoF, JGB, and TONAR rates differences dataframe
            mof_jgb_tnr_td:      the MoF, JGB, and TONAR rates time differences
            ----------

            Return
            ------
            bs_port_ts_chg: a dataframe containing:
                                P&L USD Bond Portfolio
                                P&L USD Swap Portfolio
                                Treasury Change
                                Swap Change
            ljj_mjs_chg:    a dataframe containing:
                                P&L Liabilities
                                P&L JGB + JPY Swap Portfolio
                                MoF Change
                                JGB Change
                                Swap Change
            ------
            """

            bs_port_ts_chg_names = ['P&L USD Bond Portfolio',
                                    'P&L USD Swap Portfolio', 
                                    'Treasury Change', 
                                    'Swap Change']
            # Names of the assets we want to look at

            fx_hedge_data = fx_hedging_pnl_df.T
            # Gets the FX Hedge P&L dataframe and transposes it
            ust_sofr_data = ust_sofr_diff_df.T[ust_sofr_td]
            # Gets the UST and SOFR data from the UST and SOFR Differences dataframe.
            # The dataframe is then transposed and locked onto the final columns
            # containing the differences.

            index = ust_sofr_data.index[1:]
            # The index variable to be used when creating the final dataframe - 
            # This exists for consistency purposes.

            pl_usd_bond = - fx_hedge_data['P&L Bond'].values
            # P&L USD Bond Portfolio
            pl_usd_swap = fx_hedge_data['P&L Swap'].values
            # P&L USD Swap Portfolio
            treasury_change = ust_sofr_data.iloc[1:, 0].values
            # Treasury Change
            swap_change = ust_sofr_data.iloc[1:, 1].values
            # Swap Change

            data_list_bsts = [pl_usd_bond, 
                              pl_usd_swap,
                              treasury_change,
                              swap_change]
            # The data are put in an iterable list so that we may use a for loop

            bs_port_ts_chg = pd.DataFrame(*[{
                f'{name}': [i for i in x] for name, x in zip(bs_port_ts_chg_names,
                                                             data_list_bsts)}], index = index).T
            # Adds all the data fields into a single dataframe with respect to their names

            ljj_mjs_chg_names = ['P&L Liabilities', 
                                 'P&L JGB + JPY Swap Portfolio', 
                                 'MoF Change', 
                                 'JGB Change', 
                                 'Swap Change']
            # Names of the assets we want to look at

            alm_hedge_data = alm_hedging_pnl_df.T
            # Gets the ALM Hedge P&L dataframe and transposes it
            mof_jgb_tnr_data = mof_jgb_tnr_diff_df.T[mof_jgb_tnr_td]
            # Gets the MoF, JGB, and TONAR data from the MoF, JGB, and TONAR Differences dataframe.
            # The dataframe is then transposed and locked onto the final columns
            # containing the differences.

            pl_lb = alm_hedge_data['P&L BEL'].values
            # P&L Liabilities
            pl_jgb_jpy = -(alm_hedge_data['P&L JGB'] + alm_hedge_data['P&L Swap'])
            # P&L JGB + JPY Swap Portfolio
            mof_chg = mof_jgb_tnr_data.iloc[1:, 0]
            # MoF Change
            jgb_chg = mof_jgb_tnr_data.iloc[1:, 1]
            # JGB Change
            swp_chg = mof_jgb_tnr_data.iloc[1:, 2]
            # Swap Change

            data_list_ljj_mjs = [pl_lb, 
                                 pl_jgb_jpy, 
                                 mof_chg, 
                                 jgb_chg, 
                                 swp_chg]
            # The data are put in an iterable list so that we may use a for loop

            ljj_mjs_chg = pd.DataFrame(*[{
                f'{name}': [i for i in x] for name, x in zip(ljj_mjs_chg_names,
                                                             data_list_ljj_mjs)}], index = index).T
            # Adds all the data fields into a single dataframe with respect to their names

            return bs_port_ts_chg, ljj_mjs_chg

        bs_ts, ljj_mjs = p_and_l(self,
                                 fx_hedging_pnl_df,
                                 ust_sofr_diff_df,
                                 ust_sofr_td,
                                 alm_hedging_pnl_df,
                                 mof_jgb_tnr_diff_df,
                                 mof_jgb_tnr_td)
# ------------------------------------------------------------------------------------------------------ #
# TE Hedging DataFrame
        
        def te_hedging(self,
                       ust_sofr_diff_df: pd.DataFrame,
                       mof_jgb_tnr_diff_df: pd.DataFrame,
                       alm_hedge_dataframe: pd.DataFrame,
                       fx_hedge_dataframe: pd.DataFrame,
                       ust_sofr_td: str,
                       mof_jgb_tnr_td: str):

            """
            Parameters
            ----------
            ust_sofr_diff_df:     the UST and SOFR Differences dataframe
            mof_jgb_tnr_diff_df:  the MoF, JGB, and TONAR Differences dataframe
            alm_hedge_dataframe:  the ALM Hedge dataframe
            fx_hedge_dataframe:   the FX Hedge dataframe
            ust_sofr_td:          the UST and SOFR time difference, given two 
                                  two different dates
            mof_jgb_tnr_td:       the MoF, JGB, and TONAR time difference, given two 
                                  two different dates
            ----------

            Return
            ------
            te_hedge_alm_df: the ALM Hedging dataframe containing:
                                P&L BEL
                                P&L JGB
                                P&L Swap
            te_hedge_fx_df:  the FX Hedging dataframe containing:
                                P&L Bond
                                P&L Swap
            ------
            """
            index = ust_sofr_diff_df.T[ust_sofr_td].index[1:]

            filtered_us_rates = ust_sofr_diff_df.T[ust_sofr_td]
            # UST and SOFR difference dataframe filtered for only the differences by
            # looking only at the time differences columns
            ust_diff = filtered_us_rates.iloc[1:, 0]
            # UST Differences
            sofr_diff = filtered_us_rates.iloc[1:, 1]
            # SOFR Differences

            filtered_mjt_rates = mof_jgb_tnr_diff_df.T[mof_jgb_tnr_td]
            # MoF, JGB, and TONAR difference dataframe filtered for only the differences by
            # looking only at the time differences columns
            mof_diff = filtered_mjt_rates.iloc[1:, 0]
            # MoF Differences
            jgb_diff = filtered_mjt_rates.iloc[1:, 1]
            # JGB  Differences
            tnr_diff = filtered_mjt_rates.iloc[1:, 2]
            # TONAR Differences

            dv01_fx = fx_hedge_dataframe.T
            # Dataframe with the DV01 FX Hedge data
            dv01_fxbonds = dv01_fx.iloc[0:, 0]
            # DV01 FX Bonds
            dv01_fxswaps = dv01_fx.iloc[0:, 1]
            # DV01 FX Swaps

            dv01_alm = alm_hedge_dataframe.T
            # Dataframe with the DV01 ALM Hedge data
            dv01_almbel = dv01_alm.iloc[0:, 0]
            # DV01 ALM BEL
            dv01_almjgb = dv01_alm.iloc[0:, 1]
            # DV01 ALM JGB
            dv01_almswp = dv01_alm.iloc[0:, 2]
            # DV01 ALM Swaps

            te_hedge_pl_fxbd = ust_diff.values * dv01_fxbonds.values
            # TE Hedge P&L FX Bonds
            te_hedge_pl_fxswp = ust_diff.values * dv01_fxswaps.values
            # TE Hedge P&L FX Swaps

            te_hedge_pl_almbel = mof_diff.values * dv01_almbel.values
            # TE Hedge P&L ALM BEL
            te_hedge_pl_almjgb = jgb_diff.values * dv01_almjgb.values
            # TE Hedge P&L ALM JGB 
            te_hedge_pl_almswp = mof_diff.values * dv01_almswp.values
            # TE Hedge P&L ALM Swaps

            te_alm_names = ['P&L BEL', 'P&L JGB', 'P&L Swap']
            # TE ALM Hedging P&L names
            te_alm_data = [te_hedge_pl_almbel, 
                           te_hedge_pl_almjgb, 
                           te_hedge_pl_almswp]
            # ALM P&L data

            te_hedge_alm_df = pd.DataFrame(*[{
                f'{name}': [i for i in x] for name, x in zip(te_alm_names, 
                                                             te_alm_data)}], index=index).T
            # ALM TE Hedge dataframe
            te_hedge_alm_df.index.name = 'TE ALM Hedging'
            # Set index name to TE ALM Hedging

            te_fx_names = ['P&L Bond', 'P&L Swap']
            # TE FX Hedging P&L names
            te_fx_data = [te_hedge_pl_fxbd, 
                          te_hedge_pl_fxswp]

            te_hedge_fx_df = pd.DataFrame(*[{
                f'{name}': [i for i in x] for name, x in zip(te_fx_names, 
                                                             te_fx_data)}], index=index).T
            te_hedge_fx_df.index.name = 'TE FX Hedging'
            # Set index name to TE FX Hedging

            return te_hedge_alm_df, te_hedge_fx_df

        te_hedge_alm, te_hedge_fx = te_hedging(self,
                                               ust_sofr_diff_df,
                                               mof_jgb_tnr_diff_df,
                                               alm_hedge_dataframe,
                                               fx_hedge_dataframe,
                                               ust_sofr_td,
                                               mof_jgb_tnr_td)

        def swap_spread(self, 
                        ust_sofr_diff_df, 
                        fx_hedge_dataframe):
            """
            Parameters
            ----------
            ust_sofr_diff_df:     the UST and SOFR Differences dataframe
            fx_hedge_dataframe:   the FX Hedge dataframe
            ----------

            Return
            ------
            swap_spread_df: a dataframe returning the swap spread between DV01
                            Bonds and Swaps against the UST difference in rates.
            ------
            """
            ust_diff = ust_sofr_diff_df.iloc[:, 1:][-2:-1].values[0]

            ss_bond, ss_swap = [], []

            for i, j in zip(ust_diff, [*fx_hedging_pnl_df.iloc[:2].T.values]):
                ss_bond.append(i * j[0])
                ss_swap.append(i * j[1])

            swap_spread_df = pd.DataFrame({
                'P&L Bond': ss_bond,
                'P&L Swap': ss_swap
            }, index = ust_sofr_diff_df.columns[1:]).T

            swap_spread_df.index.name = 'Swap Spread'

            return swap_spread_df
        
        swap_spread_df = swap_spread(self, ust_sofr_diff_df, fx_hedge_dataframe)
    
        return (alm_hedge_dataframe, fx_hedge_dataframe, ust_sofr_diff_df,
                    mof_jgb_tnr_diff_df, alm_hedging_pnl_df, fx_hedging_pnl_df,
                    bs_ts, ljj_mjs, te_hedge_alm, te_hedge_fx, swap_spread_df)
# ------------------------------------------------------------------------------------------------------ #
# Curve Attribution Plots #
    
    def curve_attribution_plots(self):

        """
        Parameters
        ----------
        ust_sofr_diff_df: the UST and SOFR Differences dataframe
        ljj_mjs:          the dataframe containing
                                P&L Liabilities
                                P&L JGB + JPY Swap Portfolio
                                MoF Change
                                JGB Change
                                Swap Change
        bs_ts:            the dataframe containing
                                P&L USD Bond Portfolio
                                P&L USD Swap Portfolio 
                                Treasury Change 
                                Swap Change
        ----------

        Return
        ------

        ------
        """
# ------------------------------------------------------------------------------------------------------ #
# UST and SOFR Changes Plot # 
        
        def ust_sofr_change_plot(self, ust_sofr_diff_df):

            """
            Parameters
            ----------
            ust_sofr_diff_df: the UST and SOFR Differences dataframe
            ----------

            Return
            ------
            The UST Change vs. SOFR Changes plot
            ------
            """

            plt.plot(ust_sofr_diff_df.columns[1:-1].values,
                     ust_sofr_diff_df.iloc[-2:-1, 1:-1].values[0], 
                     marker='o',
                     label='UST Diff',
                     color='#05c3dd')
            plt.plot(ust_sofr_diff_df.columns[1:-1].values,
                     ust_sofr_diff_df.iloc[-1, 1:-1].values, 
                     marker='o',
                     label='SOFR Diff',
                     color='#7965b2') # Purple, Accent 5
            plt.title('UST Changes vs. SOFR Changes')
            plt.legend()
            plt.grid()
            plt.savefig('UST Changes vs. SOFR Changes.png')
            plt.close();
            
            return print('A plot analysing UST and SOFR Differences has been saved.')
# ------------------------------------------------------------------------------------------------------ #
# P&L Liabilities + JGB + JPY #
        
        def pl_liabilities_pl_jgb_jpy(self, ljj_mjs):
            
            """
            Parameters
            ----------
            ljj_mjs: the dataframe containing
                        P&L Liabilities
                        P&L JGB + JPY Swap Portfolio
                        MoF Change
                        JGB Change
                        Swap Change
            ----------

            Return
            ------
            The P&L Liabilities vs. the P&L JGB + P&L Swap Portfolio plot
            ------
            """

            plt.plot(ljj_mjs.T['P&L Liabilities'],
                     marker='o', 
                     label='P&L Liabilities')
            plt.plot(ljj_mjs.T['P&L JGB + JPY Swap Portfolio'], 
                     marker='o', 
                     label='P&L JGB + JPY Swap Portfolio')
            plt.legend()
            plt.title('P&L Liabilities vs. P&L JGB + JPY Swap Portfolio')
            plt.grid()
            plt.savefig('P&L Liabilities vs. P&L JGB + JPY Swap Portfolio.png')
            plt.close();
            
            return print('A plot comparing the P&L Liabilities vs. the '  
                         'P&L JGB + P&L Swap Portfolio plot has been saved.')
# ------------------------------------------------------------------------------------------------------ #
# KIKU B/S Rates Attribution Analysis Plot #
            
        def kiku_bs_rates_attribution_analysis(self, ljj_mjs):
            
            """
            Parameters
            ----------
            ljj_mjs: the dataframe containing
                        P&L Liabilities
                        P&L JGB + JPY Swap Portfolio
                        MoF Change
                        JGB Change
                        Swap Change
            ----------

            Return
            ------
            The Kiku — B/S — Rates Attribution Analysis plot
            ------
            """

            bar_width = 0.3
            space = 0.15
            index = np.arange(0, len(ljj_mjs.columns[:-1]))

            fig, ax = plt.subplots(figsize=(16,10))

            ax.bar(index + space, 
                   ljj_mjs.iloc[0, :-1], 
                   bar_width, 
                   color='#2c5697', # Dark Teal, Accent 1
                   label='P&L Liabilities')
            # Bar plot of P&L Liabilities

            ax.bar(index - space, 
                   ljj_mjs.iloc[1, :-1], 
                   bar_width, 
                   color='#00ab84', # Dark Blue, Accent 2
                   label='P&L JGB + JPY Swap Portfolio')
            # Bar plot of P&L JGB + JPY Swap Portfolio

            ax.set_xticks(index, 
                          [f'Tenor {i}' for i in ljj_mjs.columns[:-1]], fontsize=14)
            # Sets the x tickers

            ax.set_ylabel('JPY', 
                          fontsize=16)
            # Names Y axis

            ax.yaxis.get_major_formatter().set_scientific(False)
            # Turns off the scientific notation

            ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y,  p: format(int(y), ',')))
            # Adds a comma for thousands

            ax2 = ax.twinx()
            # Second x axis
            ax2.plot(ljj_mjs.T['MoF Change'][:-1], 
                     marker='o', 
                     label='MoF Change',
                     color='#05c3dd') # Acqua, Accent 3
            # Plots MoF differences

            ax2.plot(ljj_mjs.T['JGB Change'][:-1], 
                     marker='o', 
                     label='JGB Change',
                     color='#4c82a8') # Blue Grey, Accent 4
            # Plots JGB Differences

            ax2.plot(ljj_mjs.T['Swap Change'][:-1], 
                     marker='o', 
                     label='Swap Change',
                     color='#7965b2') # Purple, Accent 5
            # Plots Swap Changes

            ax2.set_ylabel('BPs', 
                          fontsize=16)
            # BPs Y label

            fig.legend(loc='center left',
                       labelspacing=1, 
                       bbox_to_anchor=(0.14, 0.77), fontsize=14)
            # Adds a legend
            ax2.tick_params(labelsize=16)
            # Tick label size

            ax.tick_params(labelsize=16)
            # 
            plt.grid()
            # Adds grid lines
            plt.title('Kiku — B/S — Rates Attribution Analysis', fontsize=22),
            # Gives a graph a title
            plt.savefig('Kiku — BS — Rates Attribution Analysis.png')
            # Saves the figure
            plt.close();
            # Closes the figure
            
            return print('A plot examining the Kiku — B/S — Rates Attribution ' \
                         'Analysis has been saved.')
# ------------------------------------------------------------------------------------------------------ #
# KIKU Bond Portfolio Rates Attribution Analysis Plot #
        
        def kiku_bondptf_rates_attribution_analysis(self, bs_ts):
                        
            """
            Parameters
            ----------
            bs_ts: the dataframe containing
                      P&L USD Bond Portfolio
                      P&L USD Swap Portfolio 
                      Treasury Change 
                      Swap Change
            ----------

            Return
            ------
            The Kiku — Bond Portfolio — Rates Attribution Analysis plot
            ------
            """ 

            bar_width = 0.3
            space = 0.15
            index = np.arange(0, len(bs_ts.columns[:-1]))

            fig, ax = plt.subplots(figsize=(16,10))

            ax.bar(index + space, 
                   bs_ts.iloc[0, :-1], 
                   bar_width, 
                   color='#2c5697', # Dark Teal, Accent 1
                   label='P&L USD Bond Portfolio')
            # P&L USD Bond Portfolio

            ax.bar(index - space, 
                   bs_ts.iloc[1, :-1], 
                   bar_width, 
                   color='#00ab84', # Dark Blue, Accent 2
                   label='P&L USD Swap Portfolio')
            # P&L USD Swap Portfolio

            ax.set_xticks(index, 
                          [f'Tenor {i}' for i in bs_ts.columns[:-1]], fontsize=14)
            # Sets the x tickers

            ax.set_ylabel('USD', 
                          fontsize=16)
            # Names Y axis

            ax.yaxis.get_major_formatter().set_scientific(False)
            # Turns off the scientific notation

            ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y,  p: format(int(y), ',')))
            # Adds a comma for thousands

            ax2 = ax.twinx()
            # Second x axis
            ax2.plot(bs_ts.T['Treasury Change'][:-1], 
                     marker='o', 
                     label='Treasury Change',
                     color='#05c3dd') # Acqua, Accent 3
            # Plots MoF differences

            ax2.plot(bs_ts.T['Swap Change'][:-1], 
                     marker='o', 
                     label='Swap Change',
                     color='#7965b2') # Purple, Accent 5
            # Plots Swap Changes

            ax2.set_ylabel('BPs', 
                          fontsize=16)
            # BPs Y label

            ax2.tick_params(labelsize=16)
            # Y ticks for BPs 
            ax.tick_params(labelsize=16)
            # Y ticks for USD
            fig.legend(loc='center',
                       labelspacing=1, 
                       bbox_to_anchor=(0.4, 0.7), fontsize=14)
            # Adds a legend

            plt.grid()
            plt.title('Kiku — Bond Portfolio — Rates Attribution Analysis', fontsize=22),
            plt.savefig('Kiku — Bond Portfolio — Rates Attribution Analysis.png')
            plt.close();
            
            return print('A plot examining the Kiku — Bond Portfolio — Rates Attribution ' \
                         'Analysis has been saved.')
            
        return (ust_sofr_change_plot(self, ust_sofr_diff_df),
                pl_liabilities_pl_jgb_jpy(self, ljj_mjs), 
                kiku_bs_rates_attribution_analysis(self, ljj_mjs),
                kiku_bondptf_rates_attribution_analysis(self, bs_ts))
# ------------------------------------------------------------------------------------------------------ #
# Convexity Attribution Sheet #
    
    def convexity_attribution(self,
                              ust_sofr_diff_df: pd.DataFrame,
                              mof_jgb_tnr_diff_df: pd.DataFrame,
                              fx_hedge_dataframe: pd.DataFrame,
                              alm_hedge_dataframe: pd.DataFrame,
                              convexities: list):
        """
        Parameters
        ----------
        ust_sofr_diff_df:    UST and SOFR Differences dataframe
        mof_jgb_tnr_diff_df: MoF, JGB, and TONER Differences dataframe
        fx_hedge_dataframe:  FX Hedge Dataframe
        alm_hedge_dataframe: ALM Hedge Dataframe
        convexities:         a list of convexities, with respect to the following assets, 
                             in the following order
                                - USD Bond Portfolio
                                - USD Payers
                                - JPY IRS Ptf
                                - JGB Ptf
                                - JPY BEL
        ----------

        Return
        ------
        df: a dataframe that contains the convexity, rate change, and the convexity P&L
            for each of the assets in the above order

            NOTE - USDJPY Forward is not included as it is currently all zero.
        ------
        """

        ust_diff = ust_sofr_diff_df.iloc[-2, 1:]
        # UST Differences
        sofr_diff = ust_sofr_diff_df.iloc[-1, 1:]
        # SOFR Differences

        mof_diff = mof_jgb_tnr_diff_df.iloc[-3, 1:]
        # MoF Differences
        jgb_diff = mof_jgb_tnr_diff_df.iloc[-2, 1:]
        # JGB Differences
        tnr_diff = mof_jgb_tnr_diff_df.iloc[-1, 1:]
        # TONAR Differences

        rates = [ust_diff, sofr_diff, tnr_diff, jgb_diff, mof_diff]
        # A list of the differences of the rates

        dv01s_fxbonds = fx_hedge_dataframe.iloc[0]
        # DV01 FX Bonds
        dv01s_fxswaps = fx_hedge_dataframe.iloc[1]
        # DV01 FX Swaps

        dv01_almbel = alm_hedge_dataframe.iloc[0]
        # DV01 ALM BEL
        dv01_almjgb = alm_hedge_dataframe.iloc[1]
        # DV01 ALM JGB
        dv01_almswp = alm_hedge_dataframe.iloc[2]
        # DV01 ALM Swaps

        dv01s = [dv01s_fxbonds, dv01s_fxswaps, dv01_almswp, dv01_almjgb, dv01_almbel]
        # A List of DV01s including both FX Hedges and ALM Hedges
# ------------------------------------------------------------------------------------------------------ #
# Convexity P&L DataFrame
        
        def convex_pl_calc(rate, dv01, convexity, name):

            """
            Parameters
            ----------
            rate:      the rate series
            dv01:      the dv01 series
            convexity: a convexity value - need to see where this comes from
            name:      the name of the asset
            ----------

            Return
            ------

            ------
            """

            rate_change = sum([i*j for i, j in zip(rate, dv01)])/sum(dv01)
            # Calculates the rate change
            convex_pl = convexity * 0.5 * rate_change**2
            # Calculates the convexity P&L
            df = pd.DataFrame({
                f'{name}': [convexity, rate_change, convex_pl]
            }, index=['Convexity', 'Rate Change', 'Convexity P&L'])
            # Puts all the data in a dataframe

            return df


        names = ['USD Bond portfolio', 
                 'USD Payers', 
                 'JPY IRS Portfolio', 
                 'JGB Portfolio',
                 'JPY BEL']
        
        return pd.concat([convex_pl_calc(rate, dv01, 
                                     convexity, name) for (rate, dv01, 
                                                           convexity, name) in zip(rates, dv01s,
                                                                                  convexities, names)], 
                     axis=1)
# ------------------------------------------------------------------------------------------------------ #
# Credit Spread Attribution Sheet # 
    
    def credit_spread_attribution(self):

        """
        Parameters
        ----------
        filenames: a list of the file names required
        dates:     a list of the dates relative to the file names
        ----------

        Return
        ------
        df1: a dataframe of USDJPY CURNCY, MtM USD Bond Portfolio, and MtM JPY
        s1:  a series of FX P&L Bonds
        s2:  a series of FX Relative Performance
        ------
        """

        oas_duration, op_as, oas, csa_mtm = [], [], [], []

        for filename, date, i in zip(self.filenames, self.dates, np.arange(0, len(dates))):

            d = self.data_sheet(filename, date)
            # Obtain data

            d = d[d['Currency'] == 'USD']

            md = d['Modified Duration'].replace('---', 0)
            # Modified Duration

            mv = d['Market Value']
            # Market Value

            op_as = d['Option-Adjusted Spread'].replace('---', 0)
            # Option-Adjusted Spread - Data Sheet

            mtm = sum(d['Accrued Balance']) + sum(mv)

            oas_duration.append(self.sumprod_sum(md, mv))
            # OAS Duration

            oas.append(self.sumprod_sum(op_as, mv))
            # Option Adjusted Spread — Credit Spread Attribution

            csa_mtm.append(mtm)
            # Credit Spread Attribution MtM

        c_pl = ((sum(csa_mtm) / len(csa_mtm)) * \
                (sum(oas_duration) / len(oas_duration)) * \
                (oas[-1] - oas[0])) / -10000
        # Credit P&L

        df1 = pd.DataFrame({
            'OAS Duration': oas_duration,
            'OAS': oas,
            'MtM': csa_mtm
        }, index = dates)

        return df1, c_pl
# ------------------------------------------------------------------------------------------------------ #
# Carry Attribution Sheet #
    
    def carry_attribution(self, mof_jgb_tnr_diff_df):

        """
        Parameters
        ----------
        filenames:       the name of the file - typically 'TD_LFI_Template_DL-Kiku-CA-JPM-FI (294204).xlsx'
        dates:           the time at wish the file should be examined
        ----------

        Return
        ------
        df1: a dataframe containing the Date, YTM, OAS, RFR, and MtM
        ------
        """
# ------------------------------------------------------------------------------------------------------ #
# Time Fraction, Risk-Free Rate, Credit Spread, Risk-Free Rate Carry, Credit Spread Carry,
# and FX Basis Carry
        
        def tf_rfr_cs_rfrc_csc_fxbc(self,
                                    dates, 
                                    rfr, 
                                    oas, 
                                    mtm, 
                                    basis=[0.0], 
                                    notional=[0.0], 
                                    mtm_cs=0.0, 
                                    swap=True):

            """
            Parameters
            ----------
            dates:    the dates of the two file names
            rfr:      the risk free rate
            oas:      the options adjusted spread amount
            mtm:      the mark to market amount
            basis:    takes the basis amount
            notional: takes the notional amount obtained from the J. Dass report
            mtm_cs:   the MtM value obtained from the credit spread attribution
            ----------

            Return
            ------
            df: a dataframe containing the time fraction, date, credit spread, risk free rate carry,
                credit spread carry, and the fx basis carry
            ------
            """

            dt_fmt = '%m/%d/%Y'
            # Date Format
            dt_strp = lambda x: dt.datetime.strptime(x, dt_fmt)
            # Date Striptime
            tf = (dt_strp(dates[1]) - dt_strp(dates[0])) / dt.timedelta(days=365)
            # Time Fraction

            rfr_bp_diff = (pd.Series(rfr).diff() / 10000).dropna()
            # Risk Free Rate Difference in bps

            cs_bp_diff = (pd.Series(oas).diff() / 10000).dropna()
            # Credit Spread Difference in bps

            if swap == True:

                df = pd.DataFrame({
                    'Time Fraction': tf,
                    'RFR': rfr_bp_diff.values,
                    'Credit Spread': cs_bp_diff.values,
                    'RFR Carry': tf * (sum(rfr) / len(rfr)) * (sum(notional) / 2) / 10000,
                    'Credit Spread Carry': (tf * sum(oas) / len(oas) * (sum(notional) / 2)) / 10000,
                    'FX Basis Carry': (1 - ((1-(basis[0]/100))**(tf))) * mtm_cs
                })

                return df

            else:

                df = pd.DataFrame({
                    'Time Fraction': tf,
                    'RFR': rfr_bp_diff.values,
                    'Credit Spread': cs_bp_diff.values,
                    'RFR Carry': (tf * (sum(rfr) / 2) * (sum(mtm) / 2)) / 10000,
                    'Credit Spread Carry': (tf * sum(oas) / len(oas) * (sum(mtm) / 2)) / 10000,
                    'FX Basis Carry': (tf * (sum(basis) / 2) * (sum(notional) / 2)) / 10000
                })

                return df        
# ------------------------------------------------------------------------------------------------------ #
# USD Bonds Account #
            
        def usd_bond(filenames, dates):

            """
            Parameters
            ----------
            filenames:    names of the two files of data to be extracted
            dates:        the dates of the two file names
            fx_filenames: the name of the fx file in list format
            ----------

            Return
            ------
            df: a dataframe containing the yield to maturity, option adjusted spread, risk-free rate, 
                and the mark to market of the USD Bond Portfolio
            ------
            """

            oas = self.credit_spread_attribution()[0]['OAS']
            # The Option Adjusted Spread is already given as a list therefore there is no 
            # need to loop through.

            ytm, mtm, rfr = [], [], []

            for f, dat, i in zip(self.filenames, self.dates, np.arange(0, len(oas))):
                d = self.data_sheet(f, dat)
                d = d[d['Currency'] == 'USD']
                mtm.append(sum([i / self.bbg_value for i in d['Base Market Value + Accrued']]))
                # Mark to Market 

                ytm.append(self.sumprod_sum(d['Yield to Maturity'], d['Market Value']) * 100)
                # Yield to Maturity

                rfr.append(ytm[i] - oas[i])
                # Risk free rate

            usd_bond_portfolio = pd.DataFrame({
                'YTM': ytm,
                'OAS': oas,
                'RFR': rfr,
                'MtM': mtm}, index = self.dates)        
            usd_bond_portfolio.index.name = 'USD Bond Portfolio'

            df = tf_rfr_cs_rfrc_csc_fxbc(self, self.dates, rfr, oas, mtm, swap=False)
            # The above function returns the dataframe containing the time fraction, date, credit spread, 
            # risk free rate carry, credit spread carry, and the fx basis carry

            return usd_bond_portfolio, df

        trcrcf_df = pd.DataFrame()

        usd_bond_carry_df, df = usd_bond(self.filenames, self.dates)

        trcrcf_df = trcrcf_df.append(df)
# ------------------------------------------------------------------------------------------------------ #
# USD Payer Account #
        
        def usd_payer(dates, murray=False):

            usd_payer_nprb = pd.DataFrame({
                'Notional': [815740000.0, 883800030.0],
                'Payer Rate': [-470.0, -463.0],
                'Receiver Rate': [189.34, 179.39],
                'Basis': [0, 0]}, index = self.dates)
            notional = usd_payer_nprb['Notional'].values

            ytm = [(x + y) for x, y in zip(usd_payer_nprb['Payer Rate'],
                                           usd_payer_nprb['Receiver Rate'])]
            oas = [0, 0]
            rfr = [(x - y) for x, y in zip(ytm, oas)]
            mtm = [54173156.00, 60039673.0]

            usd_payer_carry_df = pd.DataFrame({
                'YTM': ytm,
                'OAS': oas,
                'RFR': rfr,
                'MtM': mtm}, index = self.dates)
            usd_payer_carry_df.index.name = 'USD Payer'

            df = tf_rfr_cs_rfrc_csc_fxbc(self, 
                                         self.dates, 
                                         rfr, 
                                         oas, 
                                         mtm, 
                                         notional=notional)
            # The above function returns the dataframe containing the time fraction, date, credit spread, 
            # risk free rate carry, credit spread carry, and the fx basis carry

            return usd_payer_carry_df, df

        usd_payer_carry_df, df = usd_payer(self.dates)

        trcrcf_df = trcrcf_df.append(df)
# ------------------------------------------------------------------------------------------------------ #
# JPY Receiver ALM Account #

        def jpy_receiver_alm(dates, murray=False):

            jpy_receiver_alm_nprb = pd.DataFrame({
                'Notional': [29831000000.00, 33921000000.0],
                'Payer Rate': [0, 0],
                'Receiver Rate': [85.0, 85.0],
                'Basis': [0, 0]}, index = dates)
            notional = jpy_receiver_alm_nprb['Notional'].values

            oas = [0, 0]
            ytm = [(x + y) for x, y in zip(jpy_receiver_alm_nprb['Payer Rate'],
                                           jpy_receiver_alm_nprb['Receiver Rate'])]
            rfr = [(x - y) for x, y in zip(ytm, oas)]
            mtm = [-413612944.00, -1660381991.0]


            jpy_receiver_alm_carry_df = pd.DataFrame({
                'YTM': ytm,
                'OAS': oas,
                'RFR': rfr,
                'MtM': mtm}, index = dates)
            jpy_receiver_alm_carry_df.index.name = 'JPY Receiver ALM'

            df = tf_rfr_cs_rfrc_csc_fxbc(self,
                                         self.dates, 
                                         rfr, 
                                         oas, 
                                         mtm, 
                                         notional=notional)
            # The above function returns the dataframe containing the time fraction, date, credit spread, 
            # risk free rate carry, credit spread carry, and the fx basis carry

            return jpy_receiver_alm_carry_df, df

        jpy_receiver_alm_carry_df, df = jpy_receiver_alm(self.dates)

        trcrcf_df = trcrcf_df.append(df)
# ------------------------------------------------------------------------------------------------------ #
# JPY Receiver FX Account #
        
        def jpy_receiver_fx(dates, murray=False):

            jpy_receive_fx_nprb = pd.DataFrame({
                'Notional': [88652350000.0, 87497350000],
                'Payer Rate': [0, 0],
                'Receiver Rate': [77.0, 77.0],
                'Basis': [0, 0]}, index = self.dates)
            notional = jpy_receive_fx_nprb['Notional'].values

            oas = [0, 0]
            ytm = [(x + y) for x, y in zip(jpy_receive_fx_nprb['Payer Rate'],
                                           jpy_receive_fx_nprb['Receiver Rate'])]
            rfr = [(x - y) for x, y in zip(ytm, oas)]
            mtm = [-1355095968.0, -4877786812.0]

            jpy_receiver_fx_carry_df = pd.DataFrame({
                'YTM': ytm,
                'OAS': oas,
                'RFR': rfr,
                'MtM': mtm}, index = self.dates)
            jpy_receiver_fx_carry_df.index.name = 'JPY Receiver FX'

            df = tf_rfr_cs_rfrc_csc_fxbc(self, 
                                         self.dates, 
                                         rfr, 
                                         oas, 
                                         mtm, 
                                         notional=notional)
            # The above function returns the dataframe containing the time fraction, date, credit spread, 
            # risk free rate carry, credit spread carry, and the fx basis carry

            return jpy_receiver_fx_carry_df, df

        jpy_receiver_fx_carry_df, df = jpy_receiver_fx(self.dates)

        trcrcf_df = trcrcf_df.append(df)
# ------------------------------------------------------------------------------------------------------ #
# JPY Receiver IRS Account # 

        jpy_irs = [i + j for i, j in trcrcf_df.T.iloc[:, 2:].values]
        trcrcf_df = trcrcf_df.append(pd.DataFrame(jpy_irs, index = ['Time Fraction',
                                                                    'RFR',
                                                                    'Credit Spread',
                                                                    'RFR Carry',
                                                                    'Credit Spread Carry',
                                                                    'FX Basis Carry']).T)
# ------------------------------------------------------------------------------------------------------ #
# JGB Account #

        def jgb(filenames, dates):

            jgb_ytm, jgb_oas, jgb_mtm = [], [], []

            for i in filenames:
                d = pd.read_excel(i, engine='openpyxl')
                d.columns = [*[x for x in d.iloc[5, :]]]
                d = d.iloc[6:-2, :]
                d = d[d['Currency'] == 'JPY'][1:]
                d_ytm = self.sumprod_sum(d['Yield to Maturity'], d['Market Value']) * 100
                d_oas = self.sumprod_sum(d['Option-Adjusted Spread'], d['Market Value'])
                d_mtm = sum(d['Market Value'])
                jgb_ytm.append(d_ytm)
                jgb_oas.append(d_oas)
                jgb_mtm.append(d_mtm)

            jgb_rfr = [(x - y) for x, y in zip(jgb_ytm, jgb_oas)]

            jgb_carry_df = pd.DataFrame({
                'YTM': jgb_ytm,
                'OAS': jgb_oas,
                'RFR': jgb_rfr,
                'MtM': jgb_mtm}, index = self.dates)
            jgb_carry_df.index.name = 'JGB'

            df = tf_rfr_cs_rfrc_csc_fxbc(self,
                                         self.dates, 
                                         jgb_rfr, 
                                         jgb_oas, 
                                         jgb_mtm, 
                                         swap=False)
            # The above function returns the dataframe containing the time fraction, date, credit spread, 
            # risk free rate carry, credit spread carry, and the fx basis carry

            return jgb_carry_df, df

        jgb_carry_df, df = jgb(self.filenames, self.dates)

        trcrcf_df = trcrcf_df.append(df)
# ------------------------------------------------------------------------------------------------------ #
# Forward Account #

        def forward(filenames, dates, fx_filenames):
            # WAITING ON DASS
            notional = [102256002142.53, 113064343517.07]
            payer = [0, 0]
            receiver = [0, 0]
            basis = [-25.0, -17.25]

            forward_nprb = pd.DataFrame({
                'Notional': notional,
                'Payer Rate': payer,
                'Receiver Rate': receiver,
                'Basis': basis}, index = self.dates)

            ytm = [0, 0]
            oas = [0, 0]
            rfr = [0, 0]
            mtm = [0, 0]

            fx_spot_usdjpy_mtm_usdjpy = self.fx_spot_sheet()[0]

            mtm_cs = fx_spot_usdjpy_mtm_usdjpy['MtM JPY'][-1]

            forward_carry_df = pd.DataFrame({
                'YTM': ytm,
                'OAS': oas,
                'RFR': rfr,
                'MtM': mtm}, index = self.dates)
            forward_carry_df.index.name = 'Forward'

            df = tf_rfr_cs_rfrc_csc_fxbc(self,
                                         self.dates, 
                                         rfr, 
                                         oas, 
                                         mtm, 
                                         basis, 
                                         notional, 
                                         mtm_cs, 
                                         swap=False)
            # The above function returns the dataframe containing the time fraction, date, credit spread, 
            # risk free rate carry, credit spread carry, and the fx basis carry

            return forward_carry_df, df

        forward_carry_df, df = forward(self.filenames, self.dates, self.fx_filenames)

        trcrcf_df = trcrcf_df.append(df)
# ------------------------------------------------------------------------------------------------------ #
# BEL Account #
    
        def bel(self, dates, mof_jgb_tnr_diff_df):

            b_ytm = lambda self, x: mof_jgb_tnr_diff_df.loc[mof_jgb_tnr_diff_df.index.unique()[x]]['15Y'][0]
            
            

            bel_ytm = [b_ytm(self, 0) * 100, b_ytm(self, 1) * 100]
            oas = [0, 0]
            rfr = [(x - y) for x, y in zip(bel_ytm, oas)]
            mtm = [163285742125.0, 156373353426.0]

            bel_carry_df = pd.DataFrame({
                'YTM': bel_ytm,
                'OAS': oas,
                'RFR': rfr,
                'MtM': mtm}, index = self.dates)
            bel_carry_df.index.name = 'BEL'

            df = tf_rfr_cs_rfrc_csc_fxbc(self,
                                         self.dates, 
                                         rfr, 
                                         oas, 
                                         mtm, 
                                         swap=False)

            return bel_carry_df, df

        bel_carry_df, df = bel(self, self.dates, mof_jgb_tnr_diff_df)

        trcrcf_df = trcrcf_df.append(df)
        trcrcf_df.index = ['USD Bond Portfolio', 'USD Payer', 
                           'JPY Receiver ALM', 'JPY Receiver FX', 'JPY Receiver IRS',
                           'JGB', 'Forward', 'BEL']

        return (usd_bond_carry_df, usd_payer_carry_df, jpy_receiver_alm_carry_df, jpy_receiver_fx_carry_df,
                 jgb_carry_df, forward_carry_df, bel_carry_df, trcrcf_df)

#########################################################################################################

class reporting:
    
    def __init__(self, filenames, dates, fx_filenames, mjt_rates, us_rates, alm_hedging, fx_hedging):
        self.filenames = filenames
        self.dates = dates
        self.fx_filenames = fx_filenames
        self.mjt_rates = mjt_rates
        self.us_rates = us_rates
        self.alm_hedging = alm_hedging
        self.fx_hedging = fx_hedging
    
# ------------------------------------------------------------------------------------------------------ #
# Millions Notation
    mm = 10**6
# ------------------------------------------------------------------------------------------------------ #
# The USDJPY CURNCY Value
    bbg_value = 131.121
# ------------------------------------------------------------------------------------------------------ #
# SUMPRODUCT / SUM
    sumprod_sum = lambda self, x, y: sum([float(i) * \
                                          float(j) for i, j in zip(x.replace('---', '0'),
                                                                   y.replace('---', '0'))]) / sum(y)
# ------------------------------------------------------------------------------------------------------ #
# Attribution Class

    att = attribution(filenames, 
                      dates, 
                      fx_filenames, 
                      mjt_rates, 
                      us_rates, 
                      alm_hedging, 
                      fx_hedging)
# ------------------------------------------------------------------------------------------------------ #
# FX Spot DataFrames

    (fx_spot_usdjpy_mtm_usdjpy,
     fx_rp_plbond_fwd) = att.fx_spot_sheet()
# ------------------------------------------------------------------------------------------------------ #
# Curve Attribution DataFrames
    
    (alm_hedge_dataframe, 
     fx_hedge_dataframe, 
     ust_sofr_diff_df,
     mof_jgb_tnr_diff_df,
     alm_hedging_pnl_df,
     fx_hedging_pnl_df,
     bs_ts,
     ljj_mjs,
     te_hedge_alm,
     te_hedge_fx,
     swap_spread_df) = att.curve_attribution()
# ------------------------------------------------------------------------------------------------------ #
# Convexities DataFrames
    
    usd_bond_convex = 751.0
    usd_payer_convex = -1487.0
    jpy_irs_ptf_convex = 258453.0 + 86173.0
    jgb_ptf_convex = 50004.0
    usdjpy_fwd_convex = 0
    jpy_bel_convex = 234470.0

    convexities = [usd_bond_convex,
                   usd_payer_convex,
                   jpy_irs_ptf_convex,
                   jgb_ptf_convex,
                   jpy_bel_convex]
    
    convexities_df = att.convexity_attribution(ust_sofr_diff_df,
                                             mof_jgb_tnr_diff_df,
                                             fx_hedge_dataframe,
                                             alm_hedge_dataframe,
                                             convexities)
    
    convexities_df['USDJPY Fwd'] = [0, 0, '---']
    USDJPY_FWD_col = convexities_df.pop('USDJPY Fwd')
    convexities_df.insert(4, 'USDJPY Fwd', USDJPY_FWD_col)
    del USDJPY_FWD_col
# ------------------------------------------------------------------------------------------------------ #
# Credit Spread Attribution DataFrames #
    
    credit_spread_df, credit_spread_pnl = att.credit_spread_attribution()
# ------------------------------------------------------------------------------------------------------ #
# Carry Attribution DataFrames
    
    (usd_bond_carry_df, usd_payer_carry_df,
     jpy_receiver_alm_carry_df, jpy_receiver_fx_carry_df,
     jgb_carry_df, forward_carry_df,
     bel_carry_df,
     trcrcf_df) = att.carry_attribution(mof_jgb_tnr_diff_df)
# ------------------------------------------------------------------------------------------------------ #
# Totals Sheet #

    def total_sheet(self):
        mm = self.mm
        bbg_value = self.bbg_value
# ------------------------------------------------------------------------------------------------------ #
# Formatting

        wb = xlsxwriter.Workbook('20230109 Performance Attribution.xlsx')
        ws = wb.add_worksheet('Total')
        fmt_nmb = wb.add_format({'num_format': '#,##0.0; (#,##0.0)'})
        fmt_pct = wb.add_format({'num_format': '#,##0.0%; (#,##0.0%)'})

        ws.set_column('A:I', None, fmt_nmb)
        ws.set_column('J:J', None, fmt_pct)

        ws.write('B2', 'Date:')
        ws.write('C2', f'{dates[0]} - {dates[-1]}')

        ws.write('B3', 'USDJPY CURNCY')
        ws.write('C3', float(f'{bbg_value}'))
    
    
        titles_usd = ['Return Type', 
                      'USD Bond Portfolio [mUSD]', 
                      'USD Payer IRS Portfolio [mUSD]', 
                      'Net USD Exposure [mUSD]']
        
        titles_jpy = ['Return Type',
                      'JPY Receiver IRS Portfolio [mJPY]',
                      'JGB Portfolio [mJPY]',
                      'USDJPY Forward [mJPY]',
                      'JPY BEL [mJPY]',
                      'Net JPY Exposure [mJPY]', 
                      'Net Exposure [mUSD]']
        
        titles_net = ['Return Type', 
                      'Net P&L B/S of Kiku [mJPY]', 
                      'Net P&L B/S of Kiku [mUSD]',
                      '% Risk Exposure']

        l = ['Total Carry',
             'Rates',
             'Credit Spread',
             'FX Basis',
             'Total RFR Curve',
             '1Y', '2Y', '3Y', '5Y', '7Y', '10Y', '15Y', '20Y', '25Y', '30Y', '40Y', '50Y',
             'Total Convexity',
             'Credit Spread',
             'FX Spot',
             'Trading Costs',
             'Residual',
             'True Total (excl. FX)']

        [ws.write(f'{i}5', f'{x}') for i, x in zip(string.ascii_uppercase[1:5], titles_usd)]
        [ws.write(f'B{i}', f'{x}') for i, x in zip(np.arange(6, 31), l)]
        [ws.write(f'{i}30', f'{x}') for i, x in zip(string.ascii_uppercase[1:8], titles_jpy)]
        [ws.write(f'B{i}', f'{x}') for i, x in zip(np.arange(31, 54), l)]
        [ws.write(f'{i}5', f'{x}') for i, x in zip(string.ascii_uppercase[6:10], titles_net)]
        [ws.write(f'G{i}', f'{x}') for i, x in zip(np.arange(6, 31), l)]
# ------------------------------------------------------------------------------------------------------ #
# USD Bond Portfolio [mUSD] (ubp)

        ubp_rates_carry = self.trcrcf_df['RFR Carry'][0] / mm
        # Rates Carry
        ubp_credit_spread_carry = self.trcrcf_df['Credit Spread Carry'][0] / mm
        # Credit Spread Carry
        ubp_fx_basis_carry = self.trcrcf_df['FX Basis Carry'][0] / mm
        # FX Basis Carry
        ubp_total_carry = sum([ubp_rates_carry, 
                               ubp_credit_spread_carry, 
                               ubp_fx_basis_carry])
        # Total Carry
        ubp_rfr_curve = fx_hedging_pnl_df.T['P&L Bond'] / mm
        # Risk-Free Rate Curve 1Y-50Y
        ubp_total_rfr_curve = sum(ubp_rfr_curve)
        # Total RFR Curve
        ubp_convexity = convexities_df.T['Convexity P&L'][0] / mm
        # Total Convexity
        ubp_credit_spread = credit_spread_pnl / mm
        # Credit Spread
        ubp_fx_spot = self.fx_rp_plbond_fwd['FX P&L Bond'][-1] / mm
        # FX Spot
        ubp_trading_cost = 0.0
        # Trading Costs
        
        df1 = att.data_sheet(filenames[1], dates[1])
        df1 = df1[df1['Currency'] == 'USD']
        df2 = att.data_sheet(filenames[0], dates[0])
        df2 = df2[df2['Currency'] == 'USD']
        ubp_true_total = (sum(df1['Unrealised P&L']) - sum(df2['Unrealised P&L'])) / mm
        # True Total
        
        ubp_residual = (ubp_true_total - 
                        ubp_trading_cost - 
                        ubp_credit_spread - 
                        ubp_convexity - 
                        ubp_total_rfr_curve -
                        ubp_total_carry)
        # Residual
        
        ubp_column = [ubp_total_carry,
                      ubp_rates_carry,
                      ubp_credit_spread_carry,
                      ubp_fx_basis_carry,
                      ubp_total_rfr_curve,
                      *[i for i in ubp_rfr_curve],
                      ubp_convexity,
                      ubp_credit_spread,
                      ubp_fx_spot,
                      ubp_trading_cost,
                      ubp_residual,
                      ubp_true_total]
        
        [ws.write(f'C{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(6, 31), ubp_column)]
# ------------------------------------------------------------------------------------------------------ #
# USD Payer IRS Portfolio [mUSD] (upip)

        upip_rates_carry = self.trcrcf_df['RFR Carry'][1] / mm
        # Rates Carry
        upip_credit_spread_carry = self.trcrcf_df['Credit Spread Carry'][1] / mm
        # Credit Spread Carry
        upip_fx_basis_carry = self.trcrcf_df['FX Basis Carry'][1] / mm
        # FX Basis Carry
        upip_total_carry = sum([upip_rates_carry, 
                               upip_credit_spread_carry, 
                               upip_fx_basis_carry])
        # Total Carry
        upip_rfr_curve = fx_hedging_pnl_df.T['P&L Swap'] / mm
        # Risk-Free Rate Curve 1Y-50Y
        upip_total_rfr_curve = sum(upip_rfr_curve)
        # Total RFR Curve
        upip_convexity = convexities_df.T['Convexity P&L'][1] / mm
        # Total Convexity
        upip_credit_spread = 0.0
        # Credit Spread
        upip_fx_spot = 0.0
        # FX Spot
        upip_trading_cost = 0.0
        # Trading Costs
        
        upip_true_total = (usd_payer_carry_df['MtM'][-1] - usd_payer_carry_df['MtM'][0]) / mm
        # True Total
        
        upip_residual = (upip_true_total - 
                        upip_trading_cost - 
                        upip_credit_spread - 
                        upip_convexity - 
                        upip_total_rfr_curve -
                        upip_total_carry)
        # Residual
        
        upip_column = [upip_total_carry,
                      upip_rates_carry,
                      upip_credit_spread_carry,
                      upip_fx_basis_carry,
                      upip_total_rfr_curve,
                      *[i for i in upip_rfr_curve],
                      upip_convexity,
                      upip_credit_spread,
                      upip_fx_spot,
                      upip_trading_cost,
                      upip_residual,
                      upip_true_total]
        
        [ws.write(f'D{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(6, 31), upip_column)]
# ------------------------------------------------------------------------------------------------------ #
# Net USD Exposure [mUSD] 

        net_usd_exp_column = [x + y for x, y in zip(ubp_column, upip_column)]
        [ws.write(f'E{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(6, 31), net_usd_exp_column)]
# ------------------------------------------------------------------------------------------------------ #
# JPY Receiver IRS Portfolio [mJPY] (jri)

        jri_rates_carry = self.trcrcf_df['RFR Carry'][4] / mm
        # Rates Carry
        jri_credit_spread_carry = self.trcrcf_df['Credit Spread Carry'][4] / mm
        # Credit Spread Carry
        jri_fx_basis_carry = self.trcrcf_df['FX Basis Carry'][4] / mm
        # FX Basis Carry
        jri_total_carry = sum([jri_rates_carry, 
                               jri_credit_spread_carry, 
                               jri_fx_basis_carry])
        # Total Carry
        jri_rfr_curve = [i / mm for i in [*alm_hedging_pnl_df.T['P&L Swap'].values]]
        # Risk-Free Rate Curve 1Y-50Y
        jri_total_rfr_curve = sum(jri_rfr_curve)
        # Total RFR Curve
        jri_convexity = convexities_df.T['Convexity P&L'][2] / mm
        # Total Convexity
        jri_credit_spread = 0.0
        # Credit Spread
        jri_fx_spot = 0.0
        # FX Spot
        jri_trading_cost = 0.0
        # Trading Costs
        
        jri_true_total = (jpy_receiver_alm_carry_df['MtM'][-1] +
                          jpy_receiver_fx_carry_df['MtM'][-1] - 
                         (jpy_receiver_alm_carry_df['MtM'][0] + 
                          jpy_receiver_fx_carry_df['MtM'][0])) / mm
        # True Total
        
        jri_residual = (jri_true_total - 
                        jri_trading_cost - 
                        jri_credit_spread - 
                        jri_convexity - 
                        jri_total_rfr_curve -
                        jri_fx_spot -
                        jri_total_carry)
        # Residual
        
        jri_column = [jri_total_carry,
                      jri_rates_carry,
                      jri_credit_spread_carry,
                      jri_fx_basis_carry,
                      jri_total_rfr_curve,
                      *[i for i in jri_rfr_curve],
                      jri_convexity,
                      jri_credit_spread,
                      jri_fx_spot,
                      jri_trading_cost,
                      jri_residual,
                      jri_true_total]
        
        [ws.write(f'C{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(31, 54), jri_column)]
# ------------------------------------------------------------------------------------------------------ #
# JGB Portfolio [mJPY] (jgb)

        jgb_rates_carry = self.trcrcf_df['RFR Carry'][5] / mm
        # Rates Carry
        jgb_credit_spread_carry = self.trcrcf_df['Credit Spread Carry'][5] / mm
        # Credit Spread Carry
        jgb_fx_basis_carry = self.trcrcf_df['FX Basis Carry'][5] / mm
        # FX Basis Carry
        jgb_total_carry = sum([jgb_rates_carry, 
                               jgb_credit_spread_carry, 
                               jgb_fx_basis_carry])
        # Total Carry
        jgb_rfr_curve = [i / mm for i in [*alm_hedging_pnl_df.T['P&L JGB'].values]]
        # Risk-Free Rate Curve 1Y-50Y
        jgb_total_rfr_curve = sum(jgb_rfr_curve)
        # Total RFR Curve
        jgb_convexity = convexities_df.T['Convexity P&L'][3] / mm
        # Total Convexity
        jgb_credit_spread = 0.0
        # Credit Spread
        jgb_fx_spot = 0.0
        # FX Spot
        jgb_trading_cost = 0.0
        # Trading Costs
        
        
        jd1 = att.data_sheet(filenames[1], dates[1])
        jd1 = jd1[jd1['Currency'] == 'JPY']
        jd2 = att.data_sheet(filenames[0], dates[0])
        jd2 = jd2[jd2['Currency'] == 'JPY']
        jgb_true_total = (sum(jd1['Unrealised P&L']) - sum(jd2['Unrealised P&L'])) / mm
        # True Total
        
        jgb_residual = (jgb_true_total - 
                        jgb_trading_cost - 
                        jgb_credit_spread - 
                        jgb_convexity - 
                        jgb_total_rfr_curve -
                        jgb_fx_spot -
                        jgb_total_carry)
        # Residual
        
        jgb_column = [jgb_total_carry,
                      jgb_rates_carry,
                      jgb_credit_spread_carry,
                      jgb_fx_basis_carry,
                      jgb_total_rfr_curve,
                      *[i for i in jgb_rfr_curve],
                      jgb_convexity,
                      jgb_credit_spread,
                      jgb_fx_spot,
                      jgb_trading_cost,
                      jgb_residual,
                      jgb_true_total]
        
        [ws.write(f'D{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(31, 54), jgb_column)]

# ------------------------------------------------------------------------------------------------------ #
# USDJPY Forward [mJPY] (fwd) 
    # NOTE: The current residual value is incorrect as we are waiting on data
    # from J. Dass regarding the new basis and notional values for Dec 2022.
        
        fwd_rates_carry = self.trcrcf_df['RFR Carry'][6] / mm
        # Rates Carry
        fwd_credit_spread_carry = self.trcrcf_df['Credit Spread Carry'][6] / mm
        # Credit Spread Carry
        fwd_fx_basis_carry = - self.trcrcf_df['FX Basis Carry'][6] / mm
        # FX Basis Carry
        fwd_total_carry = sum([fwd_rates_carry, 
                               fwd_credit_spread_carry, 
                               fwd_fx_basis_carry])
        # Total Carry
        fwd_rfr_curve = [0.0] * 12
        # Risk-Free Rate Curve 1Y-50Y
        fwd_total_rfr_curve = sum(fwd_rfr_curve)
        # Total RFR Curve
        fwd_convexity = float(convexities_df.T['Convexity P&L'][4].replace('---', str(0.0))) / mm
        # Total Convexity
        fwd_credit_spread = 0.0
        # Credit Spread
        fwd_fx_spot = - self.fx_rp_plbond_fwd['FX P&L Bond'][1] * bbg_value / mm
        # FX Spot
        fwd_trading_cost = 0.0
        # Trading Costs
        
        fwd_true_total = - self.fx_rp_plbond_fwd['FX Forward'][0] / mm
        # True Total
        
        fwd_residual = (fwd_true_total - 
                        fwd_trading_cost - 
                        fwd_credit_spread - 
                        fwd_convexity - 
                        fwd_total_rfr_curve -
                        fwd_fx_spot -
                        fwd_total_carry)
        # Residual
        
        fwd_column = [fwd_total_carry,
                      fwd_rates_carry,
                      fwd_credit_spread_carry,
                      fwd_fx_basis_carry,
                      fwd_total_rfr_curve,
                      *[i for i in fwd_rfr_curve],
                      fwd_convexity,
                      fwd_credit_spread,
                      fwd_fx_spot,
                      fwd_trading_cost,
                      fwd_residual,
                      fwd_true_total]
        [ws.write(f'E{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(31, 54), fwd_column)]
        
# ------------------------------------------------------------------------------------------------------ #
# JPY Bel [mJPY] (bel)
        
        bel_rates_carry = -trcrcf_df['RFR Carry'][7] / mm
        # Rates Carry
        bel_credit_spread_carry = trcrcf_df['Credit Spread Carry'][7] / mm
        # Credit Spread Carry
        bel_fx_basis_carry = - trcrcf_df['FX Basis Carry'][7] / mm
        # FX Basis Carry
        bel_total_carry = sum([bel_rates_carry, 
                               bel_credit_spread_carry, 
                               bel_fx_basis_carry])
        # Total Carry
        bel_rfr_curve = [i / mm for i in [*alm_hedging_pnl_df.T['P&L BEL'].values]]
        # Risk-Free Rate Curve 1Y-50Y
        bel_total_rfr_curve = sum(bel_rfr_curve)
        # Total RFR Curve
        bel_convexity = convexities_df.T['Convexity P&L'][5] / mm
        # Total Convexity
        bel_credit_spread = 0.0
        # Credit Spread
        bel_fx_spot = 0.0
        # FX Spot
        bel_trading_cost = 0.0
        # Trading Costs
        
        bel_true_total = - (bel_carry_df['MtM'][-1] - bel_carry_df['MtM'][0]) / mm
        # True Total
        
        bel_residual = (bel_true_total - 
                        bel_trading_cost - 
                        bel_credit_spread - 
                        bel_convexity - 
                        bel_total_rfr_curve -
                        bel_fx_spot -
                        bel_total_carry)
        # Residual
        
        bel_column = [bel_total_carry,
                      bel_rates_carry,
                      bel_credit_spread_carry,
                      bel_fx_basis_carry,
                      bel_total_rfr_curve,
                      *[i for i in bel_rfr_curve],
                      bel_convexity,
                      bel_credit_spread,
                      bel_fx_spot,
                      bel_trading_cost,
                      bel_residual,
                      bel_true_total]
        [ws.write(f'F{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(31, 54), bel_column)]

# ------------------------------------------------------------------------------------------------------ #
# Net JPY Exposure [mJPY]

        net_jpy_x_jpy = [a+b+c+d for a, b, c, d in zip(jri_column, jgb_column, fwd_column, bel_column)]
        [ws.write(f'G{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(31, 54), net_jpy_x_jpy)]
    
# ------------------------------------------------------------------------------------------------------ #
# Net JPY Exposure [mUSD]

        net_jpy_x_usd = [i / bbg_value for i in net_jpy_x_jpy]
        [ws.write(f'H{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(31, 54), net_jpy_x_usd)]
        
# ------------------------------------------------------------------------------------------------------ #
# Net P&L B/S of Kiku [mJPY]
    
        net_pnl_jpy = [(i * bbg_value) + j for i, j in zip(net_usd_exp_column, net_jpy_x_jpy)]
        [ws.write(f'H{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(6, 31), net_pnl_jpy)]
        
# ------------------------------------------------------------------------------------------------------ #
# Net P&L B/S of Kiku [mUSD]
        
        net_pnl_usd = [i / bbg_value for i in net_pnl_jpy]
        [ws.write(f'I{i}', round(float(f'{x}'), 1)) for i, x in zip(np.arange(6, 31), net_pnl_usd)]
    
# ------------------------------------------------------------------------------------------------------ #
# % Risk Exposure

        net_pnl_usd_true_total = net_pnl_usd[-1]
        print(net_pnl_usd_true_total)
        rsk_exp = [i / net_pnl_usd_true_total for i in net_pnl_usd]
        [ws.write(f'J{i}', round(float(f'{x}'), 5)) for i, x in zip(np.arange(6, 31), rsk_exp)]        
        
# ------------------------------------------------------------------------------------------------------ #
# TE Hedging

        ws.write('K30', round(float(f'{sum(sum([*te_hedge_fx.T.values])) / 1000000}'), 1))
        ws.write('J30', 'TE - FX Hedging [mUSD]')
        
        # ws.write('K31', round(float(f'{}'))
        
        print(sum(self.swap_spread_df.T.values))
    
        wb.close()