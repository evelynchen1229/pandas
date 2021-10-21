import pandas as pd
import numpy as np

# directory of raw data

#financials = pd.read_csv('Raw Data/Raw_Financial Base.csv')
financials = pd.read_csv('test_billing.csv')

#credit = pd.read_csv('Raw Data/RAW_credit.csv')
credit = pd.read_csv('test_credit.csv')

f_cols = ['LBU_CY_ON_AMT','LBU_CY_OA_AMT','LBU_CY_PR_AMT','LBU_PY_ON_AMT','LBU_PY_OA_AMT','LBU_PY_PR_AMT','LBU_SUBSCR','LBU_CAE_VAL']

c_cols = ['CREDIT_CY_ON','CREDIT_CY_OA','CREDIT_CY_PR','CREDIT_PY_ON','CREDIT_PY_OA','CREDIT_PY_PR','ACTIVE_SUBS_VALUE','CAE_SUBS_VALUE','ACCNT_LEGCY_ID']

new_cols = {'adj_lbu_cy_on_amt': f_cols[0],
            'adj_lbu_cy_oa_amt': f_cols[1],
            'adj_lbu_cy_pr_amt': f_cols[2],
            'adj_lbu_py_on_amt': f_cols[3],
            'adj_lbu_py_oa_amt': f_cols[4],
            'adj_lbu_py_pr_amt': f_cols[5],
            'adj_lbu_subscr': f_cols[6],
            'adj_lbu_cae_val': f_cols[7]}

financial_without_credit = (financials.merge(credit, how = 'left',left_on = 'ACCOUNT', right_on='ACCNT_LEGCY_ID')
        #.merge(cus_without_credit, how='left', left_on ='CUSTOMER', right_on ='CUSTOMER')

        # get adjusted lbu cy amount
        .assign(adj_lbu_cy_on_amt = lambda x: x[f_cols[0]].fillna(0) - x[c_cols[0]].fillna(0))
        .assign(adj_lbu_cy_oa_amt = lambda x: x[f_cols[1]].fillna(0) - x[c_cols[1]].fillna(0))
        .assign(adj_lbu_cy_pr_amt = lambda x: x[f_cols[2]].fillna(0) - x[c_cols[2]].fillna(0))

        # get adjusted lbu py amount
        .assign(adj_lbu_py_on_amt = lambda x: x[f_cols[3]].fillna(0) - x[c_cols[3]].fillna(0))
        .assign(adj_lbu_py_oa_amt = lambda x: x[f_cols[4]].fillna(0) - x[c_cols[4]].fillna(0))
        .assign(adj_lbu_py_pr_amt = lambda x: x[f_cols[5]].fillna(0) - x[c_cols[5]].fillna(0))

        # get adjusted lbu cae and active sub values
        .assign(adj_lbu_subscr = lambda x: x[f_cols[6]].fillna(0) - x[c_cols[6]].fillna(0))
        .assign(adj_lbu_cae_val = lambda x: x[f_cols[7]].fillna(0) - x[c_cols[7]].fillna(0))

         # get lbu total value for cy and py
        .assign(lbu_cy_tot = lambda x: x['adj_lbu_cy_on_amt'] + x['adj_lbu_cy_oa_amt'] + x['adj_lbu_cy_pr_amt'])
        .assign(lbu_py_tot = lambda x: x['adj_lbu_py_on_amt'] + x['adj_lbu_py_oa_amt'] + x['adj_lbu_py_pr_amt'])

        # get excldclosedacct flag
        .assign(excldclosedacct = lambda x: np.where((x['REPCODE'] == 'CL') & (x['lbu_py_tot'] == 0),'Y','N'))

        # get cust level billing
        .assign(cust_subscr = lambda x: x.groupby('CUSTOMER').adj_lbu_subscr.transform('sum'))
        .assign(cust_cae_val = lambda x: x.groupby('CUSTOMER').adj_lbu_cae_val.transform('sum'))
        .assign(cust_cae_p = lambda x: x['cust_cae_val'] / x['cust_subscr'])

        .assign(cus_cy_on_amt = lambda x: x.groupby('CUSTOMER').adj_lbu_cy_on_amt.transform('sum'))
        .assign(cus_cy_oa_amt = lambda x: x.groupby('CUSTOMER').adj_lbu_cy_oa_amt.transform('sum'))
        .assign(cus_cy_pr_amt = lambda x: x.groupby('CUSTOMER').adj_lbu_cy_pr_amt.transform('sum'))
        .assign(cus_cy_tot = lambda x: x.groupby('CUSTOMER').lbu_cy_tot.transform('sum'))

        .assign(cus_py_on_amt = lambda x: x.groupby('CUSTOMER').adj_lbu_py_on_amt.transform('sum'))
        .assign(cus_py_oa_amt = lambda x: x.groupby('CUSTOMER').adj_lbu_py_oa_amt.transform('sum'))
        .assign(cus_py_pr_amt = lambda x: x.groupby('CUSTOMER').adj_lbu_py_pr_amt.transform('sum'))
        .assign(cus_py_tot = lambda x: x.groupby('CUSTOMER').lbu_py_tot.transform('sum'))

        .assign(cus_has_cy_online_spend = lambda x: np.where(x['cus_cy_on_amt'] > 0, 'Y', 'N'))
        .assign(cus_has_py_online_spend = lambda x: np.where(x['cus_py_on_amt'] > 0, 'Y', 'N'))

        # add unique id based on cust number and lbu number
        .assign(unique_ID = lambda x: x['CUSTOMER'] + x['ACCOUNT'])

        # drop orginal lbu spend and credit
        .drop(columns = f_cols + c_cols)
        .rename(columns = new_cols)

        )

# exclude closed accounts
financials = financial_without_credit[financial_without_credit['excldclosedacct'] == 'N']
financials.columns = map(str.upper, financials.columns)
#print(financial_without_credit.columns)
#financial_without_credit.to_csv('financial_without_credit.csv',index=False)
#print(financial_without_credit['unique_ID'])

financial_without_credit.to_csv('financial_without_credit_wip.csv',index=False)
financials.to_csv('financials.csv',index=False)


