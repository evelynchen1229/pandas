
import pandas as pd


# directory of raw data

financials = pd.read_csv('Raw Data/Raw_Financial Base.csv')

credit = pd.read_csv('Raw Data/RAW_credit.csv')

#print(financials.head(5))
#print(credit.head(5))
cus_without_credit = (financials[['CUSTOMER','ACCOUNT']].merge(credit, how = 'left', left_on = 'ACCOUNT', right_on='ACCNT_LEGCY_ID')
        .groupby('CUSTOMER')[['CREDIT_CY_ON','CREDIT_CY_OA','CREDIT_CY_PR','CREDIT_PY_ON','CREDIT_PY_OA','CREDIT_PY_PR']].sum()
        .reset_index()
        )
#print(cus_without_credit)

financial_without_credit = (financials.merge(credit, how = 'left',left_on = 'ACCOUNT', right_on='ACCNT_LEGCY_ID')
        .merge(cus_without_credit, how='left', left_on ='CUSTOMER', right_on ='CUSTOMER')

        .assign(adj_lbu_cy_on_amt = lambda x: x['LBU_CY_ON_AMT'].fillna(0) - x['CREDIT_CY_ON_x'].fillna(0))
        .assign(adj_lbu_cy_oa_amt = lambda x: x['LBU_CY_OA_AMT'].fillna(0) - x['CREDIT_CY_OA_x'].fillna(0))
        .assign(adj_lbu_cy_pr_amt = lambda x: x['LBU_CY_PR_AMT'].fillna(0) - x['CREDIT_CY_PR_x'].fillna(0))

        .assign(adj_lbu_py_on_amt = lambda x: x['LBU_PY_ON_AMT'].fillna(0) - x['CREDIT_PY_ON_x'].fillna(0))
        .assign(adj_lbu_py_oa_amt = lambda x: x['LBU_PY_OA_AMT'].fillna(0) - x['CREDIT_PY_OA_x'].fillna(0))
        .assign(adj_lbu_py_pr_amt = lambda x: x['LBU_PY_PR_AMT'].fillna(0) - x['CREDIT_PY_PR_x'].fillna(0))


        .assign(adj_cus_cy_on_amt = lambda x: x['CUS_CY_ON_AMT'].fillna(0) - x['CREDIT_CY_ON_y'].fillna(0))
        .assign(adj_cus_cy_oa_amt = lambda x: x['CUS_CY_OA_AMT'].fillna(0) - x['CREDIT_CY_OA_y'].fillna(0))
        .assign(adj_cus_cy_pr_amt = lambda x: x['CUS_CY_PR_AMT'].fillna(0) - x['CREDIT_CY_PR_y'].fillna(0))

        .assign(adj_cus_py_on_amt = lambda x: x['CUS_PY_ON_AMT'].fillna(0) - x['CREDIT_PY_ON_y'].fillna(0))
        .assign(adj_cus_py_oa_amt = lambda x: x['CUS_PY_OA_AMT'].fillna(0) - x['CREDIT_PY_OA_y'].fillna(0))
        .assign(adj_cus_py_pr_amt = lambda x: x['CUS_PY_PR_AMT'].fillna(0) - x['CREDIT_PY_PR_y'].fillna(0))
        )



print(financial_without_credit.columns)
financial_without_credit.to_csv('financial_withou_credit.csv',index=False)



