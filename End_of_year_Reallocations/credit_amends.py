''' THE BELOW FETCHES DATA FROM RECON (CHB), GCRM AND LAW. FILES ARE OUTPUTTED IN THE RAW DATA FOLDER
    AND ARE CONSEQUENTLY PROCESSED IN PANDAS
'''

import pandas as pd
import csv
import datetime as dt
from pandas import ExcelWriter

## CREDIT DAT FROM LAW
dsn_tns_base = cx_Oracle.makedsn('PSDB3684.LEXIS-NEXIS.COM', '1521', service_name='GBIPRD1.ISPPROD.LEXISNEXIS.COM')
conn_base = cx_Oracle.connect(user='DATAANALYTICS', password='DatPwd123Z', dsn=dsn_tns_base)


period = input('Enter current period (YYYYMM): ')

query_cy_st = f"""
select distinct per_start_dt from law.d_period_dt where per_wid =  to_char(add_months(to_date({period},'yyyymm'),-11),'yyyymm') and calendar =  'Avatar 445'

"""
cy_st = pd.read_sql(query_cy_st, con=conn_base)
cy_yr_st = cy_st['PER_START_DT'][0].strftime('%Y-%m-%d')

query_py_st = f"""
select distinct per_start_dt from law.d_period_dt where per_wid =  to_char(add_months(to_date({period},'yyyymm'),-23),'yyyymm') and calendar =  'Avatar 445'

"""
py_st = pd.read_sql(query_py_st, con=conn_base)
py_yr_st = py_st['PER_START_DT'][0].strftime('%Y-%m-%d')


query_credit = """
with sales_ops_product AS
(
        SELECT
            prod.row_wid,
            prod.prod_cd,
            prod.prod_famly,
            prod.prod_name,
            CASE
                WHEN prod.fin_report_cd != 'Unspecified' THEN 'ON'
                WHEN prod.fin_report_cd = 'Unspecified' AND
                prod.prod_medium_cd IN ('MC', 'PR','CD','DI') THEN 'PR'
                WHEN prod.fin_report_cd = 'Unspecified' AND
                prod.prod_medium_cd = 'OA' THEN 'OA'
                WHEN prod.fin_report_cd = 'Unspecified' AND
                prod.prod_medium_cd = 'ON' AND
                (
                --Product Code Detail Criteria
                (prod.prod_cd LIKE 'LEGM%' OR
                prod.prod_cd LIKE 'PSLC%' OR
                prod.prod_cd LIKE 'PSLD%' OR
                prod.prod_cd LIKE 'TSO%' OR
                prod.prod_cd LIKE 'TOLG%' OR
                prod.prod_cd LIKE 'TOLL%' OR
                prod.prod_cd LIKE 'MELM%')
                OR prod.prod_cd IN ('BXSUBS','IHSUBS','VASUBS','VCSUBS','VRSUBS','VVSUBS','WASUBS','WCSUBS','WDSUBS','WESUBS', --webinars
                'WFSUBS','WISUBS','WJSUBS','WNSUBS','WOSUBS','WPSUBS','WRSUBS','WSSUBS','WTSUBS','WUA','WVSUBS','WWSUBS','WXSUBS')
                -- Product Family Criteria
                OR (prod.spr_product_family IN ('LexisLibrary','Core', 'SmartForms', 'Middle East Online','LexisSmart', 'PSL/LPA','LexisPSL'))) THEN 'ON'
            ELSE 'OA'
            END AS salesops_medcodeclass
        FROM law.d_product prod
        WHERE 1=1
            AND (prod.spr_exclusion !='Y' OR prod.prod_cd IN ('WAENX', 'FTACA'))
),
 credit as (
 select accnt_legcy_id,
  sum(CAE_sub) cae_flg,
  sum(CY_ON) Credit_CY_ON,
  sum(CY_PR) Credit_CY_PR,
  sum(CY_OA) Credit_CY_OA,
  sum(PY_ON) Credit_PY_ON,
  sum(PY_PR) Credit_PY_PR,
  sum(CY_OA) Credit_PY_OA
  from
  (select credit.accnt_legcy_id,salesops_medcodeclass,fy,SUB_STAT,
  case when salesops_medcodeclass='ON' and fy='cy' then "Invoice Amount" end as CY_ON,
  case when salesops_medcodeclass='PR' and fy='cy' then "Invoice Amount" end as CY_PR,
  case when salesops_medcodeclass='OA' and fy='cy' then "Invoice Amount" end as CY_OA,
  case when salesops_medcodeclass='ON' and fy='py' then "Invoice Amount" end as PY_ON,
  case when salesops_medcodeclass='PR' and fy='py' then "Invoice Amount" end as PY_PR,
  case when salesops_medcodeclass='OA' and fy='py' then "Invoice Amount" end as PY_OA,
  case when SUB_STAT = 'Cancel At End' then 1 else 0 end as CAE_sub
from(
SELECT
  INV.ROW_WID, acc.ACCNT_LEGCY_ID,prod.prod_cd,sop.salesops_medcodeclass,sub.SUB_STAT,
  INV.INTEGRATION_ID AS "Invoice Line", INV.ORIG_INV_NUM AS "Original Invoice #", INV.DT_INV "Invoice Date",
  F.DOC_AMT AS "Invoice Amount", ORIG.INTEGRATION_ID AS "Original Invoice", ORIG.DT_INV AS "Original Invoice Date",
  case when INV.DT_INV between TO_DATE('{0}', 'YYYY-MM-DD') and (to_date ('{1}','YYYY-MM-DD')-1)-- Need to change this to whatever date we are going to use - likely first day of avatar year
  AND ORIG.DT_INV < TO_DATE('{0}', 'YYYY-MM-DD') then 'py'
  when  INV.DT_INV >= TO_DATE('{1}', 'YYYY-MM-DD') -- Need to change this to whatever date we are going to use - likely first day of avatar year
  AND ORIG.DT_INV < TO_DATE('{1}', 'YYYY-MM-DD')
  then 'cy'
  end as fy
FROM
  law.D_INVOICE_LN INV
  INNER JOIN law.F_INVOICE_LN F ON F.ROW_WID = INV.ROW_WID
  INNER JOIN law.D_INVOICE ORIG ON ORIG.ROW_WID = F.IN_ORIG_WID --  if doesn't work, ETL lookup must be incorrect use ORIG.INTEGRATION_ID = INV.ORIG_INV_NUM as a work around
  inner join law.d_fin_accnt_x acc on acc.row_wid = f.fa_wid
  inner join law.d_product prod on prod.row_wid = f.prod_wid
  inner join sales_ops_product sop ON sop.row_wid=prod.row_wid
  left join law.D_SUBSCR_REV sub on sub.row_wid = f.subr_wid
WHERE
  INV.DELETE_FLG = 'N' -- Invoice has not been deleted
  AND INV.DATASOURCE_NUM_ID = 10 -- Only Invoices from Genesis
  AND INV.ORIG_INV_NUM IS NOT NULL
  AND INV.ORIG_INV_NUM <> '         ' -- Looks like Genesis has this value instead of NULL
 /* AND INV.DT_INV >= TO_DATE('2021-01-01', 'YYYY-MM-DD') -- Need to change this to whatever date we are going to use - likely first day of avatar year
  AND ORIG.DT_INV < TO_DATE('2021-01-01', 'YYYY-MM-DD') -- Need to change this to whatever date we are going to use - likely first day of avatar year
  */
  --test
  AND INV.DT_INV >= TO_DATE('{0}', 'YYYY-MM-DD') -- Need to change this to whatever date we are going to use - likely first day of avatar year
  AND ORIG.DT_INV < TO_DATE('{0}', 'YYYY-MM-DD') -- Need to change this to whatever date we are going to use - likely first day of avatar year

  AND F.DOC_AMT < 0 -- Only Credits
  --and ACC.accnt_legcy_id = 'AGBA5003'
  ORDER BY INV.DT_INV, INV.INTEGRATION_ID)credit
)
  group by accnt_legcy_id
)
select *
from credit

""".format(py_yr_st,cy_yr_st)
df_credit = pd.read_sql(query_credit, con=conn_base)
output_credit =  os.getcwd()+'\\Raw Data\\Raw_credit.csv'
df_credit.to_csv(output_credit, index=False)


conn_base.close()





