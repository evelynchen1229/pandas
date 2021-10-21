''' THE BELOW FETCHES DATA FROM RECON (CHB), GCRM AND LAW. FILES ARE OUTPUTTED IN THE RAW DATA FOLDER
    AND ARE CONSEQUENTLY PROCESSED IN PANDAS
'''

import cx_Oracle
import pandas as pd
import csv
import datetime as dt
import win32com.client
from pandas import ExcelWriter
import os

# MARKETABLE CONTACTS - queries views built by Mike Thomas to retrieve the number of Marketable contacts for each customer
## Contact Mike for further info
dsn_tns_ms_contacts = cx_Oracle.makedsn('psdb3119.lexisnexis.com', '1521', service_name='CUSTHUB.ispprod.lexisnexis.com') 
conn_ms_contacts = cx_Oracle.connect(user='recon', password='Recon1121Z', dsn=dsn_tns_ms_contacts) 

query_ms_contacts = """SELECT 
                            crm_cust_id as CUSTOMER,
                            "# M/S Contacts" as "Marketable Contact Flag"
                        FROM
                            mv_customer_contacts
                        WHERE
                            customer_has_ms_contact = 'Y'
                        """

ms_contacts = pd.read_sql(query_ms_contacts, con=conn_ms_contacts)
output_file_ms_contacts =  os.getcwd()+'\\Raw Data\\MARKETABLE CONTACTS.csv'
ms_contacts.to_csv(output_file_ms_contacts, index=False)

# FEE EARNERS AND TAX PARTNERS - same as above, the views might have changed since this run

dsn_tns_fe = cx_Oracle.makedsn('psdb3119.lexisnexis.com', '1521', service_name='CUSTHUB.ispprod.lexisnexis.com') 
conn_fe = cx_Oracle.connect(user='recon', password='Recon1121Z', dsn=dsn_tns_fe) 

query_fe = """SELECT
                  crm_cust_id             AS "CRM Org ID",
                  number_of_fee_earners   AS "Legal_FE",
                  number_of_partners      AS "Tax_FE"
              FROM
                  mv_customer_contacts
              WHERE
                  ( number_of_fee_earners IS NOT NULL
                    OR number_of_partners IS NOT NULL )
                        """

fe = pd.read_sql(query_fe, con=conn_fe)
output_file_fe =  os.getcwd()+'\\Raw Data\\FE DATA.csv'
fe.to_csv(output_file_fe, index=False)

# NEXIS BILLINGS - from GCRM, query built by Mike Thomas
dsn_tns_nexis = cx_Oracle.makedsn('psdb33540-vip.lexis-nexis.com', '1521', service_name='SBL_SERVICE.isprod.lexisnexis.com') 
conn_nexis = cx_Oracle.connect(user='SDI', password='SDIPROD63', dsn=dsn_tns_nexis) 

query_nexis = """select
customer,
customer_name,
--Product_Class,
sum(value_doc)
from
(
SELECT
  ORG.LOC AS Customer
  , ORG.NAME AS Customer_Name
  , ORG.X_CUSTOMER_SUBCLASS AS "Subclass", ORG.X_SEC_SUBCLASS AS "Secondary Subclass"
  , EMP.LAST_NAME AS "Last Name", EMP.FST_NAME AS "First Name", EMP.job_title as "Position"
  , AGR.AGREE_NUM||':'||AGR.REV_NUM AS "Agreement", AGR.X_MULTI_TERM_FLG AS MYD, AGR.X_MYD_MAX_END_DT AS "MYD End"
  , AGR.EFF_START_DT AS "Sub Start", AGR.EFF_END_DT AS "Sub End"
  , AGR_CON.LAST_NAME AS "Contact Last Name", AGR_CON.FST_NAME AS "Contact First Name", AGR_CON.EMAIL_ADDR AS "Contact Email", AGR_CON.WORK_PH_NUM AS "Contact Phone"
  , ROUND(MONTHS_BETWEEN(AGR.EFF_END_DT+1, AGR.EFF_START_DT), 1) AS "Length"
  , AGR.X_TERM_TYPE AS "Term"
  , AGR.X_AAR_PCT AS "AAR%"
  , ORD.X_ORDER_SUB_TYPE AS "Order Type"
  --, MAX(CASE WHEN PROD.INTEGRATION_ID = 'urn:product:1519442' THEN 'Y' ELSE 'N' END) OVER (PARTITION BY ORG.ROW_ID, AGR.ROW_ID) AS "Has Nexis Core Feature"
  , PROD.INTEGRATION_ID AS "Product PGUID", PROD.NAME AS "Product Name", PROD.X_PRODUCT_CLASS AS Product_Class, PROD.X_GO_TO_MARKET_TYPE AS "GTM Type", PROD.X_ITEM_TYPE AS "Item Type", PROD.X_PKG_LVL AS "Package Level", PROD.X_PRODUCT_TYPE AS "Product Type"
--  , AGR.BL_CURCY_CD AS FX, AGR.X_MONTLY_NET_PRICE AS "Total Agreement Value", round(AGR.X_MONTLY_NET_PRICE /  (ROUND(MONTHS_BETWEEN(AGR.EFF_END_DT+1, AGR.EFF_START_DT), 1) ),2) as "Monthly Commitment"
  , NVL(OI.X_EXT_NET_PRI, OI.QTY_REQ * OI.NET_PRI) AS Value_Doc
 -- , ROUND(NVL(OI.X_EXT_NET_PRI, OI.QTY_REQ * OI.NET_PRI) * (12/NULLIF(ROUND(MONTHS_BETWEEN(AGR.EFF_END_DT+1, AGR.EFF_START_DT), 1), 0)), 2) AS "Annualized (Doc)"
--  , NVL(REL_QI.X_EXT_NET_PRI, REL_QI.QTY_REQ * REL_QI.NET_PRI) AS "Renewal (Doc)"
--  , ROUND(NVL(REL_QI.X_EXT_NET_PRI, REL_QI.QTY_REQ * REL_QI.NET_PRI) * (12/NULLIF(ROUND(MONTHS_BETWEEN(REL_Q.EFF_END_DT+1, REL_Q.EFF_START_DT), 1), 0)), 2) AS "Annualized Renewal (Doc)"
 -- , ROUND(CASE AGR.BL_CURCY_CD WHEN 'GBP' THEN 1.28 WHEN 'EUR' THEN 1.12281 ELSE 1 END * NVL(OI.X_EXT_NET_PRI, OI.QTY_REQ * OI.NET_PRI), 2) AS "Value (USD)"
  --, ROUND(CASE AGR.BL_CURCY_CD WHEN 'GBP' THEN 1.28 WHEN 'EUR' THEN 1.12281 ELSE 1 END * NVL(OI.X_EXT_NET_PRI, OI.QTY_REQ * OI.NET_PRI) * (12/NULLIF(ROUND(MONTHS_BETWEEN(AGR.EFF_END_DT+1, AGR.EFF_START_DT), 1), 0)), 2) AS "Annualized (USD)"
 -- , ROUND(CASE AGR.BL_CURCY_CD WHEN 'GBP' THEN 1.28 WHEN 'EUR' THEN 1.12281 ELSE 1 END * NVL(REL_QI.X_EXT_NET_PRI, REL_QI.QTY_REQ * REL_QI.NET_PRI), 2) AS "Renewal (USD)"
--  , ROUND(CASE AGR.BL_CURCY_CD WHEN 'GBP' THEN 1.28 WHEN 'EUR' THEN 1.12281 ELSE 1 END * NVL(REL_QI.X_EXT_NET_PRI, REL_QI.QTY_REQ * REL_QI.NET_PRI) * (12/NULLIF(ROUND(MONTHS_BETWEEN(REL_Q.EFF_END_DT+1, REL_Q.EFF_START_DT), 1), 0)), 2) AS "Annualized Renewal (USD)"
FROM
  SIEBEL.S_DOC_AGREE AGR
  INNER JOIN SIEBEL.S_ORG_EXT ORG ON ORG.ROW_ID = AGR.TARGET_OU_ID
  INNER JOIN SIEBEL.S_BU BU ON BU.ROW_ID = ORG.BU_ID
  INNER JOIN SIEBEL.S_ORDER ORD ON ORD.ROW_ID = AGR.ORDER_ID
  INNER JOIN SIEBEL.S_ORDER_ITEM OI ON OI.ORDER_ID = ORD.ROW_ID
  INNER JOIN SIEBEL.S_PROD_INT PROD ON PROD.ROW_ID = OI.PROD_ID
  INNER JOIN SIEBEL.S_POSTN POS ON POS.ROW_ID = ORG.PR_POSTN_ID
  INNER JOIN SIEBEL.S_CONTACT EMP ON EMP.ROW_ID = POS.PR_EMP_ID
  INNER JOIN SIEBEL.S_CONTACT AGR_CON ON AGR_CON.PAR_ROW_ID = AGR.CON_PER_ID
  LEFT OUTER JOIN SIEBEL.S_DOC_QUOTE REL_Q ON REL_Q.X_REL_QUOTE_ID = AGR.QUOTE_ID
  LEFT OUTER JOIN SIEBEL.S_QUOTE_ITEM REL_QI ON REL_QI.SD_ID = REL_Q.ROW_ID AND REL_QI.PROD_ID = OI.PROD_ID
WHERE 1=1
  AND BU.NAME = 'United Kingdom'
  AND NVL(AGR.X_TRIAL_FLG, 'N') = 'N'     -- No Trials
  AND AGR.EFF_END_DT > SYSDATE            -- Renews in future
  AND AGR.STAT_CD = 'Active'              -- Active
  AND PROD.INTEGRATION_ID NOT IN
    (
      'urn:product:1523691','urn:product:1523692','urn:product:1523693','urn:product:1523694','urn:product:1523695','urn:product:1523696','urn:product:1523697', 'urn:product:1523698'
      ,'urn:product:1523699','urn:product:1523700','urn:product:1525008','urn:product:1525009','urn:product:1525010','urn:product:1525618','urn:product:1526114' -- Remove international content
      ,'urn:product:1515564' -- Remove UK Core Feature
    )
ORDER BY
  ORG.LOC, AGR.EFF_END_DT, AGR.EFF_START_DT, AGR.AGREE_NUM, AGR.REV_NUM, PROD.INTEGRATION_ID)
  where 1=1 and nvl(Product_Class,'pippo') != 'TotalPatent'
 -- and customer = '42543YBWK'
  group by
  customer,
customer_name"""


nexis = pd.read_sql(query_nexis, con=conn_nexis)
output_file_nexis =  os.getcwd()+'\\Raw Data\\Raw_Nexis Billings.csv'
nexis.to_csv(output_file_nexis, index=False)


##CUST MAPPING WITH CUSTOMERS FROM NEXIS BILLINGS - the below query in LAW maps the old and new customers
## However, the mapping doesn't always align. Keving McGowan has provided a manually mapped file in 2020.
cust_list = nexis['CUSTOMER'].values
cust_list = tuple(cust_list)

dsn_tns = cx_Oracle.makedsn('PSDB3684.LEXIS-NEXIS.COM', '1521', service_name='GBIPRD1.ISPPROD.LEXISNEXIS.COM') 
conn = cx_Oracle.connect(user='DATAANALYTICS', password='DatPwd123Z', dsn=dsn_tns) 

query = """select distinct  c.accnt_num,  c.accnt_name, cx.integration_id, dacx.accnt_legcy_id
 from LAW.f_fin_accnt facc
join law.d_customer c on facc.cus_wid = c.row_wid
left join law.d_customer_x cx on facc.cus_wid = cx.row_wid
left join LAW.d_fin_accnt_x dacx on dacx.row_wid = facc.row_wid
where  c.accnt_num in{}""".format(cust_list)
df = pd.read_sql(query, con=conn)
output_mapping =  os.getcwd()+'\\Raw Data\\Raw_Customer Mapping.csv'
df.to_csv(output_mapping, index=False)


## BASE DATA FROM LAW - query might need to be updated, comments in the SQL file in the folder
dsn_tns_base = cx_Oracle.makedsn('PSDB3684.LEXIS-NEXIS.COM', '1521', service_name='GBIPRD1.ISPPROD.LEXISNEXIS.COM') 
conn_base = cx_Oracle.connect(user='DATAANALYTICS', password='DatPwd123Z', dsn=dsn_tns_base) 


period = input('Enter current period (YYYYMM): ')
query_base = """

with cust_det as 

        (select distinct
        c.provider_customer_key customer ,
        c.customer custname,
        c.status as cust_status,
        xcust.accnt_stat as law_customer_status,
        a.provider_account_key account,
        a.account_name accountname,
        fa_accnt.account_status,
        c.customer_segment SEGMENT,
        fa_accnt.team,
        fa_accnt.repcode repcode,
        fa_accnt.fullname fullname,
        addrs.addr_5 AS region,
        addrs.addr_4 AS city,
        addrs.post_code_r AS postcode,
        addrs.addr_6 AS country
        from pdl.plu_vw_account a
        join pdl.plu_vw_customer c on a.customer_id = c.customer_id
        LEFT JOIN pdl.gen_r_customer_delivery cd ON cd.customer_account = a.provider_account_key and cd.SYS_CURRENT_FLG = 'Y' 
        LEFT JOIN pdl.gen_r_names_addresses addrs ON addrs.address_ptr = cd.address_ptr AND addrs.sys_current_flg = 'Y'
         left JOIN (
                            SELECT
                                xaccnt.accnt_legcy_id,
                                p.postn_rep_cd as repcode,
                                p.postn_team            AS team,
                                p.emp_name              AS fullname,
                                 xaccnt.ACCNT_STAT as account_status
                            FROM
                                law.f_fin_accnt     faccnt
                                INNER JOIN law.d_fin_accnt_x   xaccnt ON faccnt.row_wid = xaccnt.row_wid
                                INNER JOIN law.d_position      p ON p.row_wid = xaccnt.pos_pr_wid
                        ) fa_accnt ON a.provider_account_key = fa_accnt.accnt_legcy_id
                        
        left join law.d_customer_x xcust on xcust.integration_id = c.provider_customer_key
        where  1=1 
        and fa_accnt.repcode   not in ('MIAA', 'MIAC')
        and fa_accnt.fullname not in ('M LEX', 'Team RepCode')
        and fa_accnt.team != 'MI-MI'
        and not   (fa_accnt.repcode   ='--' or  c.customer_segment like '%Internal%')
           ) ,
           
sales_ops_product as
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
            OR prod.prod_cd in ('BXSUBS','IHSUBS','VASUBS','VCSUBS','VRSUBS','VVSUBS','WASUBS','WCSUBS','WDSUBS','WESUBS', --webinars
            'WFSUBS','WISUBS','WJSUBS','WNSUBS','WOSUBS','WPSUBS','WRSUBS','WSSUBS','WTSUBS','WUA','WVSUBS','WWSUBS','WXSUBS')
            -- Product Family Criteria
            OR (prod.SPR_PRODUCT_FAMILY IN ('LexisLibrary','Core', 'SmartForms', 'Middle East Online','LexisSmart', 'PSL/LPA','LexisPSL'))) THEN 'ON'
        ELSE 'OA'
        END AS salesops_medcodeclass
        FROM law.D_PRODUCT PROD
        WHERE 1=1
        and (prod.spr_exclusion !='Y' or prod.prod_cd in ('WAENX', 'FTACA'))
),
   
financials as 
        (SELECT
                customer,
                account,
                accountname,
                NVL(SUM(case when FY = 'CY' and salesops_medcodeclass = 'ON' THEN amt end),0) as CY_ON_AMT,
                NVL(SUM(case when FY = 'CY' and salesops_medcodeclass = 'PR' THEN amt end),0) as CY_PR_AMT,
                NVL(SUM(case when FY = 'CY' and salesops_medcodeclass = 'OA' THEN amt end),0) as CY_OA_AMT,
                
                NVL(SUM(case when FY = 'PY' and salesops_medcodeclass = 'ON' THEN amt end),0) as PY_ON_AMT,
                NVL(SUM(case when FY = 'PY' and salesops_medcodeclass = 'PR' THEN amt end),0) as PY_PR_AMT,
                NVL(SUM(case when FY = 'PY' and salesops_medcodeclass = 'OA' THEN amt end),0) as PY_OA_AMT,
                
                sum(case when salesops_medcodeclass = 'ON' then active_subs_value end) as active_subs_value,
                sum(case when salesops_medcodeclass = 'ON' then CAE_subs_value end) as CAE_subs_value,
                max(case when salesops_medcodeclass = 'ON' then  dt_renew end) as max_dt_renew,
                max(CASE WHEN myd_flg = 'Y' and dt_renew > yr_st  then 1 else 0 end) as myd_flg,
                max(CASE WHEN  salesops_medcodeclass = 'ON'  and nvl(dt_renew, to_date('1900-01-01','YYYY-MM-DD'))  > yr_st  then 1 else 0 end) as LBU_renewal_flg,
                
                sum(nexis_flag) as Maila_nexis_flag
         FROM 
                (SELECT 
                    sub_inv.integration_id as customer,
                    sub_inv.accnt_legcy_id as account,
                    sub_inv.accnt_name as accountname,
                    sub_inv.per_wid,
                    sub_inv.prod_cd,
                    case when sub_inv.per_wid  >= to_char(add_months(to_date({0},'yyyymm'),-11),'yyyymm') then 'CY' else 'PY' END AS FY,
                    -- AS WE HAVE A BESPOKE FY, THE START OF THE YEAR IS THE START OF THE PERIOD IDENTIFIED AS THE STARTING PERIOD FOR THE CY
                    case when sub_inv.per_wid >= to_char(add_months(to_date({0},'yyyymm'),-23),'yyyymm') then (select distinct per_start_dt from law.d_period_dt where per_wid =  to_char(add_months(to_date({0},'yyyymm'),-11),'yyyymm') and calendar =  'Avatar 445') end as yr_st,
                    case when sub_inv.sub_stat  in ('Active','Cancel At End','Complete at End') then  sub_inv.amt end as active_subs_value,
                    case when sub_inv.sub_stat  in ('Cancel At End') then  amt end as CAE_subs_value,
                    case when sub_inv.sub_stat  in ('Active','Cancel At End','Complete at End','Pending','Frozen') then   sub_inv.dt_renew end as dt_renew,
                    sub_inv.myd_flg,
                    case when sub_inv.prod_cd in ('WAENX', 'FTACA') then 1 else 0 end as nexis_flag,
                    sub_inv.salesops_medcodeclass,
                    sub_inv.f_row_wid,
                    sub_inv.amt
              -- SUB_INV -- ALL SUBS WITHOUT AN INVOICE UNIONED TO THE INVOICE DATA
              -- THE 'INVOICE' DATE OF THE SUBSCRIPTION WITH NO INVOICE HAS BEEN ARBITRARILY SET TO 30 DAYS BEFORE THE SUB START DATE
                FROM 
                ( select 
                        cust.integration_id,
                        dacc.accnt_legcy_id,
                        dacc.accnt_name,
                        per.per_wid,
                        prod.prod_cd,
                        prod.prod_famly,
                        prod.prod_name,
                        dsub.sub_stat,
                        dsub.dt_start,
                        dsub.dt_end,
                        dsub.canc_dt,
                        dsub.cae_date,
                        dsub.DT_RENEW,
                        dsub.myd_flg,
                        sop.salesops_medcodeclass,
                        fsub.row_wid as f_row_wid,
                        fsub.doc_amt AS amt,
                        'SUB' as src
                    from
                        law.f_subscr_rev fsub
                        join law.d_subscr_rev    dsub on dsub.row_wid = fsub.row_wid
                        join law.d_product prod on prod.row_wid = fsub.prod_wid
                        join sales_ops_product sop on sop.row_wid=prod.row_wid
                        join law.d_fin_accnt_x dacc on dacc.row_wid=fsub.fa_wid
                        JOIN law.d_customer_x cust  ON fsub.cus_wid = cust.row_wid
                        LEFT join law.f_invoice_ln inv on   dsub.row_wid = inv.subr_wid
                        -- the alleged invoice date should be 30 days before the end of the sub, but i decided 30 days before the start of the sub...
                        JOIN law.d_period_dt per ON to_char(to_date(fsub.dt_start_wid,'yyyymmdd') - 30,'yyyymmdd') = per.per_dt_wid 
                        join pdl.gen_s_line gens on gens.sys_pguid = dsub.pguid and gens.sys_current_flg = 'Y'
                    where 1=1
                        and per.calendar = 'Avatar 445'
                        and dsub.bu_pguid = 'UK'
                        and fsub.dt_start_wid != 0
                        --  EXCLUDING RECORDS WITH  AN INVOICE
                        and inv.row_wid is null
                        and nvl(gens.sys_delete_flg,'N')!='Y'
                        and  case when dsub.grts_rsn = 'Trial Request' and fsub.doc_amt <= 0 then 1 else 0 end != 1
                        and per.per_wid between to_char(add_months(to_date({0},'yyyymm'),-23),'yyyymm')  AND {0}
                
                 UNION ALL 
                 -- INVOICE DATA
                 SELECT
                    cust.integration_id,
                    acc.accnt_legcy_id,
                    acc.accnt_name,
                    per.per_wid,
                    prod.prod_cd,
                     prod.prod_famly,
                    prod.prod_name,
                    subrev.sub_stat,
                    subrev.dt_start,
                    subrev.dt_end,
                    subrev.canc_dt,
                    subrev.cae_date,
                    subrev.DT_RENEW,
                    subrev.myd_flg,
                    sop.salesops_medcodeclass,
                    inv.row_wid as f_row_wid,
                    inv.doc_amt AS amt,
                    'INV' as src
                FROM law.f_invoice_ln inv 
                INNER JOIN law.d_customer_x cust  ON inv.cus_wid = cust.row_wid
                INNER JOIN law.d_product prod ON inv.prod_wid = prod.row_wid
                join sales_ops_product sop on sop.row_wid=prod.row_wid
                inner JOIN law.d_subscr_rev subrev ON subrev.row_wid = inv.subr_wid
                INNER JOIN law.d_fin_accnt_x acc ON acc.row_wid = inv.fa_wid
                INNER JOIN law.d_period_dt per ON inv.dt_inv_wid = per.per_dt_wid
                --join law.d_product prod on inv.prod_wid = prod.row_wid 
                WHERE 1 = 1
                        AND inv.delete_flg = 'N'
                        AND (inv.bu_pguid IN ('UK', 'Unspecified'))
                        AND per.calendar = 'Avatar 445'
                        and  case when subrev.grts_rsn = 'Trial Request' and inv.doc_amt <= 0 then 1 else 0 end != 1
                        and per.per_wid between to_char(add_months(to_date({0},'yyyymm'),-23),'yyyymm')  AND {0}
                        ) sub_inv
                        
                        WHERE 1=1 
                )
             group by 
                customer,
                account,
                accountname
    )
                
    select * from     
          (  select 
                cust_det.customer,
                cust_det.custname,
                --cust_det.cust_status as customer_status,
                cust_det.law_customer_status,
                cust_det.account_status as account_status,
                cust_det.account,
                cust_det.accountname,
                cust_det.team,
                cust_det.repcode,
                cust_det.fullname,
                cust_det.segment,
                cust_det.city,
                cust_det.postcode,
                cust_det.region,
                cust_det.country,
                sum(financials.active_subs_value) AS LBU_Subscr,
                sum(CAE_subs_value) as LBU_CAE_VAL,
                SUM(sum(financials.active_subs_value)) over (partition by  CUST_DET.CUSTOMER) as Cust_Subscr,
                SUM(sum(CAE_subs_value)) OVER   (partition by  CUST_DET.CUSTOMER) as CUST_CAE_VAL,
                NULLIF(SUM(sum(CAE_subs_value)) OVER   (partition by  CUST_DET.CUSTOMER),0) /   SUM(sum(financials.active_subs_value)) over (partition by  CUST_DET.CUSTOMER) AS Cust_CAE_p,
               case when  nvl(SUM(SUM(CY_ON_AMT)) OVER (partition by  CUST_DET.CUSTOMER),0) > 0 then 'Y' else 'N' end AS CUS_HAS_CY_ONLINE_SPEND,
               case when  nvl(SUM(SUM(PY_ON_AMT)) OVER (partition by  CUST_DET.CUSTOMER),0) >  0 then 'Y' else 'N' end AS CUS_HAS_PY_ONLINE_SPEND,
                case when  nvl(SUM(SUM(LBU_renewal_flg)) OVER (partition by  CUST_DET.CUSTOMER),0) > 0 then 'Y' else 'N' end AS Cust_ON_RenewalFlg ,
                case when  nvl(SUM(financials.LBU_renewal_flg) ,0) > 0 then 'Y' else 'N' end AS LBU_ON_RenewalFlg ,
                
               case when sum(Maila_nexis_flag) > 0 then 'Y' else 'N' end as nexis_flg,
                case   WHEN SUM(myd_flg) > 0  then 'Y' else 'N' end AS  LBU_MYD_FLG,
                max(financials.max_dt_renew) as LBUmaxRenewDate,
               max(max(financials.max_dt_renew)) OVER (partition by  CUST_DET.CUSTOMER) as CustMaxRenewDate,
                -- CURRENT YEAR LBU
                NVL(SUM(financials.CY_ON_AMT),0) AS LBU_CY_ON_AMT,
                NVL(SUM(financials.CY_PR_AMT),0) AS LBU_CY_PR_AMT,
                NVL(SUM( financials.CY_OA_AMT),0) AS LBU_CY_OA_AMT,
                NVL( SUM(financials.CY_ON_AMT) + SUM(financials.CY_PR_AMT) + SUM( financials.CY_OA_AMT),0) AS LBU_CY_TOT,
                -- CURRENT YEAR CUST
                NVL(SUM(SUM(financials.CY_ON_AMT)) OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_CY_ON_AMT,
                NVL(SUM(SUM(financials.CY_PR_AMT)) OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_CY_PR_AMT,
                NVL(SUM(SUM( financials.CY_OA_AMT)) OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_CY_OA_AMT,
                NVL(SUM(SUM(financials.CY_ON_AMT)) OVER (PARTITION BY CUST_DET.CUSTOMER) + SUM(SUM(financials.CY_PR_AMT))  OVER (PARTITION BY CUST_DET.CUSTOMER) + SUM(SUM( financials.CY_OA_AMT))  OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_CY_TOT,
        
                -- PREVIOUS YEAR LBU
                NVL(SUM(financials.PY_ON_AMT),0) AS LBU_PY_ON_AMT,
                NVL(SUM(financials.PY_PR_AMT),0) AS LBU_PY_PR_AMT,
                NVL(SUM( financials.PY_OA_AMT),0) AS LBU_PY_OA_AMT,
                NVL( SUM(financials.PY_ON_AMT) + SUM(financials.PY_PR_AMT) + SUM( financials.PY_OA_AMT),0) AS LBU_PY_TOT,
                -- PREVIOUS YEAR CUST
                NVL(SUM(SUM(financials.PY_ON_AMT))  OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_PY_ON_AMT,
                NVL(SUM(SUM(financials.PY_PR_AMT))  OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_PY_PR_AMT,
                NVL(SUM(SUM( financials.PY_OA_AMT))  OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_PY_OA_AMT,
                NVL(SUM(SUM(financials.PY_ON_AMT)) OVER (PARTITION BY CUST_DET.CUSTOMER) + SUM(SUM(financials.PY_PR_AMT))  OVER (PARTITION BY CUST_DET.CUSTOMER) + SUM(SUM( financials.PY_OA_AMT))  OVER (PARTITION BY CUST_DET.CUSTOMER),0) AS CUS_PY_TOT,
                
                {0} as YrPer,
                
                -- NEED TO EXCLUDE CL REPCODES WITH NO SPEND AT CUST LEVEL IN PREVIOUS YEAR
                CASE WHEN cust_det.repcode = 'CL' AND   (NVL( SUM(financials.PY_ON_AMT) + SUM(financials.PY_PR_AMT) + SUM( financials.PY_OA_AMT),0) ) = 0  THEN 'Y' ELSE 'N' END AS excldclosedacct
        
        from     cust_det
        left join financials on financials.customer = cust_det.customer and financials.account = cust_det.account
        GROUP BY 
                cust_det.customer,
                cust_det.custname,
                cust_det.account,
                 cust_det.region,
                cust_det.accountname,
                cust_det.team,
                cust_det.repcode,
                cust_det.fullname,
                cust_det.segment,
                cust_det.city,
                cust_det.postcode,
                cust_det.country,
                cust_det.cust_status,
                 cust_det.law_customer_status,
                  cust_det.account_status
        )
            where excldclosedacct = 'N'
""".format(period)
df_base = pd.read_sql(query_base, con=conn_base)
output_base =  os.getcwd()+'\\Raw Data\\Raw_Financial Base.csv'
df_base.to_csv(output_base, index=False)






