import openpyxl

wb = openpyxl.load_workbook(r'c:\Users\user\OneDrive\Desktop\geranium project\GERANIUM REFINEMNET\Task5\golden_solution\renewal_underwriting_review.xlsx')

ws_ar = wb['Account_Routing']
ws_es = wb['Executive_Summary']
ws_le = wb['Location_Exceptions']
ws_dn = wb['Decision_Notes']
ws_lr = wb['Loss_Reconciliation']

# Issue 3 fixes (making text live-linked via Excel formulas)

ws_ar['AD34'].value = '="Section 4 per granite_ridge_routing_notes.txt: unit count variance " & TEXT(ABS(G34-H34)/G34, "0.00%") & " (" & G34 & " list vs " & H34 & " schedule) exceeds 5% threshold; TIV variance " & TEXT(ABS(J34-K34)/J34, "0.00%") & " (" & TEXT(ABS(J34-K34), "$#,##0") & ") exceeds 3% threshold. Source: granite_ridge_renewal_accounts.csv, granite_ridge_location_schedule.csv."'

ws_ar['AD35'].value = '="Section 3: recomputed LR " & TEXT(V35, "0.0%") & " is 45%-70% (granite_ridge_loss_run_36mo.csv); Section 3: HAB-2162-C attritional incurred " & TEXT(W35, "$#,##0") & " >= $250,000 (granite_ridge_loss_run_36mo.csv)."'

ws_ar['AD36'].value = '="Section 3 decline: recomputed attritional loss ratio " & TEXT(V36, "0.0%") & " (>=70%) from granite_ridge_loss_run_36mo.csv; account list showed " & TEXT(U36, "0.0%") & " (stale) in granite_ridge_renewal_accounts.csv."'

ws_es['A41'].value = '="The October batch has 36 accounts, all routed. " & TEXT(C16, "0") & " are cleared to quote; " & TEXT(C17, "0") & " require underwriter review (Refer); GR-2195 (Humboldt Yard) hits a " & TEXT(Account_Routing!V36, "0.0%") & " recomputed loss ratio and is declined; GR-2033 (Lincolnwood Place) needs broker vacancy confirmation before we can proceed. The account list flagged every account renewal-ready, but that is informational per Section 8. Judgment calls on Campus Edge (Section 10), GR-2047 vacancy, and GR-2071 cat exclusion are documented on Decision_Notes."'

ws_le['D13'].value = '="List " & Account_Routing!G34 & " vs schedule sum " & Account_Routing!H34 & " (variance " & ABS(Account_Routing!G34-Account_Routing!H34) & " units, " & TEXT(ABS(Account_Routing!G34-Account_Routing!H34)/Account_Routing!G34, "0.00%") & ")"'

ws_le['D14'].value = '="List " & TEXT(Account_Routing!J34, "$#,##0") & " vs schedule sum " & TEXT(Account_Routing!K34, "$#,##0") & " (variance " & TEXT(ABS(Account_Routing!J34-Account_Routing!K34), "$#,##0") & ", " & TEXT(ABS(Account_Routing!J34-Account_Routing!K34)/Account_Routing!J34, "0.00%") & ")"'

# Decision_Notes fixes C2, C3, C7:C9, C13
ws_dn['C2'].value = '="Thirty of the 36 accounts show list loss ratios differing by more than five points from the recomputed 36-month ratios. Recomputations are detailed on Loss_Reconciliation."'
ws_dn['C3'].value = '="Four loss rows total " & TEXT(SUMIFS(Src_LossRun!$I$2:$I$43,Src_LossRun!$C$2:$C$43,"GR-2195"), "$#,##0") & " incurred against " & TEXT(Account_Routing!D36, "$#,##0") & " premium = " & TEXT(Account_Routing!V36, "0.0%") & ". Account list showed " & TEXT(Account_Routing!U36, "0.0%") & "."'

wb.save(r"c:\Users\user\OneDrive\Desktop\geranium project\GERANIUM REFINEMNET\Task5\golden_solution\renewal_underwriting_review.xlsx")
print("Workbook updated with formulas.")
