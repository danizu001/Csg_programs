from openpyxl import comments                             
import reports
import database 
import os

def selectronics(part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.usage_by_day_report('5D')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.mou_entries()
        reports.exception_analisys()
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.zip_and_error_report('SELECTRONICS', 'ERROR')
        reports.zip_and_error_report('SELECTRONICS', 'ZIP')
        reports.usage_balancing_report('6')
        reports.payment_report()

        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
                                
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion() 
        reports.switched_bans_trending_report("12M")
        reports.facility_bans_trending_report("12N")
        reports.adjustment_report()
        reports.occ_report()
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.aged_trial_balance_for_export_report()
        reports.mmr_by_BAN_report()
        reports.transaction_summary_report()
        reports.late_payment_charges_report()
        reports.revenue_analysis_report()
        reports.bill_by_period_report()
        reports.facility_circuit_charges_billed_disconnected()
        reports.invoice_balance_report()
        reports.number_of_invoices()
        reports.Switch_Terminating_Intrastate_Rev_Report_for_NECA()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def bendtel(part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.usage_by_day_report('5D') 
        reports.usage_by_day_report('5E-3MONTHS')
        reports.zip_and_error_report('BENDTEL', 'ERROR')
        reports.zip_and_error_report('BENDTEL', 'ZIP')
        reports.mou_entries()
        reports.exception_analisys()
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.usage_balancing_report('6')
        reports.payment_report()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
                               
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()  
        reports.switched_bans_trending_report("12C")       
        reports.adjustment_report()
        reports.occ_report()
        reports.usage_balancing_report('13')
        reports.pre_deletions()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST2':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.pos_deletions()
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.mmr_by_BAN_report()
        reports.transaction_summary_report()
        reports.bill_by_period_report()
        reports.switched_bans_trending_report("19W")
        reports.invoice_balance_report()
        reports.adjustment_occ_report()
        reports.number_of_invoices()

        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def mieac(part):
    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.mou_entries()
        reports.exception_analisys()
        reports.adjs_not_posted()
        reports.usage_by_day_report('5D')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.zip_and_error_report('MIEAC', 'ERROR')
        reports.zip_and_error_report('MIEAC', 'ZIP')
        reports.usage_balancing_report('6')
        reports.payment_report()

        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
                               
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()  
        reports.switched_bans_trending_report("12B")       
        reports.adjustment_report()
        reports.occ_report()
        reports.usage_balancing_report('13')
        reports.mmr_by_BAN_report("13C")
        reports.aged_trial_balance_for_export_report('13B')
        reports.transaction_summary_report('13D')
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST2':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.aged_trial_balance_for_export_report()
        reports.msg_mou_rev_with_lpc()
        reports.msg_mou_rev_export()
        reports.mmr_by_BAN_report()
        reports.mmr_by_clli_report()
        reports.transaction_summary_report()
        reports.revenue_analysis_report()
        reports.billing_review()
        reports.bill_by_period_report()
        reports.switched_bans_trending_report('19W')
        reports.invoice_balance_report()
        reports.number_of_invoices()
        reports.aurs_bill_date('19AX')
        reports.aurs_bill_date('19AY')                              
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")
        
        
def onvoy(part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.mou_entries()
        reports.exception_analisys()
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.usage_by_day_report('5D')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.zip_and_error_report('Onvoy', 'ZIP')
        reports.zip_and_error_report('Onvoy', 'ERROR')
        reports.usage_balancing_report('6')
        reports.payment_report()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


    if part == 'POST1':
                               
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()  
        reports.switched_bans_trending_report("12B")       
        reports.adjustment_report()
        reports.occ_report()
        reports.usage_balancing_report('13')
        reports.mmr_by_BAN_report("13C")
        reports.aged_trial_balance_for_export_report('13B')
        reports.transaction_summary_report('13D')
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST2':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.aged_trial_balance_for_export_report()
        reports.msg_mou_rev_with_lpc()
        reports.msg_mou_rev_export()
        reports.mmr_by_BAN_report()
        reports.mmr_by_clli_report()
        reports.transaction_summary_report()
        reports.revenue_analysis_report()
        reports.billing_review()
        reports.bill_by_period_report()
        reports.switched_bans_trending_report('19W')
        reports.invoice_balance_report()
        reports.number_of_invoices()
        reports.aurs_bill_date('19AX')
        reports.aurs_bill_date('19AY')
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def mta_sw(part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.mou_entries()
        reports.exception_analisys()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.usage_by_day_report('5D')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.zip_and_error_report('MTA', 'ERROR', state_billtype= 'SW')
        reports.zip_and_error_report('MTA', 'ZIP', state_billtype= 'SW')
        reports.usage_balancing_report('6',bill_type='SW')
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")
        

    if part == 'POST1':
                              
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()   
        reports.adjs_not_posted()
        reports.payment_report(fa_sw_rc='SW')
        reports.switched_bans_trending_report()
        reports.adjustment_report()
        reports.occ_report(bill_type='SW')
        reports.usage_balancing_report('13',bill_type='SW')
        reports.aged_trial_balance_for_export_report(bill_type='SW')
        reports.mmr_by_BAN_report()
        reports.mmr_by_clli_report()
        reports.mmr_by_rate_element()
        reports.transaction_summary_report(bill_type='SW')
        reports.bill_by_period_report()
        reports.switched_usage_summary_charges()
        reports.revenue_analysis_report(bill_type = 'SW')
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def mta_fa (part):

    if part == 'POST1':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.payment_report(fa_sw_rc='FA')
        reports.facility_bans_trending_report()
        reports.adjustment_report()
        reports.facility_summary_charges()
        reports.facility_charges_by_cic_clli()
        reports.facility_circuit_charges_billed_disconnected(code="19AC",disconnect='Y')
        reports.occ_report(bill_type='FA')
        reports.transaction_summary_report(bill_type='FA')
        reports.fusc_charge_by_circuit()
        reports.aged_trial_balance_for_export_report(bill_type='FA')
        reports.revenue_analysis_report(bill_type = 'FA')
        reports.occ_billed()
        reports.bill_completion()
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def mta_rc (part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.mou_entries()
        reports.exception_analisys()
        reports.usage_by_day_report('5D')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.usage_balancing_report('6',bill_type='RC')
        reports.zip_and_error_report('MTA', 'ERROR', state_billtype= 'RC')
        reports.zip_and_error_report('MTA', 'ZIP', state_billtype= 'RC')
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')


        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")
    

    if part == 'POST1':
                              
        
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()   
        reports.adjs_not_posted()
        reports.payment_report(fa_sw_rc='RC')
        reports.switched_bans_trending_report()
        reports.adjustment_report()
        reports.occ_report()
        reports.usage_balancing_report('13',bill_type='RC')
        reports.aged_trial_balance_for_export_report(bill_type='RC')
        reports.mmr_by_BAN_report()
        reports.mmr_by_clli_report()
        reports.mmr_by_rate_element()
        reports.transaction_summary_report(bill_type='RC')
        reports.bill_by_period_report()
        reports.switched_usage_summary_charges()
        reports.revenue_analysis_report(bill_type = 'RC')
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def nt_sw (part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.mou_entries()
        reports.exception_analisys()
        reports.adjs_not_posted()
        reports.usage_by_day_report('5D','OSA')
        reports.usage_by_day_report('5D','TTS')
        reports.usage_by_day_report('5E-3MONTHS','OSA')
        reports.usage_by_day_report('5E-3MONTHS','TTS')
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.zip_and_error_report('Neutral_Tandem', 'ZIP','SW')
        reports.zip_and_error_report('Neutral_Tandem', 'ERROR','SW')
        reports.usage_balancing_report('6')
        reports.payment_report(option='OSA')
        reports.payment_report(option='TTS')
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()
        reports.usage_balancing_report('13')
        reports.switched_bans_trending_report('12B',option='OSA')
        reports.switched_bans_trending_report('12B',option='TTS')
        reports.adjustment_report(option='OSA')
        reports.adjustment_report(option='TTS')
        reports.occ_report(bill_type='SW',option='OSA')
        reports.occ_report(bill_type='SW',option='TTS')
        reports.aged_trial_balance_report(option='OSA')
        reports.aged_trial_balance_report(option='TTS')
        reports.aged_trial_balance_for_export_report(option='OSA')
        reports.aged_trial_balance_for_export_report(option='TTS')
        reports.msg_mou_rev_export(option='OSA')
        reports.msg_mou_rev_export(option='TTS')
        reports.mmr_by_BAN_report(option='OSA')
        reports.mmr_by_BAN_report(option='TTS')
        reports.mmr_by_clli_report(option='OSA')
        reports.mmr_by_clli_report(option='TTS')
        reports.transaction_summary_report(option='OSA',client='NT')
        reports.transaction_summary_report(option='TTS',client="NT")
        reports.bill_by_period_report(option='OSA')
        reports.bill_by_period_report(option='TTS')
        reports.revenue_analysis_report(option='OSA')
        reports.revenue_analysis_report(option='TTS')
        reports.number_of_invoices(option='OSA')
        reports.number_of_invoices(option='TTS')
        reports.msg_mou_rev_with_lpc(code="19Q", option ='OSA')
        reports.msg_mou_rev_with_lpc(code="19Q", option ='TTS')
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def nt_fa (part):
    
    if part == 'POST1':
        
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.payment_report()
        reports.facility_bans_trending_report(bill_type='NTFA')
        reports.adjustment_report()
        reports.occ_report()
        reports.aged_trial_balance_for_export_report()
        reports.msg_mou_rev_export(bill_type='FA')
        reports.transaction_summary_report()
        reports.facility_summary_charges()
        reports.facility_circuit_charges_billed_disconnected(code='19AC')
        reports.facility_charges_by_cic_clli()
        reports.accounting_detailed_report()
        reports.accounting_report()
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def gci_sw (part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.exception_analisys()
        reports.mou_entries()
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.usage_by_day_report('5D')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.payment_report(fa_sw_rc='SW')
        reports.zip_and_error_report('GCI', 'ERROR',state_billtype= 'SW')
        reports.zip_and_error_report('GCI', 'ZIP',state_billtype= 'SW')
        reports.usage_balancing_report('6')
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")
        

    if part == 'POST1':
                                 
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()
        reports.payment_report()
        reports.switched_bans_trending_report(company='gci')
        reports.adjustment_report()
        reports.occ_report(bill_type='SW')
        reports.usage_balancing_report('13')
        reports.aged_trial_balance_for_export_report()
        reports.mmr_by_BAN_report()
        reports.mmr_by_clli_report()
        reports.mmr_by_rate_element()
        reports.transaction_summary_report(bill_type='SW')
        reports.bill_by_period_report()
        reports.switched_usage_summary_charges()
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def gci_fa (part):

    if part == 'POST1':
        
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')
        reports.adjs_not_posted()
        reports.payment_report()
        reports.facility_bans_trending_report(company='gci')
        reports.adjustment_report()
        reports.facility_summary_charges()
        reports.facility_charges_by_cic_clli()
        reports.facility_circuit_charges_billed_disconnected(code="19AC",disconnect='Y')
        reports.occ_report()
        reports.transaction_summary_report(bill_type='FA')
        reports.fusc_charge_by_circuit()
        reports.aged_trial_balance_for_export_report()
        reports.occ_billed()
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def amb_356DMI(part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        # reports.ban_and_clli_error_report('4F')
        # reports.ban_and_clli_error_report('4H')
        # reports.usage_by_day_report('5D')
        reports.zip_and_error_report('AMERICANBROADBAND', 'ERROR','MI_356D')
        reports.zip_and_error_report('AMERICANBROADBAND', 'ZIP','MI_356D')
        # reports.exception_analisys()
        # reports.mou_entries()
        # reports.adjs_not_posted()
        # reports.soc_factor_change('10')
        # reports.soc_factor_change('11')                                               
        # reports.usage_balancing_report('6')
        # reports.payment_report()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
                               
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.bill_completion()  
        reports.switched_bans_trending_report('12C')
        # reports.adjustment_occ_report()
        reports.usage_balancing_report('13')
        reports.pre_deletions()
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST2':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.pos_deletions()
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.aged_trial_balance_for_export_report()
        reports.mmr_by_BAN_report()
        reports.transaction_summary_report()
        reports.bill_by_period_report()
        reports.switched_bans_trending_report('19W')
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def amb_509BOH(part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        # reports.ban_and_clli_error_report('4F')
        # reports.ban_and_clli_error_report('4H')
        # reports.usage_by_day_report('5D')
        reports.zip_and_error_report('AMERICANBROADBAND', 'ERROR','OH_509B')
        reports.zip_and_error_report('AMERICANBROADBAND', 'ZIP','OH_509B')
        # reports.exception_analisys()
        # reports.mou_entries()
        # reports.adjs_not_posted()
        # reports.soc_factor_change('10')
        # reports.soc_factor_change('11')                                                           
        # reports.usage_balancing_report('6')
        # reports.payment_report()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
                                 
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.bill_completion()
        reports.switched_bans_trending_report('12C')
        #reports.adjustment_occ_report()
        reports.usage_balancing_report('13')
        reports.pre_deletions()

        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST2':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.pos_deletions()
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.aged_trial_balance_for_export_report()
        reports.mmr_by_BAN_report()
        reports.transaction_summary_report()
        reports.bill_by_period_report()
        reports.switched_bans_trending_report('19W')
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def amb_590GIN(part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4F')
        reports.ban_and_clli_error_report('4H')
        reports.usage_by_day_report('5D')
        reports.zip_and_error_report('AMERICANBROADBAND', 'ERROR','IN_590G')
        reports.zip_and_error_report('AMERICANBROADBAND', 'ZIP','IN_590G')
        reports.exception_analisys()
        reports.mou_entries()
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')                                                         
        reports.usage_balancing_report('6')
        reports.payment_report()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        
        reports.usage_by_day_report('5E-3MONTHS')
        reports.bill_completion()                         
        reports.switched_bans_trending_report('12C')
        #reports.adjustment_occ_report()
        reports.usage_balancing_report('13')
        reports.pre_deletions()

        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST2':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.pos_deletions()
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.aged_trial_balance_for_export_report()
        reports.mmr_by_BAN_report()
        reports.transaction_summary_report()
        reports.bill_by_period_report()
        reports.switched_bans_trending_report('19W')
        reports.number_of_invoices()
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


def peerless (part):

    if part == 'PRE':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.ban_and_clli_error_report('4H')
        reports.usage_by_day_report('5E-3MONTHS')
        reports.usage_balancing_report('6')
        reports.payment_report()
        reports.mou_entries()
        reports.exception_analisys()
        reports.zip_and_error_report('Peerless_Network', 'ZIP')
        reports.zip_and_error_report('Peerless_Network', 'ERROR')
        reports.adjs_not_posted()
        reports.soc_factor_change('10')
        reports.soc_factor_change('11')

        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

    if part == 'POST1':
                               
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.bill_completion()  
        reports.active_bans()
        reports.switched_bans_trending_report('12C')
        #Report 12F we can do a function to count values when they are equal to Y
        reports.revenue_analysis_report(code = '12H')
        reports.usage_balancing_report('13')
        reports.aged_trial_balance_for_export_report(code='13B')
        reports.mmr_by_BAN_report(code='13C')
        reports.transaction_summary_report(code='13D',client='peerless')
        reports.adjustment_report(code='13E')
        reports.occ_report(code='13F')

        reports.macros('PEER_SW_BAN_Trending_12C')
        print("SWBT is formated")
        reports.macros('PEER_Rev_Analysis_12H')
        print("Rev Analysis is formated")
        reports.macros('PEER_Adj_Formatting')
        print("Adjs is formated")
        reports.macros('PEER_OCC_Formatting')
        print("OCCs is formated")

        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")


    if part == 'POST2':
        database.log_file.write('========================================================================= Reports =========================================================================\n')
        reports.usage_balancing_report('16')
        reports.aged_trial_balance_report()
        reports.aged_trial_balance_for_export_report()
        reports.msg_mou_rev_export()
        reports.mmr_by_BAN_report()
        reports.transaction_summary_report(client='peerless')
        reports.rev_analysis7()
        reports.revenue_analysis_report()
        reports.bill_by_period_report()
        reports.switched_bans_trending_report('19W')
        reports.invoice_balance_report()
        reports.mmr_bill_date()
        reports.gl_extract()
        reports.occ_report(code='19AS')
        reports.soc_jurisdiction()
        reports.rev_by_clli()
        reports.number_of_invoices()
        reports.macros('PEER_MMR_19F')
        # reports.macros('PEER_MMR_Billed_Date') #need to know if date format can be ignored in order to delete the macro from here
        reports.macros('PEER_Rev_Analysis_19S')
        reports.macros('PEER_Billed_by_Period')
        # reports.macros('PEER_GL_Extract') #need to know if date format can be ignored in order to delete the macro from here
        reports.macros('PEER_Inv_Bal_Juris')
        reports.macros('PEER_Threshold_by_Pd')
        reports.macros('PEER_SOC_Juris')
        reports.macros('PEER_PT_Export')
        
        
        # saving logs to logs.txt
        database.log_file.write('\n')
        database.log_file.flush()
        os.fsync(database.log_file.fileno())
        print("Finished")

