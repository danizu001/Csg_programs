import pandas as pd
def call(wire_path,changes_path,selection,path):
    compare_df=[]
    if(selection=='0'):
        report_path=wire_path
        wc_export = pd.read_csv(report_path,sep='\t',dtype=str)
        wc_export = wc_export.assign(Concat1=wc_export['npa']+wc_export['nxx']+wc_export['from_line_num']+wc_export['to_line_num'])#Create another column with the concatenation of certain columns
        wc_export = wc_export.assign(Concat2=wc_export['clli'].str.slice(0, 4)+wc_export['ncta_state']+wc_export['clli'].str.slice(6))#Create another column of concatenation
        report_path=changes_path
        wc_changes = pd.read_excel(report_path, converters={'NPA':str,'NXX':str,'H_COORD':lambda x: str(x).zfill(5), 'V_COORD':lambda x: str(x).zfill(5), 'OCN':lambda x: str(x) if str(x)=='476' else str(x).zfill(4), 'LINE_FR':lambda x: str(x).zfill(4), 'LINE_TO':lambda x: str(x).zfill(4), 'Tandem_HCOORD':lambda x: str(x).zfill(5), 'Tandem_VCOORD':lambda x: str(x).zfill(5)})
        wc_changes["LINE_FR"]= wc_changes["LINE_FR"].str.zfill(4)
        wc_changes["LINE_TO"]= wc_changes["LINE_TO"].str.zfill(4)
        wc_changes = wc_changes.assign(Middle=wc_changes['Switch'].str.slice(4,6))#Create a column to take the middle of switch
        wc_changes = wc_changes.assign(Compare=wc_changes['LOC_STATE'] == wc_changes['Middle'])#Create a column for compare Loc_state and middle
        wc_changes = wc_changes.assign(Concat=wc_changes['NPA']+wc_changes['NXX']+wc_changes['LINE_FR']+wc_changes['LINE_TO'])#Create another column of concatenation
        for i in wc_changes['Concat']:
            try:
                index=wc_export['Concat1'].tolist().index(i)#Search in the first dataframe where is the data of the second dataframe
                value=wc_export['Concat2'][index]#after the searching we locate the value and put a data of the same row in a value
                compare_df.append(value)
            except:
                compare_df.append('#N/A')#if don't find anything we put n/a instead the value
        wc_changes = wc_changes.assign(Compare_df=compare_df)#add the last list to a column
        wc_changes = wc_changes.assign(Bool=wc_changes['Switch'].str.slice(0, 4)+wc_changes['LOC_STATE']+wc_changes['Switch'].str.slice(6)==wc_changes['Compare_df'])#add a column with true or false depend of the comparison of the concatenation with the last list
        wc_changes.loc[wc_changes['Compare_df'] == '#N/A', 'Bool'] = '#N/A'#if is N/a the values of the last column changes tu N/A too
        filter_wc = wc_changes[(wc_changes['Bool'] == False)]#filter and false
        filter_wc['Territory']=" "
        column_filter = filter_wc.columns.tolist()
        name_columns = column_filter[0:2]+column_filter[16:18]+column_filter[2:16]+[column_filter[21]]+column_filter[18:21]
        filter_wc = filter_wc[name_columns]#Put the columns in order that Jeff needs
        filter_wc = filter_wc.reset_index(drop=True)
        filter_wc.loc[filter_wc['LOC_STATE'] == 'XX' , 'Middle'] = 'XX'#if Loc state is XX middle convert to XX, this is because there are states with XX and for the comparison we don't want to mark them as false
        for i in range(len(filter_wc['LOC_STATE'])):
            if filter_wc['LOC_STATE'][i] == 'XX':
                filter_wc.loc[i,'Switch'] = filter_wc['Switch'][i][:4]+'XX'+filter_wc['Switch'][i][6:]#Put the XX in the first column in the middle
            if filter_wc['LOC_STATE'][i]!='XX' and filter_wc['Compare_df'][i]=='#N/A':#search if there are a value different to XX that match with N/A
                print('Warning with this value of Switch and Concat:' + filter_wc['Switch'][i] + ',' + filter_wc['Concat'][i]+' LOC_STATE is not XX and the Bool is #N/A')
        filter_wc = filter_wc.assign(last_column=filter_wc['Switch'].str.slice(0, 4)+filter_wc['LOC_STATE']+filter_wc['Switch'].str.slice(6))
        filter_wc['Switch'] = filter_wc['last_column']
        filter_wc['Compare'] = filter_wc['LOC_STATE'] == filter_wc['Middle']#Change the comparison with the new data of Middle
        filter_wc = filter_wc.assign(Bool2=filter_wc['Switch'].str.slice(4, 6))
        filter_wc['Middle'] = filter_wc['Bool2']
        del filter_wc['Bool2']
        filter_wc['Compare'] = filter_wc['LOC_STATE'] == filter_wc['Middle']#Change the comparison with the new data of Middle
        filter_wc.to_csv(path+'\Peerless CLLI Map Missmatches.txt', index=None, sep='\t')
    if(selection=='1'):
        report_path=wire_path
        wc_export = pd.read_csv(report_path,sep='\t',dtype=str)
        wc_export = wc_export.assign(Concat1=wc_export['npa']+wc_export['nxx']+wc_export['from_line_num']+wc_export['to_line_num'])
        wc_export = wc_export.assign(Concat2=wc_export['clli'].str.slice(0, 4)+wc_export['ncta_state']+wc_export['clli'].str.slice(6))
        report_path=changes_path
        wc_changes = pd.read_excel(report_path, converters={'NPA':str,'NXX':str,'H_COORD':lambda x: str(x).zfill(5), 'V_COORD':lambda x: str(x).zfill(5), 'OCN':lambda x: str(x) if str(x)=='476' else str(x).zfill(4), 'LINE_FR':lambda x: str(x).zfill(4), 'LINE_TO':lambda x: str(x).zfill(4), 'Tandem_HCOORD':lambda x: str(x).zfill(5), 'Tandem_VCOORD':lambda x: str(x).zfill(5)})
        wc_changes["LINE_FR"]= wc_changes["LINE_FR"].str.zfill(4)
        wc_changes["LINE_TO"]= wc_changes["LINE_TO"].str.zfill(4)
        wc_changes = wc_changes.assign(Middle=wc_changes['Switch'].str.slice(4,6))
        wc_changes = wc_changes.assign(Compare=wc_changes['LOC_STATE'] == wc_changes['Middle'])
        wc_changes = wc_changes.assign(Concat=wc_changes['NPA']+wc_changes['NXX']+wc_changes['LINE_FR']+wc_changes['LINE_TO'])
        for i in wc_changes['Concat']:
            try:
                index=wc_export['Concat1'].tolist().index(i)
                value=wc_export['Concat2'][index]
                compare_df.append(value)
            except:
                compare_df.append('#N/A')
        wc_changes = wc_changes.assign(Compare_df=compare_df)
        wc_changes = wc_changes.assign(Bool=wc_changes['Switch'].str.slice(0, 4)+wc_changes['LOC_STATE']+wc_changes['Switch'].str.slice(6)==wc_changes['Compare_df'])
        wc_changes.loc[wc_changes['Compare_df'] == '#N/A', 'Bool'] = '#N/A'
        if False in list(wc_changes['Bool']):
            print('There are mismatches at the moment')
        else:
            print('All is correct')