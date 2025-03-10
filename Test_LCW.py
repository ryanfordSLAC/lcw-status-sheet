import oracledb
import pandas as pd
import xlsxwriter

import smtplib
# MIMEMultipart send emails with both text content and attachments.
from email.mime.multipart import MIMEMultipart
# MIMEText for creating body of the email message.
from email.mime.text import MIMEText
# MIMEApplication attaching application-specific data (like CSV files) to email messages.
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders


pd.set_option('display.max_colwidth', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.options.display.float_format = '{:.2e}'.format 

class return_LCW:
	def return_info():

		odts_username = 'LCWSTATUS'
		odts_password = 'mkVEbrfsJW2#36%T'
		odts_dsn = 'slacprod.slac.stanford.edu/slacprod'
		connection = oracledb.connect (
			user=odts_username,
			password=odts_password,
			dsn=odts_dsn)

		if connection.is_healthy():
				from pandas import DataFrame
				print("Connection is Healthy")
				cursor = connection.cursor()
				cursor2 = connection.cursor()
				querystring0 = "select * from RAD_PROTECT_ADMIN.VW_LCWREPORT"
				querystring1 = ("SELECT SECAREA_DISPLAY, SYSTEMNAME_DISPLAY,DISPLAY_ORDER, COPPER_OR_SS, NOTES, \
					SOURCE_WATER, COLLECTION_TANK, ACCELSECT_COMPCOOLED \
					FROM RAD_PROTECT_ADMIN.LCWSTAT_REPORT WHERE IS_ACTIVE = 'Y' AND DISPLAY_ORDER > 0 \
					ORDER BY DISPLAY_ORDER")
				querystring2 = "select * from RAD_PROTECT_ADMIN.VW_MAX_LCW WHERE CONCENTRATION > 0"
					#FROM RAD_PROTECT_ADMIN.LCWSTAT_REPORT WHERE TABLE_NO = '1' AND IS_ACTIVE = 'Y' AND DISPLAY_ORDER > 0 \
				query = cursor.execute(querystring2)
				df = DataFrame(query)
				# Max concentration and the date it was taken on
				df1 = df[[0,1,2,11,21,23]]
				df1a = df1.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 21:'Type', 23:'Conc'})
				df2 = df1a.loc[df1a.groupby(['ID','Loc','SubLoc','Type'])['Conc'].transform('max') == df1a['Conc']]
				max_conc0 = df2.groupby(['ID','Loc','SubLoc','Type']).agg(date_of_max_conc = ('Date', 'max'), max_conc = ('Conc', 'max'))
				max_conc0['date_of_max_conc'] = pd.to_datetime(max_conc0['date_of_max_conc']).dt.strftime('%m-%d-%Y')
				max_conc1 = max_conc0.replace({pd.NaT: None})

				#print(df1.columns)
				#print(max_conc1.columns)
				#print("max conc1")
				#print(max_conc1)

				#Latest sample date with the concentration
				query2 = cursor.execute(querystring0)
				df3 = DataFrame(query2)
				df4 = df3[[0,1,2,11,23]]
				df4a = df4.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 23:'Conc'})
				df5 = df4a.loc[df4a.groupby(['ID','Loc','SubLoc'])['Date'].transform('max') == df4a['Date']]
				df6 = df5.groupby(['ID','Loc','SubLoc']).agg(date_of_latest_conc = ('Date', 'max'), latest_conc = ('Conc', 'max'))
				df6['date_of_latest_conc'] = pd.to_datetime(df6['date_of_latest_conc']).dt.strftime('%m-%d-%Y')
				#print("df6")
				#print(df6)

				query3 = cursor2.execute(querystring1)
				df8 = DataFrame(query3)
				df9 = df8.rename(columns={0:'Loc', 1:'SubLoc', 2:'ID', 3:'CU, SS, AL', 4:'Discharge Notes', 5:'Source of Water', 6:'Collection Tank', 7:'Accelerator Sectors and/or Components Cooled'})
				#print('df9')
				#print(df9)

				combined0 = df6.merge(max_conc1, how='left', on=['ID', 'Loc', 'SubLoc'])
				combined = combined0.sort_values(by='Loc')
				#print('combined')
				#print(combined)

				combined1 = combined.merge(df9, how='left', on=['Loc', 'SubLoc'], indicator=True)
				combined2 = combined1.sort_values(by='Loc')

				#print('combined2')
				#print(combined2)
			
				table1 = combined2[combined2['Loc'].isin(['Positron Vault', 'BSY Sumps','BSY Collimator Sumps', 'BDE Sump', 'North Research Yard', 'BTH East'])]
				table1_sorted = table1.sort_values(by='ID')
				table1_ordered = table1_sorted[['Loc', 'Source of Water', 'Collection Tank', 'max_conc', 'date_of_max_conc','latest_conc','date_of_latest_conc','Discharge Notes']]
				table1_noindex = table1_ordered.reset_index(drop=True)
				table1_renamed = table1_noindex.rename(columns={'Loc': 'Location', 'max_conc':'Historical Maximum Tritium Concentrations (pCi/L)', 'date_of_max_conc':'Sample Taken', 'latest_conc':'Latest Tritium Concentrations (pCi/L)', 'date_of_latest_conc':'Latest Sample Taken'})
			

			##################################################### Table 2a

	
				cursor2_2 = connection.cursor()

				query_2 = cursor2_2.execute(querystring2)
				df_2 = DataFrame(query_2)
				# Max concentration and the date it was taken on
				df1_2 = df_2[[0,1,2,11,21,23]]
				df1a_2 = df1_2.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 21:'Type', 23:'Conc'})
				df2_2 = df1a_2.loc[df1a.groupby(['ID','Loc','SubLoc','Type'])['Conc'].transform('max') == df1a_2['Conc']]
				max_conc0_2 = df2_2.groupby(['ID','Loc','SubLoc','Type']).agg(date_of_max_conc = ('Date', 'max'), max_conc = ('Conc', 'max'))
				max_conc0_2['date_of_max_conc'] = pd.to_datetime(max_conc0_2['date_of_max_conc']).dt.strftime('%m-%d-%Y')
				max_conc1_2 = max_conc0_2.replace({pd.NaT: None})

				#print(df1.columns)
				#print(max_conc1.columns)
				#print("max conc1")
				#print(max_conc1)

				#Latest sample date with the concentration
				query2_2 = cursor.execute(querystring0)
				df3_2 = DataFrame(query2_2)
				df4_2 = df3_2[[0,1,2,11,23]]
				df4a_2 = df4_2.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 23:'Conc'})
				df5_2 = df4a_2.loc[df4a.groupby(['ID','Loc','SubLoc'])['Date'].transform('max') == df4a_2['Date']]
				df6_2 = df5_2.groupby(['ID','Loc','SubLoc']).agg(date_of_latest_conc = ('Date', 'max'), latest_conc = ('Conc', 'max'))
				df6_2['date_of_latest_conc'] = pd.to_datetime(df6_2['date_of_latest_conc']).dt.strftime('%m-%d-%Y')
				#print("df6")
				#print(df6)

				query3_2 = cursor2.execute(querystring1)
				df8_2 = DataFrame(query3_2)
				df9_2 = df8_2.rename(columns={0:'Loc', 1:'SubLoc', 2:'ID', 3:'CU, SS, AL', 4:'Discharge Notes', 5:'Source of Water', 6:'Collection Tank', 7:'Accelerator Sectors and/or Components Cooled'})
				#print('df9')
				#print(df9)

				combined0_2 = df6_2.merge(max_conc1_2, how='left', on=['ID', 'Loc', 'SubLoc'])
				combined_2 = combined0_2.sort_values(by='Loc')
				#print('combined')
				#print(combined)

				combined1_2 = combined_2.merge(df9_2, how='left', on=['Loc', 'SubLoc'], indicator=True)
				combined2_2 = combined1_2.sort_values(by='Loc')

				#print('combined2')
				#print(combined2)
			
				table1_2 = combined2_2[combined2_2['Loc'].isin(['Sec 3', 'Sec 6', 'Sec 9', 'Sec 11-08', 'Sec 12', 'Sec 15', 'Sec 18', 'Sec 21', 'Sec 24', 'Sec 27', 'Sec 30', 'SSRL', 'RAMSY', 'IR 8'])]
				table1_sorted_2 = table1_2.sort_values(by='ID')
				table1_ordered_2 = table1_sorted_2[['Loc', 'Source of Water', 'Collection Tank', 'max_conc', 'date_of_max_conc','latest_conc','date_of_latest_conc','Discharge Notes']]
				table1_noindex_2 = table1_ordered_2.reset_index(drop=True)
				table1_renamed_2 = table1_noindex_2.rename(columns={'Loc': 'Location', 'max_conc':'Historical Maximum Tritium Concentrations (pCi/L)', 'date_of_max_conc':'Sample Taken', 'latest_conc':'Latest Tritium Concentrations (pCi/L)', 'date_of_latest_conc':'Latest Sample Taken'})
			


			##################################################### End Table 2a

			##################################################### Begin Table 2b

				cursor2_3 = connection.cursor()

				query_3 = cursor2_3.execute(querystring2)
				df_3 = DataFrame(query_3)
				# Max concentration and the date it was taken on
				df1_3 = df_3[[0,1,2,11,21,23]]
				df1a_3 = df1_3.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 21:'Type', 23:'Conc'})
				df2_3 = df1a_3.loc[df1a.groupby(['ID','Loc','SubLoc','Type'])['Conc'].transform('max') == df1a_3['Conc']]
				max_conc0_3 = df2_3.groupby(['ID','Loc','SubLoc','Type']).agg(date_of_max_conc = ('Date', 'max'), max_conc = ('Conc', 'max'))
				max_conc0_3['date_of_max_conc'] = pd.to_datetime(max_conc0_3['date_of_max_conc']).dt.strftime('%m-%d-%Y')
				max_conc1_3 = max_conc0_3.replace({pd.NaT: None})

				#print(df1.columns)
				#print(max_conc1.columns)
				#print("max conc1")
				#print(max_conc1)

				#Latest sample date with the concentration
				query2_3 = cursor.execute(querystring0)
				df3_3 = DataFrame(query2_3)
				df4_3 = df3_3[[0,1,2,11,23]]
				df4a_3 = df4_3.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 23:'Conc'})
				df5_3 = df4a_3.loc[df4a.groupby(['ID','Loc','SubLoc'])['Date'].transform('max') == df4a_3['Date']]
				df6_3 = df5_3.groupby(['ID','Loc','SubLoc']).agg(date_of_latest_conc = ('Date', 'max'), latest_conc = ('Conc', 'max'))
				df6_3['date_of_latest_conc'] = pd.to_datetime(df6_3['date_of_latest_conc']).dt.strftime('%m-%d-%Y')
				#print("df6")
				#print(df6)

				query3_3 = cursor2.execute(querystring1)
				df8_3 = DataFrame(query3_3)
				df9_3 = df8_3.rename(columns={0:'Loc', 1:'SubLoc', 2:'ID', 3:'CU, SS, AL', 4:'Discharge Notes', 5:'Source of Water', 6:'Collection Tank', 7:'Accelerator Sectors and/or Components Cooled'})
				#print('df9')
				#print(df9)

				combined0_3 = df6_3.merge(max_conc1_3, how='left', on=['ID', 'Loc', 'SubLoc'])
				combined_3 = combined0_3.sort_values(by='Loc')
				#print('combined')
				#print(combined)

				combined1_3 = combined_3.merge(df9_3, how='left', on=['Loc', 'SubLoc'], indicator=True)
				combined2_3 = combined1_3.sort_values(by='Loc')

				#print('combined2')
				#print(combined2)
			
				table1_3 = combined2_3[combined2_3['Loc'].isin(['SLC North Arc', 'SLC South Arc', 'IR 2', 'IR 4', 'IR 12'])]
				table1_sorted_3 = table1_3.sort_values(by='ID')
				table1_ordered_3 = table1_sorted_3[['Loc', 'Source of Water', 'Collection Tank', 'max_conc', 'date_of_max_conc','latest_conc','date_of_latest_conc','Discharge Notes']]
				table1_noindex_3 = table1_ordered_3.reset_index(drop=True)
				table1_renamed_3 = table1_noindex_3.rename(columns={'Loc': 'Location', 'max_conc':'Historical Maximum Tritium Concentrations (pCi/L)', 'date_of_max_conc':'Sample Taken', 'latest_conc':'Latest Tritium Concentrations (pCi/L)', 'date_of_latest_conc':'Latest Sample Taken'})
			

			##################################################### End Table 2b


			##################################################### Begin Table 3

				cursor2_4 = connection.cursor()

				query_4 = cursor2_4.execute(querystring2)
				df_4 = DataFrame(query_4)
				# Max concentration and the date it was taken on
				df1_4 = df_4[[0,1,2,11,21,23]]
				df1a_4 = df1_4.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 21:'Type', 23:'Conc'})
				df2_4 = df1a_4.loc[df1a.groupby(['ID','Loc','SubLoc','Type'])['Conc'].transform('max') == df1a_4['Conc']]
				max_conc0_4 = df2_4.groupby(['ID','Loc','SubLoc','Type']).agg(date_of_max_conc = ('Date', 'max'), max_conc = ('Conc', 'max'))
				max_conc0_4['date_of_max_conc'] = pd.to_datetime(max_conc0_4['date_of_max_conc']).dt.strftime('%m-%d-%Y')
				max_conc1_4 = max_conc0_4.replace({pd.NaT: None})


				#Latest sample date with the concentration
				query2_4 = cursor.execute(querystring0)
				df3_4 = DataFrame(query2_4)
				df4_4 = df3_4[[0,1,2,11,23]]
				df4a_4 = df4_4.rename(columns={0:'ID', 1:'Loc', 2:'SubLoc', 11:'Date', 23:'Conc'})
				df5_4 = df4a_4.loc[df4a.groupby(['ID','Loc','SubLoc'])['Date'].transform('max') == df4a_4['Date']]
				df6_4 = df5_4.groupby(['ID','Loc','SubLoc']).agg(date_of_latest_conc = ('Date', 'max'), latest_conc = ('Conc', 'max'))
				df6_4['date_of_latest_conc'] = pd.to_datetime(df6_4['date_of_latest_conc']).dt.strftime('%m-%d-%Y')
				#print("df6")
				#print(df6)

				query3_4 = cursor2.execute(querystring1)
				df8_4 = DataFrame(query3_4)
				df9_4 = df8_4.rename(columns={0:'Loc', 1:'SubLoc', 2:'ID', 3:'CU, SS, AL', 4:'System Notes', 5:'Source of Water', 6:'Collection Tank', 7:'Accelerator Sectors and/or Components Cooled'})
				#print('df9')
				#print(df9)

				combined0_4 = df6_4.merge(max_conc1_4, how='left', on=['ID', 'Loc', 'SubLoc'])
				combined_4 = combined0_4.sort_values(by='Loc')
				#print('combined')
				#print(combined)

				combined1_4 = combined_4.merge(df9_4, how='left', on=['Loc', 'SubLoc'], indicator=True)
				combined2_4 = combined1_4.sort_values(by='Loc')

				#print('combined2')
				#print(combined2)
			
				table1_4 = combined2_4[combined2_4['Loc'].isin(['2', '5', '9', '10', '17', '18', '19', '20', '21', '24', '28', '29', '30', 'R. Yard', '1801', 'EBD', 'BSY'])]
				table1_sorted_4 = table1_4.sort_values(by='ID')
				table1_ordered_4 = table1_sorted_4[['Loc', 'SubLoc','Accelerator Sectors and/or Components Cooled', 'max_conc', 'date_of_max_conc','latest_conc','date_of_latest_conc','System Notes', 'CU, SS, AL']]
				table1_noindex_4 = table1_ordered_4.reset_index(drop=True)
				table1_renamed_4 = table1_noindex_4.rename(columns={'Loc': 'Sector/Area','SubLoc':'System', 'max_conc':'Historical Maximum Tritium Concentrations (pCi/L)', 'date_of_max_conc':'Sample Taken', 'latest_conc':'Latest Tritium Concentrations (pCi/L)', 'date_of_latest_conc':'Latest Sample Taken'})
			

			##################################################### End Table 3


				filename = "C:\\Users\\ryanford\\OneDrive - SLAC National Accelerator Laboratory\\2025\\LCW.xlsx" #test
				#filename = "C:\\LCW\\LCW.xlsx" #prod
				with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
					#df.to_excel(writer, sheet_name='Raw Data', index=False)
					table1_renamed.to_excel(writer, sheet_name='Table 1', index=False, startrow=1)
					workbook = writer.book
					worksheet = writer.sheets['Table 1']
					cell_format1 = workbook.add_format({'text_wrap': True, 'bottom': 2, 'top': 2, 'left': 2, 'right': 2})
					cell_wrap = workbook.add_format({'text_wrap': True})
					scientific_format = workbook.add_format({'num_format': '0.00E+00'})
					header_format = workbook.add_format({'text_wrap': True, 'bottom': 2, 'top': 2, 'left': 2, 'right': 2, 'bold': True, 'fg_color': '#93ff33'})
					table_name_format = workbook.add_format({'align': "center", 'valign':'vcenter', 'bold': True, 'font_size': 14, 'text_wrap': True})			
					worksheet.conditional_format('A2:H8', {'type': 'no_errors', 'format': cell_format1})
					worksheet.set_column('A:H', 20, cell_wrap)
					worksheet.set_column('D:D', 20, scientific_format)					
					worksheet.set_column('F:F', 20, scientific_format)
					worksheet.set_row(0, 40, table_name_format)

					worksheet.merge_range('A1:H1', 'Table 1: The following water discharge locations require 500ml sample AND RP approval prior to discharge to Sanitary Sewer', table_name_format )
					worksheet.merge_range('A10:H10', 'Note: [BLANK/NULL] indicates sample taken had no detectable radioactivity. System may be radioactive due to detectable activity in Resin.', table_name_format)
					worksheet.merge_range('A11:H11', '* Sampling required if BSY ABD Sump is discharged to this tank, otherwise no sample or approval required.', table_name_format)

					#Table 2a
					table1_renamed_2.to_excel(writer, sheet_name='Table 2a', index=False, startrow=1)
					header_format_2 = workbook.add_format({'text_wrap': True, 'bottom': 2, 'top': 2, 'left': 2, 'right': 2, 'bold': True, 'fg_color': '#ff6133'})
					worksheet_2 = writer.sheets['Table 2a']
					worksheet_2.conditional_format('A2:H16', {'type': 'no_errors', 'format': cell_format1})
					worksheet_2.set_column('A:H', 20, cell_wrap)
					worksheet_2.set_column('D:D', 20, scientific_format)					
					worksheet_2.set_column('F:F', 20, scientific_format)
					worksheet_2.set_row(0, 40, table_name_format)

					worksheet_2.merge_range('A1:H1', 'Table 2a: The following water discharge locations require 500 ml sample only prior to discharge to Sanitary Sewer. RP Approval is not required.', table_name_format)
					worksheet_2.merge_range('A18:H18', 'Note: [BLANK/NULL] indicates sample taken had no detectable radioactivity. System may be radioactive due to detectable activity in Resin.', table_name_format)

					#Table 2b
					table1_renamed_3.to_excel(writer, sheet_name='Table 2b', index=False, startrow=1)
					header_format_3 = workbook.add_format({'text_wrap': True, 'bottom': 2, 'top': 2, 'left': 2, 'right': 2, 'bold': True, 'fg_color': '#338dff'})
					worksheet_3 = writer.sheets['Table 2b']
					worksheet_3.conditional_format('A2:H7', {'type': 'no_errors', 'format': cell_format1})
					worksheet_3.set_column('A:H', 20, cell_wrap)
					worksheet_3.set_column('D:D', 20, scientific_format)					
					worksheet_3.set_column('F:F', 20, scientific_format)
					worksheet_3.set_row(0, 40, table_name_format)
					worksheet_3.merge_range('A1:H1', 'Table 2B: The following locations do not require sampling prior to discharge. These locations will be sampled annually if water is present.', table_name_format)
					worksheet_3.merge_range('A9:H9', 'Note: [BLANK/NULL] indicates sample taken had no detectable radioactivity. System may be radioactive due to detectable activity in Resin.', table_name_format)

					#Table 3
					table1_renamed_4.to_excel(writer, sheet_name='Table 3', index=False, startrow=1)
					header_format_4 = workbook.add_format({'text_wrap': True, 'bottom': 2, 'top': 2, 'left': 2, 'right': 2, 'bold': True, 'fg_color': '#fff933'})
					worksheet_4 = writer.sheets['Table 3']
					worksheet_4.conditional_format('A2:I24', {'type': 'no_errors', 'format': cell_format1})
					worksheet_4.set_column('A:I', 20, cell_wrap)
					worksheet_4.set_column('D:D', 20, scientific_format)					
					worksheet_4.set_column('F:F', 20, scientific_format)
					worksheet_4.set_row(0, 40, table_name_format)					
					worksheet_4.merge_range('A1:I1', 'Table 3: The following LCW systems requiring RP coverage for breaching system,sampling, changing Resin Bottles and Filters. Contact RPFO #4299 prior to breaching these systems', table_name_format)
					worksheet_4.merge_range('A26:I26', 'Note: [BLANK/NULL] indicates sample taken had no detectable radioactivity. System may be radioactive due to detectable activity in Resin.', table_name_format)

					data = {'Tritium Concentrations':['MDA', '> MDA and ≤ DWL', '≥ DWL and ≤ 5X DWL', '> 5X DWL'], 'RCA/Secondary Containment':['Containment - No Action\nWetted Personnel - No Actions\nVerification - Water Sample from listed systems*\nReporting - Notify RP#4299', 'Containment - Secure Source, Contain, May Dry\nWetted Personnel - No Actions if GERT or RWT\nVerification - Water Sample from listed systems*\nReporting - Notify RP#4299', 'Containment - Secure Source, Contain, May Dry Post Contamination Area to control access\nWetted Personnel - Remove wetted clothing, remove water, dose assessment\nVerification - Water Sample from listed systems*\nReporting - Notify RP#4299', 'Containment - Secure Source, Contain, Clean up, Post Contamination Area\nWetted Personnel - Remove wetted clothing, remove water, dose assessment\nVerification - Water Sample, Contamination Surveys\nReporting - Notify RP#4299'], 'Environmental':['Containment - Secure Source, Contain, May allow water to dry, Protect storm drain, if possible\nWetted Personnel - No Actions\nVerification - Water Sample from listed systems*\nReporting - Notify RP#4299', 'Containment - Secure Source, Contain, May allow water to dry, Protect storm drain, if possible\nWetted Personnel - No Actions if GERT or RWT\nVerification - Water Sample from listed systems*\nReporting - Notify RP#4299', 'Containment - Secure Source, Contain, Clean up, as feasible, May allow water to dry, Protect storm drain, if possible, Post Contamination Area to control access\nWetted Personnel - Remove wetted clothing, remove water, dose assessment\nVerification - Water Sample from listed systems*\nReporting - Notify RP#4299', 'Containment - Secure Source, Contain, Protect storm drain, if possible, Clean up, Post Contamination Area\nWetted Personnel - Remove wetted clothing, remove water, dose assessment\nVerification - Water Sample, Contamination Surveys\nReporting - Notify RP#4299']}

					dataframe1_5 = DataFrame(data)

					dataframe1_5.to_excel(writer, sheet_name='Spill Response Actions', index=False, startrow=0, startcol=0, header=True)
					header_format_5 = workbook.add_format({'text_wrap': True, 'bottom': 2, 'top': 2, 'left': 2, 'right': 2, 'bold': True, 'fg_color': '#e875fa'})					
					worksheet_5 = writer.sheets['Spill Response Actions']
					worksheet_5.set_column('A:A', 30, cell_wrap)
					worksheet_5.set_column('B:D', 80, cell_wrap)
					worksheet_5.set_row(5, 80, cell_wrap)
					worksheet_5.conditional_format('A1:C6', {'type': 'no_errors', 'format': cell_format1})
					worksheet_5.merge_range('A6:C6', 'Listed systems are those systems listed on the RPFO LCW Status sheet\nhttps://sharepoint.slac.stanford.edu/sites/esh/rp/fo-group/FO%20Documents/LCWStatus.pdf\nDWL EPA Drinking Water Limit (2E4 pCi/liter)\nMDA Minimum Detectable Activity (1E3 pCi/liter)', cell_wrap)


					for col_num, value in enumerate(table1_renamed.columns.values):
						worksheet.write(1, col_num, value, header_format)
					for col_num, value in enumerate(table1_renamed_2.columns.values):
						worksheet_2.write(1, col_num, value, header_format_2)
					for col_num, value in enumerate(table1_renamed_3.columns.values):
						worksheet_3.write(1, col_num, value, header_format_3)
					for col_num, value in enumerate(table1_renamed_4.columns.values):
						worksheet_4.write(1, col_num, value, header_format_4)
					for col_num, value in enumerate(dataframe1_5.columns.values):
						worksheet_5.write(0, col_num, value, header_format_5)
		else:
			print("Unusable Connection.  Please check the database and network settings.")
			return
		cursor.close()
		connection.close()

	def send_email():

			import os
			#os.chdir('C:\\Users\\ryanford\\OneDrive - SLAC National Accelerator Laboratory\\2025\\') #test
			os.chdir('C:\\LCW\\') #prod
			#current_directory = os.getcwd()
			text = ("Today's LCW Status Sheet is shown below"
					'\n')

			# MIMEMultipart() creates a container for an email message that can hold
			# different parts, like text and attachments and in next line we are
			# attaching different parts to email container like subject and others.

			message = MIMEMultipart('alternative')
			message['Subject'] = 'LCW Status Sheet'
			message['From'] = 'melanien@slac.stanford.edu'
			message['To'] = 'lcw-status-sheet@slac.stanford.edu' #recipient_email
			message.attach(MIMEText(text, 'html'))

			part = MIMEBase('application', "octet-stream")
			part.set_payload(open("LCW.xlsx", "rb").read())
			encoders.encode_base64(part)
			part.add_header('Content-Disposition', 'attachment; filename="LCW.xlsx"')
			message.attach(part)
				
			try:

				with smtplib.SMTP('SMTPOUT.slac.stanford.edu', 25, timeout = 5) as server:
					server.sendmail('melanien@slac.stanford.edu', 'lcw-status-sheet@slac.stanford.edu', message.as_string())
					server.quit()

			except Exception as e:
					return
		
return_LCW.return_info()
#return_LCW.send_email() #comment out for testing