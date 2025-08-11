# -*- coding: utf-8 -*-
"""
Created on Mon Apr 21 09:25:38 2025

@author: Netadmin
"""

import os
import pandas as pd
from pathlib import Path
import numpy as np 
import datetime

    
    
# graphics_dir = Path(r'C:\Users\Netadmin\Documents\FitterWelderStatsCharts\2025-04-21')


# df_file = Path(r'C:/Users/Netadmin/documents/FitterWelderPerformanceCSVs/all_both_2025-03_2025-04-18-18-15-57.csv')



# main_df = pd.read_csv(df_file, index_col=0)
# # this is a place holder until I figure out how to alert on when employees are terminated but being credited with work /
# # did not work any during that month but are credited
# main_df = main_df[~main_df['direct'].isna()]

# main_df = main_df.sort_values('Earned Hours', ascending=False)

# main_df = main_df.rename(columns={'direct':'Direct Hours',
#                                   'indirect':'Indirect Hours',
#                                   'total':'Total Hours',
#                                   'missed':'Missed Hours',
#                                   'notcounted':'Other Hours',
#                                   'Tonnage':'Tons',
#                                   'Tonnage per Piece':'Tons per Piece'})

# state = 'MD'
# classification = 'Fitter'
# display_cols = ['Name','Earned Hours','Tons','Defect Quantity','Tons per Piece',
#                 'Direct Hours','Total Hours','Missed Hours']


# graphics = os.listdir(graphics_dir)
# state_graphics = [i for i in graphics if f'State-{state}' in i]

#%%

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, PageBreak, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import NextPageTemplate
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame


styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    'TitleStyle',
    parent=styles['Title'],
    fontSize=16,
    alignment=1,  # Center align
    spaceAfter=12
)

centered_note_style = ParagraphStyle(
    name='CenteredNote',
    parent=styles['Normal'],
    alignment=TA_CENTER,
    fontSize=10,
    fontName='Helvetica-Oblique',  # Use the italic version of Helvetica
)


class pdf_report():
    def __init__(self, state, csv_file=None, output_pdf=None, aggregate_data=None, past_agg_data=None):
        self.state = state
        self.aggregate_data = aggregate_data
        if csv_file is None and not aggregate_data is None:
            self.csv_file = self.aggregate_data['filepath']
            self.get_bad_dfs()
            
        elif not csv_file is None:
            self.csv_file = csv_file
            
        elif csv_file is None and aggregate_data is None:
            raise Exception("Did not pass either 'csv_file' or 'aggregate_data'")
            
        self.past_agg_data = past_agg_data.copy()
         
        
            
            
        self.extrapolate_month_year()
        self.main_df = self.load_df(self.csv_file)
        #if we have past aggregate data files to look at 
        if not self.past_agg_data is None :
            # go thru each key in the past agg data and convert to df
            for k in self.past_agg_data.keys():
                # convert filepaths to nice dataframes 
                self.past_agg_data[k] = self.load_df(self.past_agg_data[k])
        
            # if we don't already have the current months df in the past agg dict then add id 
            if self.past_agg_data.get(0) is None:
                self.past_agg_data[0] = self.main_df.copy()
            
        
        self.output_pdf = output_pdf
        
        if not os.path.dirname(output_pdf):
            os.makedirs(os.path.dirname(output_pdf))
        
        # This will store the current values to display in the header
        self.header_context = {
            'state': '',
            'classification': '',
            'month': '',
            'year': ''
        }
        
        self.elements = []
        
        
        
        
        
    def build_report(self):
        self.init_document()
    
        # Title page (no header)
        self.elements.append(NextPageTemplate('NoHeader'))
        self.add_TitlePage()
        
        
        # setup the headers 
        # you kinda have to add the header template after the plot, to set it on that page?
        self.add_template("Fitter", "Fitter")
        self.add_template("Welder", "Welder")
        self.add_template("AllEmployees", "Fitters & Welders")
        
        
        # Page 1 - Different tpyes of hours for all employees 
        self.elements.append(NextPageTemplate("AllEmployees"))
        self.do_pagebreak()
        self.add_AllHourTypeComparison()
        
        self.do_pagebreak()
        self.add_MonthOverMonth_Hours('Total Hours', None)
        
        # self.do_pagebreak()
        self.add_MonthOverMonth_Hours('Missed Hours', None)
        
        for classification in ['Fitter','Welder']:
            self.elements.append(NextPageTemplate(classification))
            self.do_pagebreak()
            self.add_state_table(classification)
           
            self.do_pagebreak()
            self.add_MonthOverMonth_Hours('Total Hours', classification)
            
            # self.do_pagebreak()
            self.add_MonthOverMonth_Hours('Missed Hours', classification)
            
           
            # how bad the bad entries effect them
            self.do_pagebreak()
            self.add_badEntries(classification)
        
            
            # Page 3 defects fitter
            self.do_pagebreak()
            self.add_DefectsPlot(classification)
            
            # Page 
            self.do_pagebreak()
            self.add_AverageTonsPerPiece(classification)
            
            # Page 
            self.do_pagebreak()
            self.add_DirectAndTotalEfficiency(classification)
            
            # Page 
            self.do_pagebreak()
            self.add_EarnedAndTotalHours(classification)
            
            # Page 
            self.do_pagebreak()
            self.add_TonsByEmployee(classification)
            
                
        self.build_document()

    
        
        
    def get_bad_dfs(self):
        self.bad_dfs = {}
        classes = ['fitters','welders']
        for classification in classes:
            df_wrong_state = self.aggregate_data[f'wrong_state_{classification}'][self.state]
            df_bad_id = self.aggregate_data[f'fix_{classification}'][self.state]
            df_didnt_work_that_month = self.aggregate_data[f'employee_didnt_work_{classification}'][self.state]
            
            if not df_wrong_state is None:
                df_wrong_state['reason'] = 'Employee Works at Different Shop'
            if not df_bad_id is None:
                df_bad_id['reason'] = 'Invalid Employee ID'
            if not df_didnt_work_that_month is None:
                df_didnt_work_that_month['reason'] = 'Employee Did Not Work in The Month'
            
            
            dfs = [df_wrong_state, df_bad_id, df_didnt_work_that_month]
            dfs = [i for i in dfs if not i is None]
            dfs_together = pd.concat(dfs, axis=0)
            
            
            
            if classification == 'fitters':
                self.bad_dfs['Fitter']  = dfs_together
                
            if classification == 'welders':
                self.bad_dfs['Welder'] = dfs_together
        
        
    def load_df(self, filepath):
        
        self.display_cols = ['Name','Earned Hours','Tons','Defect Quantity','Tons per Piece',
                        'Direct Hours','Total Hours','Missed Hours']

        # load the file
        main_df = pd.read_csv(filepath, index_col=0)
        # this is a place holder until I figure out how to alert on when employees are terminated but being credited with work /
        # did not work any during that month but are credited
        main_df = main_df[~main_df['direct'].isna()]

        main_df = main_df.sort_values('Earned Hours', ascending=False)

        main_df = main_df.rename(columns={'direct':'Direct Hours',
                                          'indirect':'Indirect Hours',
                                          'total':'Total Hours',
                                          'missed':'Missed Hours',
                                          'notcounted':'Other Hours',
                                          'Tonnage':'Tons',
                                          'Tonnage per Piece':'Tons per Piece'})

        return main_df
        
        
        
        
    def extrapolate_month_year(self):
        str_file = os.path.basename(self.csv_file)
        split_name = str_file.split('_')
        if 'prod_' in str_file:
            split_name = split_name[2].split('-')
        else:
            split_name = split_name[1].split('-')
        self.month_name = split_name[0]
        self.year = split_name[1]
        # date = datetime.datetime.strptime(month_name, '%B')
        # self.month_name = datetime.date(year=2000, month=int(month_number), day=1).strftime('%B')
        # self.year = split_name.split('_')[1]
        
        

        

    def init_document(self):
        # how to format the type of the output_pdf b/c we cant pass it a Path()
        output_pdf = str(self.output_pdf) if isinstance(self.output_pdf, Path) else self.output_pdf
        # generaet doc
        self.doc = BaseDocTemplate(output_pdf, pagesize=A4)
        # init the no header page template
        frame = Frame(36, 36, A4[0] - 72, A4[1] - 72, id='normal')
        self.doc.addPageTemplates([
            PageTemplate(id='NoHeader', frames=frame)
        ])   
        


        # --- Create PDF ---
        # doc = SimpleDocTemplate(output_pdf, pagesize=A4)
        
        
    def build_document(self):
        self.doc.build(self.elements)
        print(f'File saved to: {self.output_pdf}')
        
    

        
    def make_header_func(self, state, classification, month, year):
        def header(canvas_obj, doc):
            canvas_obj.saveState()
            canvas_obj.setFont("Helvetica", 10)
    
            left = f"{state} {classification}".strip()
            right = f"{month} {year}".strip()
    
            canvas_obj.drawString(40, A4[1] - 40, left)
            text_width = canvas_obj.stringWidth(right, "Helvetica", 10)
            canvas_obj.drawString(A4[0] - 40 - text_width, A4[1] - 40, right)
    
            canvas_obj.restoreState()
        return header

    def add_template(self, name, classification):
        frame = Frame(36, 36, A4[0] - 72, A4[1] - 72, id='normal')
        header_func = self.make_header_func(
            self.state,
            classification,
            self.month_name,
            self.year
        )
        self.doc.addPageTemplates([
            PageTemplate(id=name, frames=frame, onPage=header_func)
        ])

        
        
    
    def apply_header_template(self, classification):
        self.header_context['state'] = self.state
        self.header_context['classification'] = classification
        self.header_context['month'] = self.month_name
        self.header_context['year'] = self.year
    
    
    def do_pagebreak(self):
        self.elements.append(PageBreak())
        
    def add_TitlePage(self):
        print('Building add_TitlePage...')

        today = datetime.datetime.now()
        formatted_date = today.strftime('%B %d, %Y')  # e.g., April 21, 2025

        # --- Title and Subtitle Styles ---
        title_style = ParagraphStyle(
            name='Title',
            fontSize=32,
            alignment=1,  # Center
            spaceAfter=20
        )
        
        subtitle_style = ParagraphStyle(
            name='Subtitle',
            fontSize=20,
            alignment=1,  # Center
            spaceAfter=40
        )

        footer_style = ParagraphStyle(
            name='Footer',
            fontSize=10,
            alignment=1  # Center
        )

        # --- Add vertical space to center content ---
        self.elements.append(Spacer(1, 180))  # Push content toward vertical center

        # --- Main Title ---
        title = Paragraph(f"Fitter & Welder Stats for {self.state}", title_style)
        self.elements.append(title)

        # --- Subtitle ---
        subtitle = Paragraph(f"{self.month_name} {self.year}", subtitle_style)
        self.elements.append(subtitle)

        # --- Footer (push to bottom of page) ---
        self.elements.append(Spacer(1, 300))  # Adjust as needed based on your page size
        footer = Paragraph(f"Report Generated on {formatted_date}", footer_style)
        self.elements.append(footer)
        self.elements.append(Spacer(1, 12)) 
        quesetions = Paragraph("Questions? Comments? cwilson@mtrade.net", footer_style)
        self.elements.append(quesetions)

        # --- Page break to end title page ---
        # self.do_pagebreak()
        
        
        
    def add_state_table(self, classification):
        print('Building add_state_table...')

        state_df = self.main_df[(self.main_df['Location'] == self.state) & 
                                (self.main_df['Classification'] == classification)].copy()
        state_df = state_df[self.display_cols]
        state_df['DL Efficiency']  = state_df['Earned Hours'] / state_df['Direct Hours']
        state_df['TTL Efficiency'] = state_df['Earned Hours'] / state_df['Total Hours']
        
        state_df['DL Efficiency'] = state_df['DL Efficiency'].replace(np.inf, 0)
        state_df['TTL Efficiency'] = state_df['TTL Efficiency'].replace(np.inf, 0)
        
        state_df_display = state_df.copy()
        rounding_display_dict = {'Earned Hours':'.0f',
                                 'Tons':'.1f',
                                 'Defect Quantity':'.2f',
                                 'Tons per Piece':'.3f',
                                 'Direct Hours':'.0f',
                                 'Total Hours':'.0f',
                                 'Missed Hours':'.0f',
                                 'DL Efficiency':'.0%',
                                 'TTL Efficiency':'.0%'}
        
        
        for col in rounding_display_dict.keys():
            fmt = rounding_display_dict[col]
            state_df_display[col] = state_df_display[col].apply(lambda x: f"{x:{fmt}}")
            
        state_df_display['Name'] = state_df_display['Name'].apply(lambda x: x if len(x) <= 20 else x[:20] + '...')
        
        
        col_headers = state_df_display.columns.tolist()
        col_headers = [Paragraph(i.replace(' ','<br/>')) for i in col_headers]
        

        
        # Convert DataFrame to list-of-lists
        data = [col_headers] + state_df_display.values.tolist()
        
        colWidths = [130,47,35,50,40,40,40,44,58,58]
        # Create Table
        pdf_table = Table(data, colWidths=colWidths)
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        
        # --- Add title above table ---
        title_text = f"Table Summary of {classification} Stats"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        # add the table in
        self.elements.append(pdf_table)
        
        note_text = (
        f"The values of Earned Hours, Tons, Defects, are only in relation to when the employee is "
        f"credited as being the {classification} for pieces. Again, these are pieces that the employee "
        f"has earned from fabricating as a {classification.upper()}."
        )
        table_note = Paragraph(note_text, centered_note_style)
        self.elements.append(Spacer(1, 12))
        self.elements.append(table_note)
        
        # Add a page break before graphics
        # self.do_pagebreak()
        
    def add_AllHourTypeComparison(self):
        print('Building add_AllHourTypeComparison...')
        from fitter_welder_stats.fitter_welder_stats_graphing import hours_comparison_by_employee
        graphic = hours_comparison_by_employee(self.main_df, self.state, SAVEFILES=True)
        if not os.path.exists(graphic):
            raise Exception('Could not find {graphic}')


        title_text = "Classification of Hours Worked"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        
        img = Image(graphic, width=520, height=640)
        
        self.elements.append(img)
        
    def add_MonthOverMonth_Hours(self, hours_type, classification):
        print(f'Building add_MonthOverMonth_Hours {hours_type} {classification}')
        HOURS_TYPE_VALID = ['Total Hours','Direct Hours','Missed Hours']
        from fitter_welder_stats.fitter_welder_stats_graphing import mom_hours_worked_by_shop
        graphic = mom_hours_worked_by_shop(self.past_agg_data, self.state, 
                                           hours_type, classification,
                                           self.month_name, self.year,
                                           SAVEFILES=True)
        
        title_text = f'{hours_type} by Month'
        if not classification is None:
            title_text += f" for {classification}"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        img = Image(graphic, width=520, height = 325)
        
        self.elements.append(img)
        
        
    def add_DefectsPlot(self, classification):
        print(f'Building add_DefectsPlot {classification}')
        from fitter_welder_stats.fitter_welder_stats_graphing import defects_by_employee
        graphic = defects_by_employee(self.main_df, self.state, classification, SAVEFILES=True)
        if not os.path.exists(graphic):
            raise Exception('Could not find {graphic}')
            
        title_text = f"{classification} Type Defects"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        img = Image(graphic, width=487.5, height=650)
        
        self.elements.append(img)
        
        
    def add_AverageTonsPerPiece(self, classification):
        print(f'Building add_AverageTonsPerPiece {classification}')
        from fitter_welder_stats.fitter_welder_stats_graphing import tonnage_per_piece_by_employee
        graphic = tonnage_per_piece_by_employee(self.main_df, self.state, classification, SAVEFILES=True)    
        if not os.path.exists(graphic):
            raise Exception('Could not find {graphic} ')
                    
            
        title_text = f"{classification} Tons Per Piece Completed"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        
        img = Image(graphic, width=520, height=650)
        
        self.elements.append(img)
        
        
    def add_DirectAndTotalEfficiency(self, classification):
        print(f'Building add_DirectAndTotalEfficiency {classification}...')
        from fitter_welder_stats.fitter_welder_stats_graphing import total_direct_labor_efficiency
        graphic = total_direct_labor_efficiency(self.main_df, self.state, classification, SAVEFILES=True)   
        if not os.path.exists(graphic):
            raise Exception('Could not find {graphic}')
            
        title_text = f"{classification} Efficiency (Based on Direct & Total Hours)"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        img = Image(graphic, width=520, height=650)
        
        self.elements.append(img)
        
        
    def add_EarnedAndTotalHours(self, classification):
        print(f'Building add_EarnedAndTotalHours {classification}...')

        from fitter_welder_stats.fitter_welder_stats_graphing import earned_hours_by_employee
        graphic = earned_hours_by_employee(self.main_df, self.state, classification, SAVEFILES=True)
        if not os.path.exists(graphic):
            raise Exception('Could not find {graphic}')
            
        title_text = f"{classification} Earned Hours & Hours Worked"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        img = Image(graphic, width=480, height=600)
        
        self.elements.append(img)
        
        note_text = (
        f"* Employees with more than 100 Earned Hours credited as a {classification}."
        )
        table_note = Paragraph(note_text, centered_note_style)
        self.elements.append(Spacer(1, 12))
        self.elements.append(table_note)
        
    def add_TonsByEmployee(self, classification):
        print(f'Building add_TonsByEmployee {classification}...')

        from fitter_welder_stats.fitter_welder_stats_graphing import weight_by_employee
        graphic = weight_by_employee(self.main_df, self.state, classification, SAVEFILES=True)
        if not os.path.exists(graphic):
            raise Exception('Could not find {graphic}')
            
        title_text = f"Tons Completed as {classification}"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        img = Image(graphic, width=480, height=600)
        
        self.elements.append(img)
        
        note_text = (
        f"* Employees with more than 3 Tons credited as a {classification}."
        )
        table_note = Paragraph(note_text, centered_note_style)
        self.elements.append(Spacer(1, 12))
        self.elements.append(table_note)   
        
        
    def add_badEntries(self, classification):
        print(f'Building add_badEntries {classification}...')

        from fitter_welder_stats.fitter_welder_stats_graphing import bad_EmployeeCredit, bad_entriesPieChartAndTable
        
        graphic1 = bad_EmployeeCredit(self.bad_dfs[classification], self.state, classification, SAVEFILES=True)
        graphic2 = bad_entriesPieChartAndTable(self.main_df, self.bad_dfs[classification], self.state, classification, SAVEFILES=True)
        if not os.path.exists(graphic1['filepath']):
            raise Exception('Could not find {graphic1}')
            
        if not os.path.exists(graphic2['filepath']):
            raise Exception('Could not find {graphic2}')
            
        title_text = f"Effect of Bad Entries into Fablisting - {classification}"
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        
        img = Image(graphic1['filepath'], width=470, height=300)
        
        self.elements.append(img)
        self.elements.append(Spacer(1, 24))  # 12 points = approx 1 line
        
        img = Image(graphic2['filepath'], width=480, height=250)
        
        self.elements.append(img)
        
        self.elements.append(Spacer(1, 12))
        
        note_text = (
        f"<b>* Invalid Employee ID</b>: No ID was entered, OR Employee ID did not match any ID in TimeClock."
        f"<br/><b>* Employee Did Not Work in The Month</b>: Employee ID returned an Employee who did not log hours in Timeclock for the month at hand."
        f"<br/><b>* Employee Works at Different Shop</b>: Employee ID returned an Employee who has a Productive Status indicating a different shop."
        )
        
        table_note = Paragraph(note_text, centered_note_style)
        self.elements.append(Spacer(1, 12))
        self.elements.append(table_note)           
        

        
        

        

#%%



# x = pdf_report('MD', df_file, Path(r'C:/Users/Netadmin/Documents/FitterWelderStatsPDFReports/output2.pdf'))

# x.build_report()

