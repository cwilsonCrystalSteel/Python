# -*- coding: utf-8 -*-
"""
Created on Mon Apr 21 09:25:38 2025

@author: Netadmin
"""

import os
import pandas as pd
from pathlib import Path
import numpy as np 

output_dir = Path().home() / 'documents' / 'FitterWelderStatsPDFReports'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
    
output_pdf = str(output_dir / "output.pdf")
    
graphics_dir = Path(r'C:\Users\Netadmin\Documents\FitterWelderStatsCharts\2025-04-21')


df_file = Path(r'C:/Users/Netadmin/documents/FitterWelderPerformanceCSVs/all_both_2025-03_2025-04-18-18-15-57.csv')

main_df = pd.read_csv(df_file, index_col=0)
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

state = 'MD'
classification = 'Fitter'
display_cols = ['Name','Earned Hours','Tons','Defect Quantity','Tons per Piece',
                'Direct Hours','Total Hours','Missed Hours']



from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, PageBreak, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet


styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    'TitleStyle',
    parent=styles['Title'],
    fontSize=16,
    alignment=1,  # Center align
    spaceAfter=12
)



# --- Create PDF ---
doc = SimpleDocTemplate(output_pdf, pagesize=A4)
elements = []

def add_state_table(state, classification):
    state_df = main_df[(main_df['Location'] == state) & (main_df['Classification'] == classification)]
    state_df = state_df[display_cols]
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
    
    
    col_headers = state_df_display.columns.tolist()
    col_headers = [Paragraph(i.replace(' ','<br/>')) for i in col_headers]
    

    
    # Convert DataFrame to list-of-lists
    data = [col_headers] + state_df_display.values.tolist()
    
    colWidths = [140,47,35,50,40,40,40,44,58,58]
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
    title_text = f"{state} {classification} Stats"
    title = Paragraph(title_text, title_style)
    elements.append(title)
    
    # add the table in
    elements.append(pdf_table)
    
    note_text = (
    f"The values of Earned Hours, Tons, Defects, are only in relation to when the employee is "
    f"credited as being the {classification} for pieces. Again, these are pieces that the employee "
    f"has earned from fabricating as a {classification.upper()}."
    )
    table_note = Paragraph(note_text, styles['Normal'])
    elements.append(Spacer(1, 12))
    elements.append(table_note)
    
    # Add a page break before graphics
    elements.append(PageBreak())


add_state_table('MD','Welder')
# Build the PDF
doc.build(elements)

#%%

graphics = os.listdir(graphics_dir)
state_graphics = [i for i in graphics if f'State-{state}' in i]

graphcis_display = []





# Add images from the directory
for filename in sorted(os.listdir(graphics_dir)):
    if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
        img_path = os.path.join(graphics_dir, filename)
        img = Image(img_path, width=400, height=300)  # Resize as needed
        elements.append(img)
        elements.append(PageBreak())