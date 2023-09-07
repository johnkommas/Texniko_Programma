#  Copyright (c) Ioannis E. Kommas 2023. All Rights Reserved

import pandas as pd
import os
import send_mail
from send_mail import users as mail_users

# SETUP OPTION TO DISPLAY ALL DATA IN PRINT
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.max_rows', None)


def export(path_to_file, df, year):
    # RENAME COLUMN NAMES
    df = df.rename(columns={'Tίτλος': 'ΤΙΤΛΟΣ',
                            'Κατηγορία': 'ΚΑΤΗΓΟΡΙΑ',
                            'Υπηρεσία': 'ΥΠΗΡΕΣΙΑ',
                            'ΚΑ Εξόδων': 'ΚΑ ΕΞΟΔΩΝ',
                            'Συνολικός Πρ/σμός Έργου': 'ΣΥΝΟΛΙΚΟΣ ΠΡΟΥΠΟΛΟΓΙΣΜΟΣ ΕΡΓΟΥ',
                            'Υφ. Νομική Δέσμευση': 'ΥΦΙΣΤΑΜΕΝΗ ΝΟΜΙΚΗ ΔΕΣΜΕΥΣΗ',
                            'Εξοφλημένα Τιμολόγια': 'ΕΞΟΦΛΗΜΕΝΑ ΤΙΜΟΛΟΓΙΑ',
                            'Υπόλοιπο Πληρωθέν Υφ. Νομ. Δεσμ.': 'ΥΠΟΛΟΙΠΟ ΠΛΗΡΩΘΕΝ ΥΦ. ΝΟΜ. ΔΕΣΜ.',
                            'Προταθέντα 2024': 'ΠΡΟΤΑΘΕΝΤΑ 2024',
                            'ΤΑΚΤΙΚΑ/ΙΔΙΟΙ ΠΟΡΟΙ': 'ΤΑΚΤΙΚΑ/ΙΔΙΟΙ ΠΟΡΟΙ',
                            'ΣΑΤΑ ΠΟΕ/ΝΕΑ ΣΑΤΑ': 'ΣΑΤΑ ΠΟΕ/ΝΕΑ ΣΑΤΑ',
                            'ΔΙΑΦΟΡΑ/ΑΝΤΑΠΟΔ/ΧΡΗΜ/ΣΕΙΣ': 'ΔΙΑΦΟΡΑ/ΑΝΤΑΠΟΔ/ΧΡΗΜ/ΣΕΙΣ',
                            'Παρατηρήσεις': 'ΠΗΓΗ ΧΡΗΜ/ΣΗΣ'
                            })

    # CHOOSE COLUMNS AND ORDER
    s_df = df[['ΤΙΤΛΟΣ', 'ΚΑΤΗΓΟΡΙΑ', 'ΥΠΗΡΕΣΙΑ', 'ΚΑ ΕΞΟΔΩΝ', 'ΣΥΝΟΛΙΚΟΣ ΠΡΟΥΠΟΛΟΓΙΣΜΟΣ ΕΡΓΟΥ',
               'ΥΦΙΣΤΑΜΕΝΗ ΝΟΜΙΚΗ ΔΕΣΜΕΥΣΗ', 'ΕΞΟΦΛΗΜΕΝΑ ΤΙΜΟΛΟΓΙΑ', 'ΥΠΟΛΟΙΠΟ ΠΛΗΡΩΘΕΝ ΥΦ. ΝΟΜ. ΔΕΣΜ.',
               'ΠΡΟΤΑΘΕΝΤΑ 2024', 'ΤΑΚΤΙΚΑ/ΙΔΙΟΙ ΠΟΡΟΙ', 'ΣΑΤΑ ΠΟΕ/ΝΕΑ ΣΑΤΑ', 'ΔΙΑΦΟΡΑ/ΑΝΤΑΠΟΔ/ΧΡΗΜ/ΣΕΙΣ',
               'ΠΗΓΗ ΧΡΗΜ/ΣΗΣ']]

    # FIRE UP EXCEL WRITER
    with pd.ExcelWriter(path_to_file, engine='xlsxwriter') as writer:
        # CREATE WORKBOOK
        workbook = writer.book

        # ADD FORMATS BELOW
        number_8_pink = workbook.add_format({
            'num_format': '€#,##0.00',
            'align': 'center',
            'valign': 'vcenter',
            'bold': False,
            'text_wrap': True,
            'font_color': '#FF00FF',
            "font_size": 8,
            "font_name": "Avenir Next"})

        number_8_black = workbook.add_format({
            'num_format': '€#,##0.00',
            'align': 'center',
            'valign': 'vcenter',
            'bold': False,
            'text_wrap': True,
            'font_color': 'black',
            "font_size": 8,
            "font_name": "Avenir Next"})

        number_8_green = workbook.add_format({
            'num_format': '€#,##0.00',
            'align': 'center',
            'valign': 'vcenter',
            'bold': False,
            'text_wrap': True,
            'font_color': '#808000',
            "font_size": 8,
            "font_name": "Avenir Next"})

        number_10_black_bold = workbook.add_format({
            'num_format': '€#,##0.00',
            'align': 'center',
            'valign': 'vcenter',
            'bold': True,
            'text_wrap': True,
            'font_color': 'black',
            "font_size": 10,
            "font_name": "Avenir Next"})

        normal_10 = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'bold': False,
            'text_wrap': True,
            "font_size": 10,
            "font_name": "Avenir Next"})

        normal_8 = workbook.add_format({
            'align': 'left',
            'bold': False,
            'valign': 'vcenter',
            'text_wrap': True,
            "font_size": 8,
            "font_name": "Avenir Next"})

        normal_bold_10 = workbook.add_format({
            'align': 'left',
            'bold': True,
            'valign': 'vcenter',
            'text_wrap': True,
            "font_size": 10,
            "font_name": "Avenir Next"})

        normal_bold_10_center = workbook.add_format({
            'align': 'center',
            'bold': True,
            'valign': 'vcenter',
            'text_wrap': True,
            "font_size": 10,
            "font_name": "Avenir Next"})

        header_format = workbook.add_format({
            'text_wrap': True,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            "font_size": 10,
            "font_name": "Avenir Next",
        })

        # PUT DATA INSIDE EXCEL
        s_df.to_excel(writer, sheet_name='TODAY', startcol=2, startrow=1, index=None)

        # ADD COLUMNS WITH HEADER FORMAT
        for col_num, value in enumerate(s_df.columns.values):
            writer.sheets['TODAY'].write(1, col_num + 2, value, header_format)

        # FIRE UP WORKSHEET TO WORK WITH
        worksheet = writer.sheets['TODAY']

        # ADD ΔΡΑΣΗ ONCE PER GROUP
        colors = ['#A3A3A3', "D9D9D9"]
        i = 0
        start = 3
        worksheet.write('B2', 'ΔΡΑΣΗ', header_format)
        for drasi in df['Δράση'].unique():
            center_vert_text = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'rotation': 270,
                "font_name": "Avenir Next",
                'bold': True,
                'font_size': 8,
                'bg_color': colors[i],
                'border': 1
            })
            sql_answer = df[df['Δράση'] == drasi]
            end = start + sql_answer.shape[0] - 1
            worksheet.merge_range(f'B{start}:B{end}', drasi, center_vert_text)
            start = end + 1
            if i == 1:
                i = 0
            else:
                i += 1

        # ADD A/A
        worksheet.write('A2', 'A/A', header_format)
        for i in range(len(df)):
            worksheet.write(i + 2, 0, i + 1)

        # ADD TITLE ΤΕΧΝΙΚΟ ΠΡΟΓΡΑΜΜΑ
        worksheet.merge_range(f'A1:C1', f'ΤΕΧΝΙΚΟ ΠΡΟΓΡΑΜΑ {year}', normal_bold_10_center)

        # Autofit the worksheet.
        worksheet.autofit()
        worksheet.set_column(f'A:A', 4, normal_10)
        worksheet.set_column(f'B:B', 4)
        worksheet.set_column('C:C', 50, normal_bold_10)
        worksheet.set_column('D:D', 6.5, normal_8)
        worksheet.set_column('E:E', 22, normal_8)
        worksheet.set_column('F:F', 15, normal_bold_10_center)
        worksheet.set_column('G:G', 17, number_8_pink)
        worksheet.set_column('H:H', 17, number_8_green)
        worksheet.set_column('I:I', 17, number_8_black)
        worksheet.set_column('J:J', 17, number_8_black)
        worksheet.set_column('K:K', 17, number_8_black)
        worksheet.set_column('L:L', 17, number_8_black)
        worksheet.set_column('M:M', 17, number_8_black)
        worksheet.set_column('N:N', 17, number_8_black)
        worksheet.set_column('O:O', 21, normal_bold_10)

        # FREEZE PANES
        worksheet.freeze_panes(2, 3)

        # ADD TOTALS AT THE BOTTOM
        worksheet.write(f"G{start}", s_df['ΣΥΝΟΛΙΚΟΣ ΠΡΟΥΠΟΛΟΓΙΣΜΟΣ ΕΡΓΟΥ'].sum().round(2), number_10_black_bold)
        worksheet.write(f"H{start}", s_df['ΥΦΙΣΤΑΜΕΝΗ ΝΟΜΙΚΗ ΔΕΣΜΕΥΣΗ'].sum().round(2), number_10_black_bold)
        worksheet.write(f"I{start}", s_df['ΕΞΟΦΛΗΜΕΝΑ ΤΙΜΟΛΟΓΙΑ'].sum().round(2), number_10_black_bold)
        worksheet.write(f"J{start}", s_df['ΥΠΟΛΟΙΠΟ ΠΛΗΡΩΘΕΝ ΥΦ. ΝΟΜ. ΔΕΣΜ.'].sum().round(2), number_10_black_bold)
        worksheet.write(f"K{start}", s_df['ΠΡΟΤΑΘΕΝΤΑ 2024'].sum().round(2), number_10_black_bold)
        worksheet.write(f"L{start}", s_df['ΤΑΚΤΙΚΑ/ΙΔΙΟΙ ΠΟΡΟΙ'].sum().round(2), number_10_black_bold)
        worksheet.write(f"M{start}", s_df['ΣΑΤΑ ΠΟΕ/ΝΕΑ ΣΑΤΑ'].sum().round(2), number_10_black_bold)
        worksheet.write(f"N{start}", s_df['ΔΙΑΦΟΡΑ/ΑΝΤΑΠΟΔ/ΧΡΗΜ/ΣΕΙΣ'].sum().round(2), number_10_black_bold)


def run():
    cwd = os.path.dirname(os.path.abspath(__file__))
    year = 2023  # int(input("ΕΤΟΣ:"))
    files = f'{cwd}/DATA/{year}'
    df = pd.read_excel(f'{files}/1.xls')
    file = f'{cwd}/final.xlsx'
    export(file, df, year)
    send_mail.send_mail(mail_users.get('mail'), mail_users.get('Title'), 'FILE', file, 'final.xlsx')


if __name__ == "__main__":
    run()
