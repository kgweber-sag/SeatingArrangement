def Attendees_from_spreadsheet(excelfile):
    try:
        df = pd.read_excel(excelfile)
        df_attending = df[df.attending == 'Y']
        attendees = []
        for ind, row in df_attending.iterrows():
            head_table = True if row['head table']==1.0 else False
            attendees.append(Attendee(row['name'], row['gender'], row['seniority'],
                                      row['division'], assign_head_table=head_table))
        return attendees
    except FileNotFoundError:
        return []