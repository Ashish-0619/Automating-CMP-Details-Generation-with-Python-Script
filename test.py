import pandas as pd
from datetime import datetime, timedelta


excel_file = 'CMP_data.xlsx'
df = pd.read_excel(excel_file)


current_date = datetime(2023, 10, 31)


less_than_7_days = []
between_7_and_30_days = []
between_30_and_60_days = []
between_60_and_90_days = []
greater_than_90_days = []


for index, row in df.iterrows():
    open_date = row['Created']
    open_date_1 = pd.to_datetime(open_date)
    open_date = open_date_1.date()
    print(open_date)
    # days_open = (current_date - open_date).days

    age = 0
    temp_date = open_date           #example: 21-09-2023, current: 29-09-2023
    while temp_date <= current_date.date():
        if temp_date.weekday() < 5:
            age += 1
        temp_date += timedelta(days=1)
    print(age)
    if age <= 7:
        less_than_7_days.append(row['Key'])
    elif 7 < age <= 30:
        between_7_and_30_days.append(row['Key'])
    elif 30 < age <= 60:
        between_30_and_60_days.append(row['Key'])
    elif 60 < age <= 90:
        between_60_and_90_days.append(row['Key'])
    else:
        greater_than_90_days.append(row['Key'])


print("Tickets open less than 7 days:", less_than_7_days )
print("Tickets open between 7 and 30 days:", between_7_and_30_days)
print("Tickets open between 30 and 60 days:", between_30_and_60_days)
print("Tickets open between 60 and 90 days:", between_60_and_90_days)
print("Tickets open greater than 90 days:", greater_than_90_days)



print("Tickets Worked",  	49)
print("# of Ticket Backlog(<7 Days) ", len(less_than_7_days))
print("# of Ticket Backlog(GT_7D_LT_30D)", 	len(between_7_and_30_days))
print("# of Ticket Backlog(GT_30D_LT_60D)", len(between_30_and_60_days)) 	
print("# of Ticket Backlog(GT_60D_LT_90D)", len(between_60_and_90_days))	
print("# of Ticket Backlog(GT_90D )", len(greater_than_90_days))

