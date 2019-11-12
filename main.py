from insert_event import insert_event
from xlsx_to_csv import xlsx_to_csv
import pandas as pd


def GoogleCalendar_insert(user_name, file_name):
    """
    終わりの時間を23時にせずに開始時間と同刻にしている
    """
    df = xlsx_to_csv(user_name=user_name, file_name=file_name)
    df_itigou = df.loc['あかり一号店']
    for time in df_itigou[f'{user_name}の出勤時間']:
        time = str(time).replace(' ', 'T')  # GoogleCalendarの時間の形式に変換
        insert_event(start_time=time, end_time=time, summary='あかり一号店')

    df_hanare = df.loc['あかりはなれ']
    for time in df_hanare[f'{user_name}の出勤時間']:
        time = str(time).replace(' ', 'T')  # GoogleCalendarの時間の形式に変換
        insert_event(start_time=time, end_time=time, summary='あかりはなれ')

    shift_num = len(df_itigou[f'{user_name}の出勤時間']) + len(df_hanare[f'{user_name}の出勤時間'])
    print(f"{shift_num}個のシフトがありました")


if __name__ == '__main__':
    user_name = '加納'
    file_name = '2019.11シフト.xlsx'
    GoogleCalendar_insert(user_name, file_name)
