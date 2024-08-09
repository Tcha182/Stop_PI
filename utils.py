import datetime
import pandas as pd

# Mapping from day index to French day names
french_days = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi', 'dimanche']

def get_day_name(date):
    date_str = str(date)
    date_obj = datetime.datetime.strptime(date_str, '%Y%m%d')
    return french_days[date_obj.weekday()]

def generate_group_name(dates):
    day_order = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi', 'dimanche']
    day_names = sorted({get_day_name(date) for date in dates}, key=day_order.index)
    
    if len(day_names) == 1:
        return day_names[0]
    if day_names == ['samedi', 'dimanche']:
        return 'week-end'
    if day_names == day_order:
        return 'tous les jours'
    
    day_indices = [day_order.index(day) for day in day_names]
    ranges, start, end = [], day_indices[0], day_indices[0]
    
    for i in range(1, len(day_indices)):
        if day_indices[i] == day_indices[i-1] + 1:
            end = day_indices[i]
        else:
            if end - start >= 2:
                ranges.append((start, end))
            else:
                ranges.append((start,))
                if end != start:
                    ranges.append((end,))
            start, end = day_indices[i], day_indices[i]
    
    if end - start >= 2:
        ranges.append((start, end))
    else:
        ranges.append((start,))
        if end != start:
            ranges.append((end,))
    
    range_names = []
    for r in ranges:
        if len(r) > 1:
            range_names.append(f"{day_order[r[0]]} à {day_order[r[1]]}")
        else:
            range_names.append(day_order[r[0]])
    
    # Flatten the list and join items
    result = []
    temp = []
    for name in range_names:
        if ' à ' in name:
            if temp:
                result.append(', '.join(temp))
                temp = []
            result.append(name)
        else:
            temp.append(name)
    if temp:
        result.append(', '.join(temp))
    
    return ', '.join(result)

def group_dates_by_timetables(df, stop_id):
    filtered_stop_data = df[df['stop_id'] == stop_id]
    grouped = filtered_stop_data.groupby('date')
    timetables = {}

    for date, group in grouped:
        timetable = tuple(map(tuple, group[['arrival_time']].values))
        timetable_hash = hash(timetable)
        
        if timetable_hash in timetables:
            timetables[timetable_hash][0].append(date)
        else:
            timetables[timetable_hash] = ([date], group)
    
    named_groups = {generate_group_name(dates): (dates, timetable) for dates, timetable in timetables.values()}
    return named_groups


def organize_times_by_hour(times, headsign_indices=None):
    if times.isna().all():
        return pd.DataFrame(), pd.DataFrame()

    # Split times into hours and minutes
    split_times = times.apply(lambda x: x.split(':'))
    truncated_hours = split_times.apply(lambda x: x[0])
    truncated_minutes = split_times.apply(lambda x: x[1])

    if truncated_hours.empty:
        return pd.DataFrame(), pd.DataFrame()

    unique_hours = sorted(truncated_hours.unique(), key=lambda x: int(x))
    timetable = pd.DataFrame(index=range(truncated_hours.value_counts().max()), columns=unique_hours)
    headsign_index_table = pd.DataFrame(index=range(truncated_hours.value_counts().max()), columns=unique_hours)

    for hour in unique_hours:
        # Get the minutes corresponding to the current hour
        minutes_in_hour = truncated_minutes[truncated_hours == hour].tolist()
        minutes_in_hour.sort(key=lambda x: int(x))  # Ensure minutes are sorted within each hour

        if headsign_indices is not None:
            headsign_indices_in_hour = headsign_indices[truncated_hours == hour].tolist()

        for i, minute in enumerate(minutes_in_hour):
            timetable.at[i, hour] = minute
            if headsign_indices is not None:
                headsign_index_table.at[i, hour] = headsign_indices_in_hour[i]

    # Rename columns to have 'h' suffix and ensure unique column names
    new_columns = {}
    hour_counts = {}
    for col in timetable.columns:
        hour_int = int(col)
        hour_str = f"{hour_int % 24:02}h"
        if hour_str in hour_counts:
            hour_counts[hour_str] += 1
            hour_str += ' ' * hour_counts[hour_str]  # Append spaces to make it unique
        else:
            hour_counts[hour_str] = 0
        new_columns[col] = hour_str

    timetable.rename(columns=new_columns, inplace=True)
    timetable.fillna('', inplace=True)
    
    headsign_index_table.rename(columns=new_columns, inplace=True)
    headsign_index_table.fillna('', inplace=True)

    return timetable, headsign_index_table



