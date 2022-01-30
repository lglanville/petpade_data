from datetime import datetime, timedelta
import csv
import argparse
from statistics import mean
import sqlite3
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import NamedStyle
from openpyxl.worksheet.table import Table
from time import sleep

DT_FORMATS = ["%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%d/%m/%Y %I:%M:%S %p"]


def round_time(dt, roundTo=60):
    "round time to nearest minute"
    seconds = (dt.replace(tzinfo=None) - dt.min).seconds
    rounding = (seconds+roundTo/2) // roundTo * roundTo
    return dt + timedelta(0,rounding-seconds,-dt.microsecond)


def parse_dt(dt_string):
    "parse a date string into a timestamp, then round to minute"
    timestamp = None
    rounded = None
    for format in DT_FORMATS:
        try:
            timestamp = datetime.strptime(dt_string, format)
            rounded = round_time(timestamp)
            break
        except Exception as e:
            pass
    return timestamp, rounded


def merge_pos(rows):
    "merge several rows re position of doggo to find pos with longest duration"
    positions = Counter()
    for row in rows:
        duration = int(row['duration'])
        if row['position'].lower().startswith('lying'):
            positions['Lying'] += duration
        elif row['position'].lower().startswith('eating'):
            positions['Standing'] += duration
        else:
            positions[row['position']] += duration
    c = positions.most_common(1)
    if c == []:
        return None, None
    else:
        return c[0]


def merge_activity(rows):
    "merge several rows re activity of doggo"
    activity = sum([row['activity'] for row in rows])
    activity_group = '\n'.join(set([row['activity_group'] for row in rows]))
    if activity_group == '':
        activity = None
    return activity, activity_group


def fetch(rows, fields):
    if len(rows) == 1:
        return [rows[0][field] for field in fields]
    elif len(rows) > 1:
        raise ValueError('several values for this timestamp:')
    else:
        return [None for field in fields]


def iter_timestamp(start, end, interval=timedelta(minutes=1), skip_times=[]):
    """iterate in one minute intervals between two datetimes,
    excluding specified periods"""
    timestamp = start
    while timestamp < end:
        if not any([skipstart < timestamp <= skipend for skipstart, skipend in skip_times]):
            yield timestamp
        timestamp += interval

def build_sqlite(data_file, prox_file, datatable_file):
    """compile sqlite db in memory based on csv files"""
    con = sqlite3.connect(":memory:", detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
    cur = con.cursor()
    cur.execute("""CREATE TABLE position (time timestamp,
    rounded_time timestamp, position text, duration text)""")
    cur.execute("""CREATE TABLE pulse (time timestamp,
    rounded_time timestamp, pulse integer)""")
    cur.execute("""CREATE TABLE respiration (time timestamp,
    rounded_time timestamp, respiration integer)""")
    cur.execute("""CREATE TABLE vvti (time timestamp,
    rounded_time timestamp, vvti real)""")
    cur.execute("""CREATE TABLE activity (time timestamp,
    rounded_time timestamp, activity real, activity_group text)""")
    cur.execute("""CREATE TABLE proximity (time timestamp,
    rounded_time timestamp, tas_tim integer, tas_rev integer)""")
    cur.execute("""CREATE TABLE vector_mag (time timestamp,
    rounded_time timestamp, vector_mag real)""")
    con.commit()
    con.row_factory = sqlite3.Row

    with open(data_file, encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for x in range(2):
            next(reader)
        fieldnames = next(reader)
        for row in reader:
            dict_row = dict(zip(fieldnames, row))
            timestamp, rounded = parse_dt(dict_row['Time'])
            if dict_row['Data type'] == 'Position':
                cur.execute(
                    "INSERT INTO position VALUES (?, ?, ?, ?)",
                    (timestamp, rounded, dict_row['Position'], dict_row['Position Duration']))
            elif dict_row['Data type'] == 'VVTI':
                cur.execute(
                    "INSERT INTO vvti VALUES (?, ?, ?)",
                    (timestamp, rounded, dict_row['VVTI']))
            elif dict_row['Data type'] == 'Respiration':
                cur.execute(
                    "INSERT INTO respiration VALUES (?, ?, ?)",
                    (timestamp, rounded, dict_row['Respiration']))
            elif dict_row['Data type'] == 'Pulse':
                cur.execute(
                    "INSERT INTO pulse VALUES (?, ?, ?)",
                    (timestamp, rounded, dict_row['Pulse']))
            elif dict_row['Data type'] == 'Activity':
                cur.execute(
                    "INSERT INTO activity VALUES (?, ?, ?, ?)",
                    (timestamp, rounded, dict_row['Activity'], dict_row['Activity Group']))
    con.commit()
    with open(prox_file, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            timestamp, rounded = parse_dt(row['Timestamp'])
            cur.execute(
                "INSERT INTO proximity VALUES (?, ?, ?, ?)",
                (timestamp, rounded, row['TAS1H24200119 (Tim)'], row['TAS1H50200437 (Rev)']))
    with open(datatable_file, encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for x in range(10):
            next(reader)
        fieldnames = next(reader)
        fieldnames = [f.strip() for f in fieldnames]
        for row in reader:
            dict_row = dict(zip(fieldnames, row))
            ts = f"{dict_row['Date']} {dict_row['Time']}"
            timestamp, rounded = parse_dt(ts)
            cur.execute(
                "INSERT INTO vector_mag VALUES (?, ?, ?)",
                (timestamp, rounded, dict_row['Vector Magnitude']))
    con.commit()
    return con


def merge_prox_rows(rows):
    "merge several rows re proximity of doggo"
    dists = []
    rev_dists = []
    present = []
    rev_present = []
    for row in rows:
        tas = row['tas_tim']
        rev = row['tas_rev']
        try:
            tas = int(tas)
            dist = (((0.0012*(tas**2))+(0.0936*tas)+(1.9262)))
            dists.append(dist)
            present.append(1)
        except ValueError:
            present.append(0)
        try:
            tas_rev = int(rev)
            rev_dist = (((0.0012*(tas_rev**2))+(0.0936*tas_rev)+(1.9262)))
            rev_dists.append(rev_dist)
            rev_present.append(1)
        except ValueError:
            rev_present.append(0)
    if dists != []:
        ave_dist = mean(dists)
    else:
        ave_dist = 0
    if present != []:
        ave_pres = mean(present)
    else:
        ave_pres = 0
    if rev_dists != []:
        ave_rev_dists = mean(rev_dists)
    else:
        ave_rev_dists = 0
    if rev_present != []:
        ave_rev_present = mean(rev_present)
    else:
        ave_rev_present = 0
    return ave_dist, ave_pres, ave_rev_dists, ave_rev_present


def main(data_file, prox_file, datatable_file, output, start, end, skip_rows=0, skip_times=[]):
    "do all the things"
    start_time = datetime.now()
    wb = Workbook()
    date_style = NamedStyle(name='datetime', number_format='DD/MM/YYYY HH:MM')
    RAW_sheet = wb.create_sheet('RAW')
    act_sheet = wb.create_sheet('Activity')
    act_sheet.append(['Time', 'Activity', 'Activity Group'])
    pos_dur_sheet = wb.create_sheet('PositionDuration')
    pos_dur_sheet.append(['Time', 'Position', 'Duration'])
    pulse_sheet = wb.create_sheet('Pulse')
    pulse_sheet.append(['Time', 'Pulse'])
    respiration_sheet = wb.create_sheet('Respiration')
    respiration_sheet.append(['Time', 'Respiration'])
    vvti_sheet = wb.create_sheet('VVTI')
    vvti_sheet.append(['Time', 'VVTI'])
    raw = wb.create_sheet('RAW')
    prox_sheet = wb.create_sheet('Proximity')
    prox_sheet.append(
        [
            'Time/min', 'Tim Distance/min', 'Tim Present/min',
            'Rev Distance/min', 'Rev Present/Min', 'ACtivity', 'Pulse',
            'Respiration', 'VVTI', 'Position', 'Vector Magnitude'])

    with open(data_file, encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for row in reader:
            RAW_sheet.append(row)
    con = build_sqlite(data_file, prox_file, datatable_file)
    cursor = con.cursor()
    for timestamp in iter_timestamp(start, end, skip_times=skip_times):
        cursor.execute("SELECT rounded_time, position, duration FROM position WHERE rounded_time = '%s'" % timestamp)
        position, duration = merge_pos(cursor.fetchall())
        pos_dur_sheet.append([timestamp, position, duration])
        cursor.execute("SELECT rounded_time, activity, activity_group FROM activity WHERE rounded_time = '%s'" % timestamp)
        activity, activity_group = merge_activity(cursor.fetchall())
        act_sheet.append([timestamp, activity, activity_group])
        cursor.execute("SELECT rounded_time, pulse FROM pulse WHERE rounded_time = '%s'" % timestamp)
        pulse = fetch(cursor.fetchall(), ['pulse'])[0]
        pulse_sheet.append([timestamp, pulse])
        cursor.execute("SELECT rounded_time, vvti FROM vvti WHERE rounded_time = '%s'" % timestamp)
        vvti = fetch(cursor.fetchall(), ['vvti'])[0]
        vvti_sheet.append([timestamp, vvti])
        cursor.execute("SELECT rounded_time, respiration FROM respiration WHERE rounded_time = '%s'" % timestamp)
        respiration = fetch(cursor.fetchall(), ['respiration'])[0]
        respiration_sheet.append([timestamp, respiration])
        cursor.execute("SELECT vector_mag FROM vector_mag WHERE rounded_time = '%s'" % timestamp)
        vec_mag = cursor.fetchone()[0]
        cursor.execute("SELECT rounded_time, tas_tim, tas_rev FROM proximity WHERE rounded_time = '%s'" % timestamp)
        prox_sheet.append([timestamp, *merge_prox_rows(cursor.fetchall()), activity, pulse, respiration, vvti, position, vec_mag])
    for worksheet in wb.worksheets:
        for cell in worksheet['A'][1:]:
            cell.style = date_style
    t_coords = prox_sheet[prox_sheet.max_row][-1].coordinate

    tab = Table(displayName="proximity", ref=f"A1:{t_coords}")
    prox_sheet.add_table(tab)
    calc_heads = [
        'Start', 'End', 'Tim Present Count', 'Tim Present %',
        'Tim Present <=0.5m', 'Tim Present <=1m', 'Tim Present <=2m',
        'Tim Present <=0.5m %', 'Tim Present <=1m %', 'Tim Present <=2m %',
        'Rev Present Count', 'Rev Present %', 'Rev Present <=0.5m',
        'Rev Present <=1m', 'Rev Present <=2m', 'Rev Present <=0.5m %',
        'Rev Present <=1m %', 'Rev Present <=2m %']
    for cell, head in zip(prox_sheet["N1:AE1"][0], calc_heads):
        cell.value = head
    days = [start]
    for s, e in skip_times:
        days.append(s)
        days.append(e)
    days.append(end)
    days = [(days[x-1], days[x]) for x in range(1, len(days), 2)]

    for row, days in zip(prox_sheet["N2:AE16"], days):
        row[0].value = days[0]
        row[1].value = days[1]
        row[2].value = f'=SUMIFS(tab[Tim Present/min], tab[Timestamp], ">=" & {row[0].coordinate},  tab[Timestamp], "<=" & {row[0].coordinate})'
    wb.save(output)
    con.close()
    print(datetime.now() - start_time)

if __name__ == "__main__":
    if __name__ == '__main__':
        parser = argparse.ArgumentParser(
            description='script to merge a bunch of doggo data')
        parser.add_argument('data', metavar='i', help='input export file')
        parser.add_argument('proximity', metavar='p', help='input proximity sheet')
        parser.add_argument('datatable', metavar='p', help='input datatable sheet')
        parser.add_argument('output', metavar='o', help='output xlsx file')
        parser.add_argument(
            '--start', '-s', type=datetime.fromisoformat, help='start of monitoring period')
        parser.add_argument(
            '--end', '-e', type=datetime.fromisoformat, help='end of monitoring period')
        parser.add_argument(
            '--chargetimes', '-t', nargs='+', type=datetime.fromisoformat, default=[], help='time periods to skip')
        args = parser.parse_args()
        skiptimes = [(args.chargetimes[x-1], args.chargetimes[x]) for x in range(1, len(args.chargetimes), 2)]
        main(
            args.data, args.proximity, args.datatable, args.output, args.start, args.end, skip_times=skiptimes)
