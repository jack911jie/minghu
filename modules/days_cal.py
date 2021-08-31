import datetime

def calculate_age(birth_s='20181215'):
    birth_d = datetime.datetime.strptime(birth_s, "%Y%m%d")
    today_d = datetime.datetime.now()
    birth_t = birth_d.replace(year=today_d.year)
    if today_d > birth_t:
        age = today_d.year - birth_d.year
    else:
        age = today_d.year - birth_d.year - 1
    return age


def calculate_days(date_input='20181215'):
    today=datetime.datetime.now()
    date_s = datetime.datetime.strptime(date_input, "%Y%m%d")    
    return (today-date_s).days

def calculate_days_2(dates='20200301',datee='20210301'):
    date_s=datetime.datetime.strptime(dates, "%Y%m%d")
    date_e=datetime.datetime.strptime(datee, "%Y%m%d")
    return (date_e-date_s).days