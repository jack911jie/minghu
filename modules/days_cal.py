import datetime

def calculate_age(birth_s='20181215'):
    birth_d = datetime.datetime.strptime(str(birth_s), "%Y%m%d")
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

class Dates:
    def __init__(self):
        pass

    def check_date(self,s,e):
        if len(s)!=8:            
            print('开始日期长度有误。')
            return 'error'
        if int(s[4:6])>12:
            
            print('开始日期月份大于12')
            return 'error'
        if len(e)!=8:            
            print('结束日期长度有误。')
            return 'error'
        if int(e[4:6])>12:
            print('结束日期月份大于12')
            return 'error'

        if int(s[4:6]) in [1,3,5,7,8,10,12] and int(s[6:])>31:
            print('开始日期月份日期不能大于31')
            return 'error'
        if int(s[4:6]) in [4,6,9,11] and int(s[6:])>30:
            print('开始日期月份日期不能大于30')
            return 'error'
        if int(e[4:6]) in [1,3,5,7,8,10,12] and int(e[6:])>31:
            print('结束日期月份日期不能大于31')
            return 'error'
        if int(e[4:6]) in [4,6,9,11] and int(e[6:])>30:
            print('结束日期月份日期不能大于30')
            return 'error'
        if int(s[0:4])//4==int(s[0:4])/4 or int(s[0:4])//400==int(s[0:4])/400:
            if int(s[4:6])==2:
                if int(s[6:])>29:
                    print('闰年2月日期不能大于28')
                    return 'error'
        else:
            if int(s[4:6])==2:
                if int(s[6:])>28:
                    print('平年2月日期不能大于28')
                    return 'error'

        if int(e[0:4])//4==int(e[0:4])/4 or int(e[0:4])//400==int(e[0:4])/400:
            if int(e[4:6])==2:
                if int(e[6:])>29:
                    print('闰年2月日期不能大于28')
                    return 'error'
        else:
            if int(e[4:6])==2:
                if int(e[6:])>28:
                    print('平年2月日期不能大于28')
                    return 'error'

        return 'OK'

    def month_days(self,y):
        # print(y)
        # y=str(y)
        if y//400==y/400 or y//4==y/4:
            return {'1':31,'2':29,'3':31,'4':30,'5':31,'6':30,'7':31,'8':31,'9':30,'10':31,'11':30,'12':31}
        else:
            return {'1':31,'2':28,'3':31,'4':30,'5':31,'6':30,'7':31,'8':31,'9':30,'10':31,'11':30,'12':31}

        

    def dif_y_m_d(self,s='20211301',e='20210302'):
        
        if self.check_date(s,e) !='OK':
            exit(0)
        
        s1=datetime.datetime.strptime(s,'%Y%m%d')
        y1=s1.year
        m1=s1.month
        d1=s1.day

        s2=datetime.datetime.strptime(e,'%Y%m%d')
        y2=s2.year
        m2=s2.month
        d2=s2.day

        # print(y1,m1,d1,y2,m2,d2)

        if y1<y2:
            if m2>m1:
                d_y=y2-y1
                if d2>=d1:
                    d_m=m2-m1
                    d_d=d2-d1
                else:
                    d_m=m2-m1-1
                    m_d=self.month_days(y2)
                    if m2==1:
                        d_d=int(m_d[str(m2+12-1)])-d1+d2
                    else:
                        d_d=int(m_d[str(m2-1)])-d1+d2
            elif m2<m1:
                d_y=0
                if d2>=d1:
                    d_m=m2+12-m1
                    d_d=d2-d1
                else:
                    d_m=m2+12-m1-1
                    m_d=self.month_days(y2)
                    if m2==1:
                        d_d=int(m_d[str(m2+12-1)])-d1+d2
                    else:
                        d_d=int(m_d[str(m2-1)])-d1+d2
            else: #m2==m1
                if d2>=d1:
                    d_y=y2-y1
                    d_m=0
                    d_d=d2-d1
                else:
                    d_y=y2-y1-1
                    d_m=11
                    m_d=self.month_days(y2)
                    if m2==1:
                        d_d=int(m_d[str(m2+12-1)])-d1+d2
                    else:
                        d_d=int(m_d[str(m2-1)])-d1+d2
        elif y1==y2:
            d_y=0
            if m2>m1:                
                if d2>=d1:
                    d_m=m2-m1
                    d_d=d2-d1
                else:
                    d_m=m2-m1-1
                    m_d=self.month_days(y2)
                    if m2==1:
                        d_d=int(m_d[str(m2+12-1)])-d1+d2
                    else:
                        d_d=int(m_d[str(m2-1)])-d1+d2
            elif m2==m1:
                if d2>=d1:
                    d_m=0
                    d_d=d2-d1
                else:
                    print('日期写反')
                    return 'error'
            else:
                print('日期写反')
                return 'error'
        else:
            print('日期写反')
            return 'error'

        return [d_y,d_m,d_d]