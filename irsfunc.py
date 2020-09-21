import openpyxl
import datetime
from dateutil.relativedelta import relativedelta
import math
from business_calendar import Calendar


class irsfunc():


    def __init__(self, assessmentdate):
        self.assessmentdate = assessmentdate  # 估值日，如'20190102'
        # 外汇交易中心公布的IRS利率曲线，短端点则用对应浮动利率的即期值，通过interestread（）函数获得
        self.interestoriginal = {'FR007': [], 'SHIBOR3M': [], 'LPR1Y': []}
        self.interesthistory = {
            'FR007': [],
            'SHIBOR3M': [],
            'LPR1Y': []}  # 浮动利率的历史值，通过interestread（）函数获得
        self.asset = {}  # 持仓的IRS资产，通过irsput（）函数获得
        self.interest = {'FR007': {}, 'SHIBOR3M': {},
                         'LPR1Y': {}}  # IRS的收益率曲线，通过interestline（）函数计算

    def interestread(self, interestaddress):
        """读入各类IRS的利率曲线和浮动利率的对应历史值，填好self.interestoriginal和self.interesthistory"""
        assementdate = datetime.datetime.strptime(
            self.assessmentdate, '%Y%m%d')
        wb = openpyxl.load_workbook(interestaddress)
        # 1读入FR007的利率互换曲线和FR007的历史值
        wfr007irs = wb['FR007IRS']
        for i in wfr007irs.rows:
            if i[0].value == assementdate:
                self.interestoriginal['FR007'].append(i[10].value / 100)
                for k in range(9):
                    self.interestoriginal['FR007'].append(i[k + 1].value / 100)

        wfr007 = wb['FR007']
        for i in wfr007.rows:
            if isinstance(i[0].value, datetime.datetime):
                self.interesthistory['FR007'].append(
                    [i[0].value, i[1].value / 100])
        # 2读入LPR1Y的利率互换曲线和FR007的历史值（待写）

        wlpr1y = wb['LPR1Y']
        for i in wlpr1y:
            if isinstance(i[0].value, datetime.datetime):
                self.interesthistory['LPR1Y'].append(
                    [i[0].value, i[1].value / 100])
                if (assementdate - i[0].value).days >= 0:
                    lpr1ynow = i[1].value / 100
        self.interestoriginal['LPR1Y'].append(lpr1ynow)
        wlpr1yirs = wb['LPR1YIRS']
        for i in wlpr1yirs:
            if i[0].value == assementdate:
                for k in range(7):
                    self.interestoriginal['LPR1Y'].append(i[k + 1].value / 100)

        # 3读入SHIBOR3M的利率互换曲线和FR007的历史值（待写）

    def interestline(self, change=0):
        """计算IRS的收益率曲线，填好self.interest"""
        call = Calendar()

        assementdate = datetime.datetime.strptime(
            self.assessmentdate, '%Y%m%d')

        interestoriginal = self.interestoriginal
        changedic = {
            'FR007': {
                '1W': 0,
                '1M': 1,
                '3M': 2,
                '6M': 3,
                '9M': 4,
                '1Y': 5,
                '2Y': 6,
                '3Y': 7,
                '4Y': 8,
                '5Y': 9},
            'LPR1Y': {
                '3M': 0,
                '6M': 1,
                '9M': 2,
                '1Y': 3,
                '2Y': 4,
                '3Y': 5,
                '4Y': 6,
                '5Y': 7}}

        if change != 0:  # 如果不输入参数，就按市场利率算折现曲线，如果输入了参数，就先更改市场利率，再算曲线
            for i in change:
                interestoriginal[i[0]][changedic[i[0]][i[1]]] += i[2]
        # 1计算FR007的IRS收益率曲线
        interest = interestoriginal['FR007']
        # 1.1计算1W的值
        date1w = assementdate + datetime.timedelta(days=7)
        while not call.isworkday(date1w):
            date1w = date1w + datetime.timedelta(days=1)
        t1w = (date1w - assementdate).days
        T1w = t1w
        i1w = interest[0]
        d1w = 1 / (1 + i1w * T1w / 365)
        r1w = -math.log(d1w) * 365 / T1w
        self.interest['FR007']['1W'] = [date1w, t1w, T1w, i1w, r1w, d1w]
        # 1.2计算1M的值
        date1m = assementdate + relativedelta(months=1)
        while not call.isworkday(date1m):
            date1m = date1m + datetime.timedelta(days=1)

        t1m = (date1m - date1w).days
        T1m = (date1m - assementdate).days
        i1m = interest[1]
        d1m = 1 / (1 + i1m * T1m / 365)
        r1m = -math.log(d1m) * 365 / T1m
        self.interest['FR007']['1M'] = [date1m, t1m, T1m, i1m, r1m, d1m]
        # 1.3计算3M的值
        date3m = assementdate + relativedelta(months=3)
        while not call.isworkday(date3m):
            date3m = date3m + datetime.timedelta(days=1)
        t3m = (date3m - assementdate).days
        T3m = t3m
        i3m = interest[2]
        d3m = 1 / (1 + i3m * T3m / 365)
        r3m = -math.log(d3m) * 365 / T3m
        self.interest['FR007']['3M'] = [date3m, t3m, T3m, i3m, r3m, d3m]
        # 1.4计算6M的值
        date6m = assementdate + relativedelta(months=6)
        while not call.isworkday(date6m):
            date6m = date6m + datetime.timedelta(days=1)
        t6m = (date6m - date3m).days
        T6m = (date6m - assementdate).days
        i6m = interest[3]
        d6m = (1 - i6m * t3m / 365 * d3m) / (1 + i6m * t6m / 365)
        r6m = -math.log(d6m) * 365 / T6m
        self.interest['FR007']['6M'] = [date6m, t6m, T6m, i6m, r6m, d6m]
        # 1.5计算9M的值
        date9m = assementdate + relativedelta(months=9)
        while not call.isworkday(date9m):
            date9m = date9m + datetime.timedelta(days=1)
        t9m = (date9m - date6m).days
        T9m = (date9m - assementdate).days
        i9m = interest[4]
        d9m = (1 - i9m * t3m / 365 * d3m - i9m *
               t6m / 365 * d6m) / (1 + i9m * t9m / 365)
        r9m = -math.log(d9m) * 365 / T9m
        self.interest['FR007']['9M'] = [date9m, t9m, T9m, i9m, r9m, d9m]
        # 1.6计算1Y的值
        date1y = assementdate + relativedelta(months=12)
        while not call.isworkday(date1y):
            date1y = date1y + datetime.timedelta(days=1)
        t1y = (date1y - date9m).days
        T1y = (date1y - assementdate).days
        i1y = interest[5]
        d1y = (1 - i1y * t3m / 365 * d3m - i1y * t6m / 365 *
               d6m - i1y * t9m / 365 * d9m) / (1 + i1y * t1y / 365)
        r1y = -math.log(d1y) * 365 / T1y
        self.interest['FR007']['1Y'] = [date1y, t1y, T1y, i1y, r1y, d1y]
        # 1.6计算1.25Y,1.5Y,1.75Y和2Y的值(用二分法求值)
        date125y = assementdate + relativedelta(months=15)
        while not call.isworkday(date125y):
            date125y = date125y + datetime.timedelta(days=1)
        t125y = (date125y - date1y).days
        T125y = (date125y - assementdate).days

        date15y = assementdate + relativedelta(months=18)
        while not call.isworkday(date15y):
            date15y = date15y + datetime.timedelta(days=1)
        t15y = (date15y - date125y).days
        T15y = (date15y - assementdate).days

        date175y = assementdate + relativedelta(months=21)
        while not call.isworkday(date175y):
            date175y = date175y + datetime.timedelta(days=1)
        t175y = (date175y - date15y).days
        T175y = (date175y - assementdate).days

        date2y = assementdate + relativedelta(months=24)
        while not call.isworkday(date2y):
            date2y = date2y + datetime.timedelta(days=1)
        t2y = (date2y - date175y).days
        T2y = (date2y - assementdate).days
        i2y = interest[6]

        ks = r1y - 0.04
        kb = r1y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            # 先假设一个r2y的值，假定它等于r1y
            r2y = (ks + kb) / 2

            # 根据r2y的假设值，用线性插值法计算三个期限的即期收益率
            r125y = r1y + (T125y - T1y) / (T2y - T1y) * (r2y - r1y)
            r15y = r1y + (T15y - T1y) / (T2y - T1y) * (r2y - r1y)
            r175y = r1y + (T175y - T1y) / (T2y - T1y) * (r2y - r1y)

            # 根据即期收益率，计算折现因子
            d125y = math.exp(-r125y * T125y / 365)
            d15y = math.exp(-r15y * T15y / 365)
            d175y = math.exp(-r175y * T175y / 365)
            d2y = math.exp(-r2y * T2y / 365)

            # cal的理论值是1，即如果一开始假设的r2y使得cal的值为1，那么这个r2y就是真实的r2y
            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y * \
                   d125y + t15y * d15y + t175y * d175y + t2y * d2y) * i2y / 365 + d2y

            # 根据cal计算的值，用二分法不断逼近计算得到r2y
            if cal > 1:
                ks = r2y
            else:
                kb = r2y
        self.interest['FR007']['125Y'] = [
            date125y, t125y, T125y, 0, r125y, d125y]
        self.interest['FR007']['15Y'] = [date15y, t15y, T15y, 0, r15y, d15y]
        self.interest['FR007']['175Y'] = [
            date175y, t175y, T175y, 0, r175y, d175y]
        self.interest['FR007']['2Y'] = [date2y, t2y, T2y, i2y, r2y, d2y]
        # 1.7计算2.25Y,2.5Y,2.75Y和3Y的值(用二分法求值)
        date225y = assementdate + relativedelta(months=27)
        while not call.isworkday(date225y):
            date225y = date225y + datetime.timedelta(days=1)
        t225y = (date225y - date2y).days
        T225y = (date225y - assementdate).days

        date25y = assementdate + relativedelta(months=30)
        while not call.isworkday(date25y):
            date25y = date25y + datetime.timedelta(days=1)
        t25y = (date25y - date225y).days
        T25y = (date25y - assementdate).days

        date275y = assementdate + relativedelta(months=33)
        while not call.isworkday(date275y):
            date275y = date275y + datetime.timedelta(days=1)
        t275y = (date275y - date25y).days
        T275y = (date275y - assementdate).days

        date3y = assementdate + relativedelta(months=36)
        while not call.isworkday(date3y):
            date3y = date3y + datetime.timedelta(days=1)
        t3y = (date3y - date275y).days
        T3y = (date3y - assementdate).days
        i3y = interest[7]

        ks = r2y - 0.04
        kb = r2y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            r3y = (ks + kb) / 2

            r225y = r2y + (T225y - T2y) / (T3y - T2y) * (r3y - r2y)
            r25y = r2y + (T25y - T2y) / (T3y - T2y) * (r3y - r2y)
            r275y = r2y + (T275y - T2y) / (T3y - T2y) * (r3y - r2y)

            d225y = math.exp(-r225y * T225y / 365)
            d25y = math.exp(-r25y * T25y / 365)
            d275y = math.exp(-r275y * T275y / 365)
            d3y = math.exp(-r3y * T3y / 365)

            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y * d125y + t15y * d15y +
                   t175y * d175y + t2y * d2y + t225y * d225y + t25y * d25y +
                   t275y * d275y + t3y * d3y) * i3y / 365 + d3y

            if cal > 1:
                ks = r3y
            else:
                kb = r3y
        self.interest['FR007']['225Y'] = [
            date225y, t225y, T225y, 0, r225y, d225y]
        self.interest['FR007']['25Y'] = [date25y, t25y, T25y, 0, r25y, d25y]
        self.interest['FR007']['275Y'] = [
            date275y, t275y, T275y, 0, r275y, d275y]
        self.interest['FR007']['3Y'] = [date3y, t3y, T3y, i3y, r3y, d3y]
        # 1.8计算3.25Y,3.5Y,3.75Y和4Y的值(用二分法求值)
        date325y = assementdate + relativedelta(months=39)
        while not call.isworkday(date325y):
            date325y = date325y + datetime.timedelta(days=1)
        t325y = (date325y - date3y).days
        T325y = (date325y - assementdate).days

        date35y = assementdate + relativedelta(months=42)
        while not call.isworkday(date35y):
            date35y = date35y + datetime.timedelta(days=1)
        t35y = (date35y - date325y).days
        T35y = (date35y - assementdate).days

        date375y = assementdate + relativedelta(months=45)
        while not call.isworkday(date375y):
            date375y = date375y + datetime.timedelta(days=1)
        t375y = (date375y - date35y).days
        T375y = (date375y - assementdate).days

        date4y = assementdate + relativedelta(months=48)
        while not call.isworkday(date4y):
            date4y = date4y + datetime.timedelta(days=1)
        t4y = (date4y - date375y).days
        T4y = (date4y - assementdate).days
        i4y = interest[8]

        ks = r3y - 0.04
        kb = r3y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            r4y = (ks + kb) / 2

            r325y = r3y + (T325y - T3y) / (T4y - T3y) * (r4y - r3y)
            r35y = r3y + (T35y - T3y) / (T4y - T3y) * (r4y - r3y)
            r375y = r3y + (T375y - T3y) / (T4y - T3y) * (r4y - r3y)

            d325y = math.exp(-r325y * T325y / 365)
            d35y = math.exp(-r35y * T35y / 365)
            d375y = math.exp(-r375y * T375y / 365)
            d4y = math.exp(-r4y * T4y / 365)

            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y * d125y + t15y * d15y +
                   t175y * d175y + t2y * d2y + t225y * d225y + t25y * d25y +
                   t275y * d275y + t3y * d3y + t325y * d325y + t35y * d35y +
                   t375y * d375y + t4y * d4y) * i4y / 365 + d4y

            if cal > 1:
                ks = r4y
            else:
                kb = r4y
        self.interest['FR007']['325Y'] = [
            date325y, t325y, T325y, 0, r325y, d325y]
        self.interest['FR007']['35Y'] = [date35y, t35y, T35y, 0, r35y, d35y]
        self.interest['FR007']['375Y'] = [
            date375y, t375y, T375y, 0, r375y, d375y]
        self.interest['FR007']['4Y'] = [date4y, t4y, T4y, i4y, r4y, d4y]

        # 1.9计算4.25Y,4.5Y,4.75Y和5Y的值(用二分法求值)

        date425y = assementdate + relativedelta(months=51)
        while not call.isworkday(date425y):
            date425y = date425y + datetime.timedelta(days=1)
        t425y = (date425y - date4y).days
        T425y = (date425y - assementdate).days

        date45y = assementdate + relativedelta(months=54)
        while not call.isworkday(date45y):
            date45y = date45y + datetime.timedelta(days=1)
        t45y = (date45y - date425y).days
        T45y = (date45y - assementdate).days

        date475y = assementdate + relativedelta(months=57)
        while not call.isworkday(date475y):
            date475y = date475y + datetime.timedelta(days=1)
        t475y = (date475y - date45y).days
        T475y = (date475y - assementdate).days

        date5y = assementdate + relativedelta(months=60)
        while not call.isworkday(date5y):
            date5y = date5y + datetime.timedelta(days=1)
        t5y = (date5y - date475y).days
        T5y = (date5y - assementdate).days
        i5y = interest[9]

        ks = r4y - 0.04
        kb = r4y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            r5y = (ks + kb) / 2

            r425y = r4y + (T425y - T4y) / (T5y - T4y) * (r5y - r4y)
            r45y = r4y + (T45y - T4y) / (T5y - T4y) * (r5y - r4y)
            r475y = r4y + (T475y - T4y) / (T5y - T4y) * (r5y - r4y)

            d425y = math.exp(-r425y * T425y / 365)
            d45y = math.exp(-r45y * T45y / 365)
            d475y = math.exp(-r475y * T475y / 365)
            d5y = math.exp(-r5y * T5y / 365)

            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y * d125y + t15y * d15y +
                   t175y * d175y + t2y * d2y + t225y * d225y + t25y * d25y +
                   t275y * d275y + t3y * d3y + t325y * d325y + t35y * d35y +
                   t375y * d375y + t4y * d4y + t425y * d425y + t45y * d45y +
                   t475y * d475y + t5y * d5y) * i5y / 365 + d5y

            if cal > 1:
                ks = r5y
            else:
                kb = r5y
        self.interest['FR007']['425Y'] = [
            date425y, t425y, T425y, 0, r425y, d425y]
        self.interest['FR007']['45Y'] = [date45y, t45y, T45y, 0, r45y, d45y]
        self.interest['FR007']['475Y'] = [
            date475y, t475y, T475y, 0, r475y, d475y]
        self.interest['FR007']['5Y'] = [date5y, t5y, T5y, i5y, r5y, d5y]

        # 2计算LPR1Y的IRS收益率曲线（待写）
        interest = interestoriginal['LPR1Y']
        # 2.1计算3M的值
        date3m = assementdate + relativedelta(months=3)
        while not call.isworkday(date3m):
            date3m = date3m + datetime.timedelta(days=1)
        t3m = (date3m - assementdate).days
        T3m = t3m
        i3m = interest[0]
        d3m = 1 / (1 + i3m * T3m / 360)
        r3m = -math.log(d3m) * 365 / T3m
        self.interest['LPR1Y']['3M'] = [date3m, t3m, T3m, i3m, r3m, d3m]
        # 2.2计算6M的值
        date6m = assementdate + relativedelta(months=6)
        while not call.isworkday(date6m):
            date6m = date6m + datetime.timedelta(days=1)
        t6m = (date6m - date3m).days
        T6m = (date6m - assementdate).days
        i6m = interest[1]
        d6m = (1 - i6m * t3m / 365 * d3m) / (1 + i6m * t6m / 365)
        r6m = -math.log(d6m) * 365 / T6m
        self.interest['LPR1Y']['6M'] = [date6m, t6m, T6m, i6m, r6m, d6m]
        # 2.3计算9M的值
        date9m = assementdate + relativedelta(months=9)
        while not call.isworkday(date9m):
            date9m = date9m + datetime.timedelta(days=1)
        t9m = (date9m - date6m).days
        T9m = (date9m - assementdate).days
        i9m = interest[2]
        d9m = (1 - i9m * t3m / 365 * d3m - i9m *
               t6m / 365 * d6m) / (1 + i9m * t9m / 365)
        r9m = -math.log(d9m) * 365 / T9m
        self.interest['LPR1Y']['9M'] = [date9m, t9m, T9m, i9m, r9m, d9m]
        # 2.4计算1Y的值
        date1y = assementdate + relativedelta(months=12)
        while not call.isworkday(date1y):
            date1y = date1y + datetime.timedelta(days=1)
        t1y = (date1y - date9m).days
        T1y = (date1y - assementdate).days
        i1y = interest[3]
        d1y = (1 - i1y * t3m / 365 * d3m - i1y * t6m / 365 *
               d6m - i1y * t9m / 365 * d9m) / (1 + i1y * t1y / 365)
        r1y = -math.log(d1y) * 365 / T1y
        self.interest['LPR1Y']['1Y'] = [date1y, t1y, T1y, i1y, r1y, d1y]
        # 2.5计算1.25Y，1.5Y，1.75Y和2Y的值（用二分法求值）
        date125y = assementdate + relativedelta(months=15)
        while not call.isworkday(date125y):
            date125y = date125y + datetime.timedelta(days=1)
        t125y = (date125y - date1y).days
        T125y = (date125y - assementdate).days

        date15y = assementdate + relativedelta(months=18)
        while not call.isworkday(date15y):
            date15y = date15y + datetime.timedelta(days=1)
        t15y = (date15y - date125y).days
        T15y = (date15y - assementdate).days

        date175y = assementdate + relativedelta(months=21)
        while not call.isworkday(date175y):
            date175y = date175y + datetime.timedelta(days=1)
        t175y = (date175y - date15y).days
        T175y = (date175y - assementdate).days

        date2y = assementdate + relativedelta(months=24)
        while not call.isworkday(date2y):
            date2y = date2y + datetime.timedelta(days=1)
        t2y = (date2y - date175y).days
        T2y = (date2y - assementdate).days
        i2y = interest[4]

        ks = r1y - 0.04
        kb = r1y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            # 先假设一个r2y的值，假定它等于r1y
            r2y = (ks + kb) / 2

            # 根据r2y的假设值，用线性插值法计算三个期限的即期收益率
            r125y = r1y + (T125y - T1y) / (T2y - T1y) * (r2y - r1y)
            r15y = r1y + (T15y - T1y) / (T2y - T1y) * (r2y - r1y)
            r175y = r1y + (T175y - T1y) / (T2y - T1y) * (r2y - r1y)

            # 根据即期收益率，计算折现因子
            d125y = math.exp(-r125y * T125y / 365)
            d15y = math.exp(-r15y * T15y / 365)
            d175y = math.exp(-r175y * T175y / 365)
            d2y = math.exp(-r2y * T2y / 365)

            # cal的理论值是1，即如果一开始假设的r2y使得cal的值为1，那么这个r2y就是真实的r2y
            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y *
                   d125y + t15y * d15y + t175y * d175y + t2y * d2y) * i2y / 365 + d2y

            # 根据cal计算的值，用二分法不断逼近计算得到r2y
            if cal > 1:
                ks = r2y
            else:
                kb = r2y
        self.interest['LPR1Y']['125Y'] = [
            date125y, t125y, T125y, 0, r125y, d125y]
        self.interest['LPR1Y']['15Y'] = [date15y, t15y, T15y, 0, r15y, d15y]
        self.interest['LPR1Y']['175Y'] = [
            date175y, t175y, T175y, 0, r175y, d175y]
        self.interest['LPR1Y']['2Y'] = [date2y, t2y, T2y, i2y, r2y, d2y]

        # 2.6计算2.25Y，2.5Y，2.75Y和3Y的值（用二分法求值）
        date225y = assementdate + relativedelta(months=27)
        while not call.isworkday(date225y):
            date225y = date225y + datetime.timedelta(days=1)
        t225y = (date225y - date2y).days
        T225y = (date225y - assementdate).days

        date25y = assementdate + relativedelta(months=30)
        while not call.isworkday(date25y):
            date25y = date25y + datetime.timedelta(days=1)
        t25y = (date25y - date225y).days
        T25y = (date25y - assementdate).days

        date275y = assementdate + relativedelta(months=33)
        while not call.isworkday(date275y):
            date275y = date275y + datetime.timedelta(days=1)
        t275y = (date275y - date25y).days
        T275y = (date275y - assementdate).days

        date3y = assementdate + relativedelta(months=36)
        while not call.isworkday(date3y):
            date3y = date3y + datetime.timedelta(days=1)
        t3y = (date3y - date275y).days
        T3y = (date3y - assementdate).days
        i3y = interest[5]

        ks = r2y - 0.04
        kb = r2y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            r3y = (ks + kb) / 2

            r225y = r2y + (T225y - T2y) / (T3y - T2y) * (r3y - r2y)
            r25y = r2y + (T25y - T2y) / (T3y - T2y) * (r3y - r2y)
            r275y = r2y + (T275y - T2y) / (T3y - T2y) * (r3y - r2y)

            d225y = math.exp(-r225y * T225y / 365)
            d25y = math.exp(-r25y * T25y / 365)
            d275y = math.exp(-r275y * T275y / 365)
            d3y = math.exp(-r3y * T3y / 365)

            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y * d125y + t15y * d15y +
                   t175y * d175y + t2y * d2y + t225y * d225y + t25y * d25y +
                   t275y * d275y + t3y * d3y) * i3y / 365 + d3y

            if cal > 1:
                ks = r3y
            else:
                kb = r3y
        self.interest['LPR1Y']['225Y'] = [
            date225y, t225y, T225y, 0, r225y, d225y]
        self.interest['LPR1Y']['25Y'] = [date25y, t25y, T25y, 0, r25y, d25y]
        self.interest['LPR1Y']['275Y'] = [
            date275y, t275y, T275y, 0, r275y, d275y]
        self.interest['LPR1Y']['3Y'] = [date3y, t3y, T3y, i3y, r3y, d3y]
        # 2.7计算3.25Y，3.5Y，3.75Y和4Y的值（用二分法求值）
        date325y = assementdate + relativedelta(months=39)
        while not call.isworkday(date325y):
            date325y = date325y + datetime.timedelta(days=1)
        t325y = (date325y - date3y).days
        T325y = (date325y - assementdate).days

        date35y = assementdate + relativedelta(months=42)
        while not call.isworkday(date35y):
            date35y = date35y + datetime.timedelta(days=1)
        t35y = (date35y - date325y).days
        T35y = (date35y - assementdate).days

        date375y = assementdate + relativedelta(months=45)
        while not call.isworkday(date375y):
            date375y = date375y + datetime.timedelta(days=1)
        t375y = (date375y - date35y).days
        T375y = (date375y - assementdate).days

        date4y = assementdate + relativedelta(months=48)
        while not call.isworkday(date4y):
            date4y = date4y + datetime.timedelta(days=1)
        t4y = (date4y - date375y).days
        T4y = (date4y - assementdate).days
        i4y = interest[6]

        ks = r3y - 0.04
        kb = r3y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            r4y = (ks + kb) / 2

            r325y = r3y + (T325y - T3y) / (T4y - T3y) * (r4y - r3y)
            r35y = r3y + (T35y - T3y) / (T4y - T3y) * (r4y - r3y)
            r375y = r3y + (T375y - T3y) / (T4y - T3y) * (r4y - r3y)

            d325y = math.exp(-r325y * T325y / 365)
            d35y = math.exp(-r35y * T35y / 365)
            d375y = math.exp(-r375y * T375y / 365)
            d4y = math.exp(-r4y * T4y / 365)

            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y * d125y + t15y * d15y +
                   t175y * d175y + t2y * d2y + t225y * d225y + t25y * d25y +
                   t275y * d275y + t3y * d3y + t325y * d325y + t35y * d35y +
                   t375y * d375y + t4y * d4y) * i4y / 365 + d4y

            if cal > 1:
                ks = r4y
            else:
                kb = r4y
        self.interest['LPR1Y']['325Y'] = [
            date325y, t325y, T325y, 0, r325y, d325y]
        self.interest['LPR1Y']['35Y'] = [date35y, t35y, T35y, 0, r35y, d35y]
        self.interest['LPR1Y']['375Y'] = [
            date375y, t375y, T375y, 0, r375y, d375y]
        self.interest['LPR1Y']['4Y'] = [date4y, t4y, T4y, i4y, r4y, d4y]

        # 2.8计算4.25Y，4.5Y，4.75Y和5Y的值（用二分法求值）

        date425y = assementdate + relativedelta(months=51)
        while not call.isworkday(date425y):
            date425y = date425y + datetime.timedelta(days=1)
        t425y = (date425y - date4y).days
        T425y = (date425y - assementdate).days

        date45y = assementdate + relativedelta(months=54)
        while not call.isworkday(date45y):
            date45y = date45y + datetime.timedelta(days=1)
        t45y = (date45y - date425y).days
        T45y = (date45y - assementdate).days

        date475y = assementdate + relativedelta(months=57)
        while not call.isworkday(date475y):
            date475y = date475y + datetime.timedelta(days=1)
        t475y = (date475y - date45y).days
        T475y = (date475y - assementdate).days

        date5y = assementdate + relativedelta(months=60)
        while not call.isworkday(date5y):
            date5y = date5y + datetime.timedelta(days=1)
        t5y = (date5y - date475y).days
        T5y = (date5y - assementdate).days
        i5y = interest[7]

        ks = r4y - 0.04
        kb = r4y + 0.04
        cal = 0
        while abs(cal * 10000000000 - 10000000000) > 1:
            r5y = (ks + kb) / 2

            r425y = r4y + (T425y - T4y) / (T5y - T4y) * (r5y - r4y)
            r45y = r4y + (T45y - T4y) / (T5y - T4y) * (r5y - r4y)
            r475y = r4y + (T475y - T4y) / (T5y - T4y) * (r5y - r4y)

            d425y = math.exp(-r425y * T425y / 365)
            d45y = math.exp(-r45y * T45y / 365)
            d475y = math.exp(-r475y * T475y / 365)
            d5y = math.exp(-r5y * T5y / 365)

            cal = (t3m * d3m + t6m * d6m + t9m * d9m + t1y * d1y + t125y * d125y + t15y * d15y +
                   t175y * d175y + t2y * d2y + t225y * d225y + t25y * d25y +
                   t275y * d275y + t3y * d3y + t325y * d325y + t35y * d35y +
                   t375y * d375y + t4y * d4y + t425y * d425y + t45y * d45y +
                   t475y * d475y + t5y * d5y) * i5y / 365 + d5y

            if cal > 1:
                ks = r5y
            else:
                kb = r5y
        self.interest['LPR1Y']['425Y'] = [
            date425y, t425y, T425y, 0, r425y, d425y]
        self.interest['LPR1Y']['45Y'] = [date45y, t45y, T45y, 0, r45y, d45y]
        self.interest['LPR1Y']['475Y'] = [
            date475y, t475y, T475y, 0, r475y, d475y]
        self.interest['LPR1Y']['5Y'] = [date5y, t5y, T5y, i5y, r5y, d5y]

        # 3计算SHIBOR3M的IRS收益率曲线（待写）

    def irsput(self, irsaddress):
        """读入现有的IRS资产，填好self.asset"""
        wb = openpyxl.load_workbook(irsaddress)
        ws = wb.active
        for i in ws.rows:
            if i[1].value == '未到期交易':
                code = i[11].value

                if i[20].value == 'FR007利率互换收盘曲线(均值)':
                    irstype = 'FR007'
                elif i[20].value == 'LPR1Y(季付)互换收盘均值曲线':
                    irstype = 'LPR1Y'
                # 现在只有两种，其他品种后期加
                if i[32].value == '固定利率':
                    position = -1  # -1表示利率互换空头
                    fixrate = i[33].value / 100
                else:
                    position = 1
                    fixrate = i[54].value / 100
                valuedate = datetime.datetime.strptime(
                    i[7].value.replace('-', ''), '%Y%m%d')
                period = i[10].value
                facevalue = i[16].value
                if abs(period - 365) < 5:
                    periody = 1
                elif abs(period - 1827) < 5:
                    periody = 5
                elif abs(period - 275) < 5:
                    periody = 3/4
                self.asset[code] = [
                    irstype,
                    valuedate,
                    position,
                    periody,
                    fixrate,
                    facevalue]

    def valuecal(self, output=0):
        """计算IRS的价值"""
        interestall = self.interest
        if output == 0:
            print('------各个IRS的具体现金流------')

        cal = Calendar()
        valuedic = {}
        value = 0
        valuefr007s1y=0
        valuefr007s5y=0
        valuefr007s9m=0
        valuelpr1y=0
        assessmentdate = datetime.datetime.strptime(
            self.assessmentdate, '%Y%m%d')  # 估值日

        for code, info in self.asset.items():

            listbb = []  # 用来放付息区间的应付利息
            listaa = []  # 用来放计息区间的每个重定价的应付利息

            # 1.1计算FR007的IRS的价值
            if info[0] == 'FR007':
                interesthistory = self.interesthistory['FR007']  # 取历史的FR007的值
                interest = interestall['FR007']  # 取上面计算得到的FR007的IRS收益率曲线
                facevalue = info[5]  # IRS的名义本金
                initialdate = info[1]  # IRS的起始日

                m = 0  # 对于每一个已经开始的利率互换，有且只有一个区间，它的即期利率要取收益率曲线上的最短点，远期利率要取历史值，必须单独处理这个特别的区间，处理之后令m=1

                for i in range(int(4 * info[3])):  # 1年对应4个付息日，5年对应4*5个付息日

                    dayb = initialdate + \
                        relativedelta(months=3 * i)  # 每次付息区间的开始日
                    while not cal.isworkday(dayb):
                        dayb = dayb + datetime.timedelta(days=1)

                    daye = initialdate + \
                        relativedelta(months=3 * (i + 1))  # 每次付息区间的结束日
                    while not cal.isworkday(daye):
                        daye = daye + datetime.timedelta(days=1)
                    daycal = (daye - dayb).days  # 付息区间的天数
                    # 付息日与估计日的差，大于0则说明还没有付息，要进入下面付息计算，小于0说明已经付完息，不用算了
                    daylast = (daye - assessmentdate).days
                    momeyflall = 0  # 每个付息区间的累计浮动利息

                    if daylast > 0:  # 大于0这个付息区间还没有付息
                        j = 0  # 用while和j+的组合做循环
                        dayy = -1  # 每个计息区间终止日与付息区间终止日的差值，小于0则这个付息区间还有计息区间，要计算剩下的计息区间，大于0则说明所有计息区间都算完了。一开始肯定有计息区间没算，所以初始值为-1，让循环开始

                        while dayy < 0:  # 小于0即这次计息区间已经结束，远期利率用历史值
                            daybb = dayb + \
                                datetime.timedelta(days=7 * j)  # 计息区间的开始日
                            dayee = dayb + \
                                datetime.timedelta(days=7 * (j + 1))  # 计息区间的结束日
                            daylastt = (dayee - assessmentdate).days
                            # 计息区间结束日与估值日的差值，大于0就要算即期利率，小于0就不用算
                            dayyy = (dayee - assessmentdate).days

                            if dayyy < 0:
                                ispot = 0
                                for k in interesthistory:
                                    if (k[0] - daybb).days >= 0:
                                        break
                                    ifar = k[1]  # 远期利率为计息日的前一个工作日

                                momeyfl = (facevalue + momeyflall) * \
                                    ifar * (dayee - daybb).days / 365
                                lista = [
                                    daybb, dayee, dayee, ispot, ifar, momeyfl, dayyy]
                                listaa.append(lista)
                                momeyflall += momeyfl

                            elif m == 0:
                                ispot = interest['1W'][4]
                                for k in interesthistory:
                                    if (k[0] - daybb).days >= 0:
                                        break
                                    ifar = k[1]  # 远期利率为计息日的前一个工作日
                                momeyfl = (facevalue + momeyflall) * \
                                    ifar * (dayee - daybb).days / 365
                                lista = [
                                    daybb, dayee, dayee, ispot, ifar, momeyfl, dayyy]
                                listaa.append(lista)
                                momeyflall += momeyfl
                                m = 1
                            else:
                                dd1 = 1000
                                dd2 = -1000
                                for ii, kk in interest.items():
                                    dd = (dayee - kk[0]).days
                                    if dd > 0:
                                        if dd < dd1:
                                            dd1 = dd
                                            ii1 = ii
                                    elif dd < 0:
                                        if dd > dd2:
                                            dd2 = dd
                                            ii2 = ii
                                    elif dd == 0:
                                        dd1 = dd
                                        ii1 = ii
                                        dd2 = dd
                                        ii2 = ii
                                ispot1 = interest[ii1][4]
                                ispot2 = interest[ii2][4]
                                day1 = interest[ii1][0]
                                day2 = interest[ii2][0]
                                if (day2 - day1).days == 0:
                                    ratio = 0
                                else:
                                    ratio = (dayee - day1).days / \
                                        (day2 - day1).days
                                ispot = ispot1 + (ispot2 - ispot1) * ratio
                                ifar = (math.exp(
                                    ispot * daylastt / 365 - listaa[-1][3] * listaa[-1][6] / 365) - 1) * 365 / 7
                                if (dayee - daye).days > 0:
                                    momeyfl = (facevalue + momeyflall) * \
                                        ifar * (daye - daybb).days / 365
                                    lista = [
                                        daybb, daye, dayee, ispot, ifar, momeyfl, daylastt]

                                else:
                                    momeyfl = (facevalue + momeyflall) * \
                                        ifar * (dayee - daybb).days / 365
                                    lista = [
                                        daybb, dayee, dayee, ispot, ifar, momeyfl, daylastt]
                                listaa.append(lista)
                                momeyflall += momeyfl

                            dayy = (dayee - daye).days
                            j += 1

                        dd1 = 1000
                        dd2 = -1000
                        for ii, kk in interest.items():
                            dd = (daye - kk[0]).days
                            if dd >= 0:
                                if dd < dd1:
                                    dd1 = dd
                                    ii1 = ii
                            elif dd <= 0:
                                if dd > dd2:
                                    dd2 = dd
                                    ii2 = ii
                        ispot1 = interest[ii1][4]
                        ispot2 = interest[ii2][4]
                        day1 = interest[ii1][0]
                        day2 = interest[ii2][0]
                        if day1==day2:
                            ratio=1
                        else:
                            ratio = (daye - day1).days / (day2 - day1).days
                        ispot = ispot1 + (ispot2 - ispot1) * ratio
                        deflator = math.exp(-ispot * daylast / 365)
                        lista = [dayb, daye, daye, ispot, 0, 0, daylast]
                        listaa.append(lista)

                        momeyfixall = facevalue * daycal / 365 * info[4]
                        listb = [dayb, daye, momeyflall, momeyfixall, deflator]
                        listbb.append(listb)


                if output == 0:
                    print('>>>>', code, '的计息区间现金流')
                    for x in listaa:
                        print(x)
                    print('>>>>', code, '的付息区间现金流')
                    for x in listbb:
                        print(x)

            # 1.2计算LPR1Y的IRS的价值（待写）
            if info[0] == 'LPR1Y':
                interesthistory = self.interesthistory['LPR1Y']
                interest = interestall['LPR1Y']
                facevalue = info[5]
                initialdate = info[1]
                m = 0
                for i in range(4 * info[3]):
                    dayb = initialdate + relativedelta(months=3 * i)
                    while not cal.isworkday(dayb):
                        dayb = dayb + datetime.timedelta(days=1)

                    daye = initialdate + relativedelta(months=3 * (i + 1))
                    while not cal.isworkday(daye):
                        daye = daye + datetime.timedelta(days=1)
                    daycal = (daye - dayb).days  # 付息区间的天数
                    # 付息日与估计日的差，大于0则说明还没有付息，要进入下面付息计算，小于0说明已经付完息，不用算了
                    daylast = (daye - assessmentdate).days

                    if daylast > 0:
                        if m == 0:
                            ispot = interest['3M'][4]
                            for k in interesthistory:
                                if (k[0] - dayb).days >= 0:
                                    break
                                ifar = k[1]
                            momeyfl = facevalue * daycal / 360 * ifar
                            deflator = math.exp(-ispot * \
                                                (daye - assessmentdate).days / 365)
                            lista = [
                                dayb, daye, daye, ispot, ifar, momeyfl, daylast]
                            listaa.append(lista)
                            momeyfix = facevalue * daycal / 365 * info[4]
                            listb = [dayb, daye, momeyfl, momeyfix, deflator]
                            listbb.append(listb)
                            m = 1
                        else:
                            dd1 = 1000
                            dd2 = -1000
                            for ii, kk in interest.items():
                                dd = (daye - kk[0]).days
                                if dd > 0:
                                    if dd < dd1:
                                        dd1 = dd
                                        ii1 = ii
                                elif dd < 0:
                                    if dd > dd2:
                                        dd2 = dd
                                        ii2 = ii
                                elif dd == 0:
                                    dd1 = dd
                                    ii1 = ii
                                    dd2 = dd
                                    ii2 = ii
                            ispot1 = interest[ii1][4]
                            ispot2 = interest[ii2][4]
                            day1 = interest[ii1][0]
                            day2 = interest[ii2][0]
                            if (day2 - day1).days == 0:
                                ratio = 0
                            else:
                                ratio = (daye - day1).days / (day2 - day1).days
                            ispot = ispot1 + (ispot2 - ispot1) * ratio
                            ifar = (
                                math.exp(
                                    ispot * daylast / 365 - lista[3] * lista[6] / 365) - 1) * 360 / daycal
                            momeyfl = facevalue * daycal / 360 * ifar
                            deflator = math.exp(-ispot *
                                                (daye - assessmentdate).days / 365)
                            lista = [
                                dayb, daye, daye, ispot, ifar, momeyfl, daylast]
                            listaa.append(lista)
                            momeyfix = facevalue * daycal / 365 * info[4]
                            listb = [dayb, daye, momeyfl, momeyfix, deflator]
                            listbb.append(listb)


                if output == 0:
                    print('>>>>', code, '的计息区间现金流')
                    for x in listaa:
                        print(x)
                    print('>>>>', code, '的付息区间现金流')
                    for x in listbb:
                        print(x)

            # 1.3计算SHIBOR3M的IRS的价值（待写）

            # 2计算浮动端、固定端以及NPV
            pfloat = 0
            pfix = 0
            for x in listbb:
                pfloat += x[2] * x[4]
                pfix += x[3] * x[4]
            v = (pfloat - pfix) * info[2]

            valuedic[code] = [v, pfloat * info[2], -
                              pfix * info[2]]  # 单笔的NPV，浮动端，固定端
            value += v  # 所有IRS的NPV
            if info[0]=='LPR1Y':
                valuelpr1y+=v
            elif info[0]=='FR007':
                if info[3]==3/4:
                    valuefr007s9m += v
                elif info[3]==1:
                    valuefr007s1y += v
                elif info[3]==5:
                    valuefr007s5y += v







        if output == 11:
            print('------所有持仓IRS的估值------')
            print('所有:',value)
            print('fr007s9m:',valuefr007s9m)
            print('fr007s1y:',valuefr007s1y)
            print('fr007s5y:',valuefr007s5y)
            print('lpr1y:',valuelpr1y)
            print('------每个IRS的估值[NPV，浮动端，固定端]------')
            for x, y in valuedic.items():
                print(x, ':', y)

        return [value, valuedic,[valuefr007s9m,valuefr007s1y,valuefr007s5y,valuelpr1y]]

    def dvbp(self, output=0):
        change1 = [
            [
                'FR007', '1W', 0.00005], [
                'FR007', '1M', 0.00005], [
                'FR007', '3M', 0.00005], [
                    'FR007', '6M', 0.00005], [
                        'FR007', '9M', 0.00005], [
                            'FR007', '1Y', 0.00005], [
                                'FR007', '2Y', 0.00005], [
                                    'FR007', '3Y', 0.00005], [
                                        'FR007', '4Y', 0.00005], [
                                            'FR007', '5Y', 0.00005], [
                                                'LPR1Y', '3M', 0.00005], [
                                                    'LPR1Y', '6M', 0.00005], [
                                                        'LPR1Y', '9M', 0.00005], [
                                                            'LPR1Y', '1Y', 0.00005], [
                                                                'LPR1Y', '2Y', 0.00005], [
                                                                    'LPR1Y', '3Y', 0.00005], [
                                                                        'LPR1Y', '4Y', 0.00005], [
                                                                            'LPR1Y', '5Y', 0.00005]]
        self.interestline(change1)
        mup = self.valuecal(1)
        change2 = [['FR007', '1W', -
                    0.0001], ['FR007', '1M', -
                              0.0001], ['FR007', '3M', -
                                        0.0001], ['FR007', '6M', -
                                                  0.0001], ['FR007', '9M', -
                                                            0.0001], ['FR007', '1Y', -
                                                                      0.0001], ['FR007', '2Y', -
                                                                                0.0001], ['FR007', '3Y', -
                                                                                          0.0001], ['FR007', '4Y', -
                                                                                                    0.0001], ['FR007', '5Y', -
                                                                                                              0.0001], ['LPR1Y', '3M', -
                                                                                                                        0.0001], ['LPR1Y', '6M', -
                                                                                                                                  0.0001], ['LPR1Y', '9M', -
                                                                                                                                            0.0001], ['LPR1Y', '1Y', -
                                                                                                                                                      0.0001], ['LPR1Y', '2Y', -
                                                                                                                                                                0.0001], ['LPR1Y', '3Y', -
                                                                                                                                                                          0.0001], ['LPR1Y', '4Y', -
                                                                                                                                                                                    0.0001], ['LPR1Y', '5Y', -
                                                                                                                                                                                              0.0001]]
        self.interestline(change2)
        mdown = self.valuecal(1)
        dvbpall = mup[0] - mdown[0]
        dvbpfr007s9m=mup[2][0] - mdown[2][0]
        dvbpfr007s1y = mup[2][1] - mdown[2][1]
        dvbpfr007s5y = mup[2][2] - mdown[2][2]
        dvbplpr1y = mup[2][3] - mdown[2][3]
        dvbplist = {}
        for x in self.asset.keys():

            dvbpi0 = mup[1][x][0] - mdown[1][x][0]
            dvbpi1 = mup[1][x][1] - mdown[1][x][1]
            dvbpi2 = mup[1][x][2] - mdown[1][x][2]
            dvbplist[x] = [dvbpi0, dvbpi1, dvbpi2]
        self.interestline(change1)

        if output == 0:
            print('------整个组合的DV01------')

            print('所有', dvbpall)
            print('fr007s9m', dvbpfr007s9m)
            print('fr007s1y', dvbpfr007s1y)
            print('fr007s5y', dvbpfr007s5y)
            print('lpr1y', dvbplpr1y)
            print('------每个IRS的DV01[NPV，浮动端，固定端]------')
            for x, y in dvbplist.items():
                print(x, y)

    def stresstest(self):
        print('-------------压力测试-------------')
        change = [['FR007', '1W', 0], ['FR007', '1M', 0], ['FR007', '3M', 0],
                  ['FR007', '6M', 0],
                  ['FR007', '9M', 0], ['FR007', '1Y', 0], ['FR007', '2Y', 0],
                  ['FR007', '3Y', 0],
                  ['FR007', '4Y', 0], ['FR007', '5Y', 0], ['LPR1Y', '3M', 0],
                  ['LPR1Y', '6M', 0],
                  ['LPR1Y', '9M', 0], ['LPR1Y', '1Y', 0], ['LPR1Y', '2Y', 0],
                  ['LPR1Y', '3Y', 0],
                  ['LPR1Y', '4Y', 0], ['LPR1Y', '5Y', 0]]
        # 1计算平行上移250BP
        for i in change:
            i[2] = 0.025
        self.interestline(change)
        value = self.valuecal(1)

        print('-------------1平行上移250BP-------------')
        print('市场总价值')
        print('所有',value[0])
        print('fr007s9m', value[2][0])
        print('fr007s1y', value[2][1])
        print('fr007s5y', value[2][2])
        print('lpr1y', value[2][3])
        print('各个IRS')
        for ii, kk in value[1].items():
            print(ii, ':', kk)
        for i in change:
            i[2] = -0.025
        print('变动后曲线')
        for x, y in self.interest.items():
            print(x)
            for a, c in y.items():
                print(a, c)
        self.interestline(change)

        # 2计算平行下移250BP
        for i in change:
            i[2] = -0.025
        self.interestline(change)
        value = self.valuecal(1)

        print('-------------2平行下移250BP-------------')
        print('市场总价值')
        print('所有', value[0])
        print('fr007s9m', value[2][0])
        print('fr007s1y', value[2][1])
        print('fr007s5y', value[2][2])
        print('lpr1y', value[2][3])
        print('各个IRS')
        for ii, kk in value[1].items():
            print(ii, ':', kk)
        for i in change:
            i[2] = 0.025
        print('变动后曲线')
        for x, y in self.interest.items():
            print(x)
            for a, c in y.items():
                print(a, c)
        self.interestline(change)

        # 3变陡峭
        use = self.interest
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = (-195 * math.exp(-t / 4) + 135 *
                    (1 - math.exp(-t / 4))) / 10000
        self.interestline(change)
        value = self.valuecal(1)

        print('-------------3变陡峭-------------')
        print('市场总价值')
        print('所有', value[0])
        print('fr007s9m', value[2][0])
        print('fr007s1y', value[2][1])
        print('fr007s5y', value[2][2])
        print('lpr1y', value[2][3])
        print('各个IRS')
        for ii, kk in value[1].items():
            print(ii, ':', kk)
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = -(-195 * math.exp(-t / 4) + 135 *
                     (1 - math.exp(-t / 4))) / 10000
        print('变动后曲线')
        for x, y in self.interest.items():
            print(x)
            for a, c in y.items():
                print(a, c)
        self.interestline(change)

        # 4变陡峭

        use = self.interest
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = (240 * math.exp(-t / 4) - 90 *
                    (1 - math.exp(-t / 4))) / 10000
        self.interestline(change)
        value = self.valuecal(1)

        print('-------------4变平缓-------------')
        print('市场总价值')
        print('所有', value[0])
        print('fr007s9m', value[2][0])
        print('fr007s1y', value[2][1])
        print('fr007s5y', value[2][2])
        print('lpr1y', value[2][3])
        print('各个IRS')
        for ii, kk in value[1].items():
            print(ii, ':', kk)
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = -(240 * math.exp(-t / 4) - 90 *
                     (1 - math.exp(-t / 4))) / 10000
        print('变动后曲线')
        for x, y in self.interest.items():
            print(x)
            for a, c in y.items():
                print(a, c)
        self.interestline(change)

        # 5短期利率向上移动
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = (300 * math.exp(-t / 4)) / 10000
        self.interestline(change)
        value = self.valuecal(1)

        print('-------------5短期利率向上移动-------------')
        print('市场总价值')
        print('所有', value[0])
        print('fr007s9m', value[2][0])
        print('fr007s1y', value[2][1])
        print('fr007s5y', value[2][2])
        print('lpr1y', value[2][3])
        print('各个IRS')
        for ii, kk in value[1].items():
            print(ii, ':', kk)
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = -(300 * math.exp(-t / 4)) / 10000
        print('变动后曲线')
        for x, y in self.interest.items():
            print(x)
            for a, c in y.items():
                print(a, c)
        self.interestline(change)

        # 6短期利率向上移动
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = -(300 * math.exp(-t / 4)) / 10000
        self.interestline(change)
        value = self.valuecal(1)

        print('-------------6短期利率向下移动-------------')
        print('市场总价值')
        print('所有', value[0])
        print('fr007s9m', value[2][0])
        print('fr007s1y', value[2][1])
        print('fr007s5y', value[2][2])
        print('lpr1y', value[2][3])
        print('各个IRS')
        for ii, kk in value[1].items():
            print(ii, ':', kk)
        for i in change:
            t = use[i[0]][i[1]][2] / 365
            i[2] = (300 * math.exp(-t / 4)) / 10000
        print('变动后曲线')
        for x, y in self.interest.items():
            print(x)
            for a, c in y.items():
                print(a, c)
        self.interestline(change)


b = irsfunc('20200630')
"""
b.interestline123('C:/Users/zyzse/Desktop/interestaddress.xlsx')
b.irsput('C:/Users/zyzse/Desktop/irs交易查询与维护.xlsx')
m=b.valuecal()
for x,y in b.interest['FR007'].items():
    print(x,y)
print(b.asset.items())
for x,y in b.asset.items():
    print(x,y)
print(m)
"""
b.interestread('C:/Users/zyzse/Desktop/interestaddress.xlsx')
change = [['FR007', '3M', 0.0001], ['FR007', '6M', 0.0001],
          ['FR007', '9M', 0.0001], ['FR007', '1Y', 0.0001]]


b.interestline()
b.irsput('C:/Users/zyzse/Desktop/irs交易查询与维护.xlsx')
print('------IRS的市场报价------')
for x, y in b.interestoriginal.items():
    print(x, y)
print('------基础利率的历史值------')
for x, y in b.interesthistory.items():
    print(x, y)
print('------各个IRS的收益率曲线------')
for x, y in b.interest.items():
    print(x)
    for a, c in y.items():
        print(a, c)

m = b.valuecal(11)

b.dvbp()
print('------持有的IRS信息------')
for x, y in b.asset.items():
    print(x, ':', y)

b.stresstest()
#进行压力测试的代码
print('----------测试信息-------------')

