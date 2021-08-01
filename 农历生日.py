
import sxtwl
from datetime import datetime
# 来源：https://blog.csdn.net/qq_44907926/article/details/104582586?spm=1001.2014.3001.5501
# 感觉很好用，很好玩，留下了备用
lunar = sxtwl.Lunar()  

def fun1(day1):
    a = datetime.now().year
    date1 = datetime.strptime(day1,'%Y-%m-%d')
    day = lunar.getDayByLunar(a, date1.month, date1.day)
    day_diernian = lunar.getDayByLunar(a+1, date1.month, date1.day)
    print("你的生日在本年的阳历生日为:", day.y, "年", day.m, "月", day.d, "日")
    print("你的生日在第二年的阳历生日为:", day_diernian.y, "年", day_diernian.m, "月", day_diernian.d, "日")
    
    global year
    global month
    global day11
    year = day.y
    month = day.m
    day11 = day.d

    global month_beiyong
    global day11_beiyong
    month_beiyong = day_diernian.m
    day11_beiyong = day_diernian.d

while True:
    if __name__ == '__main__':
        print("你可以输入'q'退出去哦！")
        day1 = input('请输入你的农历生日\n（格式为：年-月-日）:')
            
        if day1 == 'q':
            break    
                  
        fun1(day1)





    b = datetime.now()       
    date2 = datetime(year,month,day11)
    days1 = date2 - b
    j = days1.days
    if j >= 0 :
        print('你本年的生日还没有过哎，距离你本年的生日还差%s天'%days1)
    elif j<0:
        #现在到第二年1月1号的天数
        year += 1
        days4 = datetime(year,1,1)
        days3 = days4 - b
        f = days3.days

        #再从第二年11日计算到你生日的天数
        days5 = datetime(year,month_beiyong,day11_beiyong)
        days6 = days5 - days4
        g = days6.days
        
        #到第二年的总天数
        sum1 = f + g
        print('你今年的生日已经过了哎，距离你第二年的生日还有%s天' % sum1)






