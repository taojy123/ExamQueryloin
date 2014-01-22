# -*- coding: cp936 -*-

import cookielib
import urllib2, urllib
import time
import re
import traceback
import xlwt
import xlrd

cj = cookielib.CookieJar()
opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cj))
opener.addheaders = [('User-agent', 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1700.76 Safari/537.36'),
                     ('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'), 
                     ('Accept-Language', 'zh-CN,zh;q=0.8,zh-TW;q=0.6,en;q=0.4'), 
                     ('Connection', 'keep-alive'),
                     ('Content-Type', 'application/x-www-form-urlencoded'),
                     ('Host', 'www.ynrsksw.cn'),
                     ('Origin', 'http://www.ynrsksw.cn')
                     ]
opener.addheaders.append( ('Accept-encoding', 'identity') )
opener.addheaders.append( ('Referer', 'http://www.ynrsksw.cn/Ynrsks/ExamQueryloin.aspx') )


def get_page(url, data=None):
    n = 0
    while n < 5:
        n = n + 1
        try:
            resp = opener.open(url, data, timeout=5)
            page = resp.read()
            return page
        except:
            traceback.print_exc()
            print "Will try after 2 seconds ..."
            time.sleep(2.0)
            continue
        break
    return "Null"

url = 'http://www.ynrsksw.cn/Ynrsks/ExamQueryloin.aspx'
vs = "/wEPDwULLTE3NjA1MDE4NDYPZBYCAgEPZBYEAgEPEA8WBh4ORGF0YVZhbHVlRmllbGQFDOiAg+ivleS7o+eggR4NRGF0YVRleHRGaWVsZAUM6ICD6K+V5ZCN56ewHgtfIURhdGFCb3VuZGdkEBVFKjIwMTPlubQxMeaciOe7j+a1juS4k+S4muaKgOacr+i1hOagvOiAg+ivlS0yMDEz5bm0MTDmnIjpgKDku7flt6XnqIvluIjmiafkuJrotYTmoLzogIPor5UwMjAxM+W5tDEw5pyI5LyB5Lia5rOV5b6L6aG+6Zeu5omn5Lia6LWE5qC86ICD6K+VJzIwMTPlubQxMOaciOaLm+agh+W4iOiBjOS4muawtOW5s+iAg+ivlTMyMDEz5bm0MTDmnIjms6jlhozln47luILop4TliJLluIjmiafkuJrotYTmoLzogIPor5UkMjAxM+W5tDEw5pyI5omn5Lia6I2v5biI6LWE5qC86ICD6K+VKjIwMTPlubQxMOaciOWHuueJiOS4k+S4muiBjOS4mui1hOagvOiAg+ivlSoyMDEz5bm0MTDmnIjlrqHorqHkuJPkuJrmioDmnK/otYTmoLzogIPor5UeMjAxM+W5tDEw5pyI5oi/5Zyw5Lqn57uP57qq5Lq6JjIwMTPlubQ55pyI54mp5Lia566h55CG5biI6LWE5qC86ICD6K+VLDIwMTPlubQ55pyI5LiA57qn5bu66YCg5biI5omn5Lia6LWE5qC86ICD6K+VJjIwMTPlubQ55pyI5LiA57qn5bu66YCg5biI55u45bqU5LiT5LiaGjIwMTPlubQ55pyI5rOo5YaM5rWL57uY5biIODIwMTPlubQ55pyI5LqR5Y2X55yB5Z+65bGC5pS/5rOV5py65YWz5a6a5ZCR5oub5b2V6ICD6K+VJjIwMTPlubQ55pyI5aSW6ZSA5ZGY5LuO5Lia6LWE5qC86ICD6K+VLDIwMTPlubQ55pyI5Zu96ZmF5ZWG5Yqh5biI5omn5Lia6LWE5qC86ICD6K+VLDIwMTPlubQ55pyI5Lu35qC86Ym06K+B5biI5omn5Lia6LWE5qC86ICD6K+VMjIwMTPlubQ55pyI5rOo5YaM6LWE5Lqn6K+E5Lyw5biI5omn5Lia6LWE5qC86ICD6K+VWTIwMTPlubQ55pyI5LqR5Y2X55yB6I2v5a2m77yI6Z2e5Li05bqK5Yy755aX77yJ5LiT5Lia5oqA5pyv77yI6I2v5aOr44CB5Lit6I2v5aOr77yJ6LWE5qC8WTIwMTPlubQ55pyI5LqR5Y2X55yB6I2v5a2m77yI6Z2e5Li05bqK5Yy755aX77yJ5LiT5Lia5oqA5pyv77yI6I2v5biI44CB5Lit6I2v5biI77yJ6LWE5qC8XzIwMTPlubQ55pyI5LqR5Y2X55yB6I2v5a2m77yI6Z2e5Li05bqK5Yy755aX77yJ5LiT5Lia5oqA5pyv77yI5Li7566h6I2v5biI44CB5Lit6I2v5biI77yJ6LWE5qC8IDIwMTPlubQ55pyI5rOo5YaM6K6+5aSH55uR55CG5biIMjIwMTPlubQ55pyI5rOo5YaM5a6J5YWo5bel56iL5biI5omn5Lia6LWE5qC86ICD6K+VLTIwMTPlubQ35pyIMjAxM+W5tOS/neWxseW4guS6i+S4muWNleS9jeiAg+ivlSkyMDEz5bm0N+aciOS6keWNl+WGnOS4muWkp+WtpuaLm+iBmOiAg+ivlS0yMDEz5bm0N+aciDIwMTPlubTlpKfnkIblt57kuovkuJrljZXkvY3ogIPor5UsMjAxM+W5tDbmnIjkuoznuqflu7rpgKDluIjmiafkuJrotYTmoLzogIPor5UsMjAxM+W5tDbmnIjms6jlhoznqI7liqHluIjmiafkuJrotYTmoLzogIPor5UgMjAxM+W5tDbmnIjkupHljZfnnIHkuInmlK/kuIDmibYpMjAxM+W5tDbmnIjotKjph4/kuJPkuJrogYzkuJrotYTmoLzogIPor5UsMjAxM+W5tDbmnIjnpL7kvJrlt6XkvZzogIXogYzkuJrmsLTlubPogIPor5UyMjAxM+W5tDbmnIjlnJ/lnLDnmbvorrDku6PnkIbkurrogYzkuJrotYTmoLzogIPor5UgMjAxM+W5tDbmnIjkuIDnuqfms6jlhozorqHph4/luIg1MjAxM+W5tDXmnIjlub/lkYrkuJPkuJrmioDmnK/kurrlkZjogYzkuJrmsLTlubPogIPor5UsMjAxM+W5tDXmnIjnjq/looPlvbHlk43or4Tku7flt6XnqIvluIjogIPor5UsMjAxM+W5tDXmnIjnm5HnkIblt6XnqIvluIjmiafkuJrotYTmoLzogIPor5VXMjAxM+W5tDXmnIgyMDEz5bm05LqR5Y2X55yB5Yac5p2R5L+h55So56S+5YWs5byA5oub6IGY6LS36K6w5Y2h55S16K+d6ZO26KGM5a6i5pyN5Lq65ZGYODIwMTPlubQ15pyI5oqV6LWE5bu66K6+6aG555uu566h55CG5biI6IGM5Lia5rC05bmz6ICD6K+VLDIwMTPlubQ15pyI566h55CG5ZKo6K+i5biI6IGM5Lia5rC05bmz6ICD6K+VMjIwMTPlubQ05pyI5rOo5YaM5ZKo6K+i5bel56iL5biI5omn5Lia6LWE5qC86ICD6K+VMDIwMTPlubQ05pyIMjAxM+W5tOS6keWNl+ecgeWFrOWKoeWRmOW9leeUqOiAg+ivlUIyMDEz5bm0NOaciOS6keWNl+ecgTIwMTPlubTpgInogZjpq5jmoKHmr5XkuJrnlJ/liLDmnZHku7vogYzlt6XkvZwjMjAxM+W5tDTmnIjogYznp7DlpJbor63nrYnnuqfogIPor5UuMjAxMuW5tDEw5pyIMjAxMuW5tOaYremAmuW4guS6i+S4muWNleS9jeiAg+ivlSAyMDEy5bm0OeaciOS6jOe6p+azqOWGjOiuoemHj+W4iFMyMDEy5bm0OeaciOaYhuaYjuW4guS4reWxguW5sumDqO+8iOenkee6p++8iei3qOmDqOmXqOernuWyl+S6pOa1geiAg+ivleeslOivleaIkOe7qRoyMDEx5bm0OeaciOmrmOe6p+S8muiuoeW4iGQyMDEw5bm0MTDmnIgyMDEw5bm05LqR5Y2X55yB5py65YWz5Y+K5Y+C5YWs5Y2V5L2N5q2j56eR57qn5bmy6YOo5pmL5Y2H5Ymv5Y6/5aSE57qn6IGM5Yqh6LWE5qC86ICD6K+VQDIwMTDlubQ55pyI5rOo5YaM5Zyf5pyo5bel56iL5biIKOawtOWIqeawtOeUteW3peeoi+awtOWcn+S/neaMgSk6MjAxMOW5tDnmnIjms6jlhozlnJ/mnKjlt6XnqIvluIgo5rC05Yip5rC055S15bel56iL56e75rCRKToyMDEw5bm0OeaciOazqOWGjOWcn+acqOW3peeoi+W4iCjmsLTliKnmsLTnlLXlt6XnqIvlnLDotKgpLjIwMTDlubQ55pyI5rOo5YaM5Zyf5pyo5bel56iL5biIKOawtOW3pee7k+aehCk6MjAxMOW5tDnmnIjms6jlhozlnJ/mnKjlt6XnqIvluIgo5rC05Yip5rC055S15bel56iL6KeE5YiSKSAyMDEw5bm0OeaciOazqOWGjOeOr+S/neW3peeoi+W4iC4yMDEw5bm0OeaciOazqOWGjOeUteawlOW3peeoi+W4iCjlj5HovpPlj5jnlLUpKzIwMTDlubQ55pyI5rOo5YaM55S15rCU5bel56iL5biIKOS+m+mFjeeUtSk0MjAxMOW5tDnmnIjms6jlhozlhaznlKjorr7lpIflt6XnqIvluIgo57uZ5rC05o6S5rC0KTQyMDEw5bm0OeaciOazqOWGjOWFrOeUqOiuvuWkh+W3peeoi+W4iCjmmpbpgJrnqbrosIMpLjIwMTDlubQ55pyI5rOo5YaM5YWs55So6K6+5aSH5bel56iL5biIKOWKqOWKmykgMjAxMOW5tDnmnIjms6jlhozljJblt6Xlt6XnqIvluIg3MjAxMOW5tDnmnIjms6jlhozlnJ/mnKjlt6XnqIvluIgo5riv5Y+j5LiO6Iiq6YGT5bel56iLKTIyMDEw5bm0OeaciOS4gOe6p+azqOWGjOe7k+aehOW3peeoi+W4iOi1hOagvOiAg+ivlTIyMDEw5bm0OeaciOS6jOe6p+azqOWGjOe7k+aehOW3peeoi+W4iOi1hOagvOiAg+ivlT4yMDEw5bm0OeaciOazqOWGjOWcn+acqOW3peeoi+W4iO+8iOWyqeWcn++8ieaJp+S4mui1hOagvOiAg+ivlUcyMDEw5bm0NuaciOS6keWNl+ecgeS6i+S4muWNleS9jeWumuWQkeaLm+iBmOWIsOWGnOadkeWfuuWxguacjeWKoemhueebrjIyMDEw5bm0NeaciOS4gOe6p+azqOWGjOW7uuetkeW4iOaJp+S4mui1hOagvOiAg+ivlSwyMDEw5bm0NeaciOS6jOe6p+azqOWGjOW7uuetkeW4iOi1hOagvOiAg+ivlTwyMDA55bm0MTHmnIjkuozjgIHkuInnuqfnv7vor5HkuJPkuJrotYTmoLzvvIjmsLTlubPvvInogIPor5UwMjAwN+W5tDEx5pyI5LiA57qn5Zyw6ZyH5a6J5YWo5oCn6K+E5Lu35bel56iL5biIFUUCMDECMjMCMzACMzkCMjkCMjYCMDcCNjQCODgCNDACMzQCNzkCNzICOTUCMDYCMDICMjQCMjUCNTkCNjACNjECMzYCMzMCNjcCOTYCNjMCNjICMjICOTECMDUCMzgCMzICNzACNjgCMzUCMjECOTICMzcCMTgCMzECOTkCOTcCODECNTQCNzECNTcCODcCOTgCNzMCNzQCNzUCNzYCNzcCNzgCMTECMTICMTMCMTQCMTUCMTYCMTcCMDMCMDQCMDgCNTICMjcCMjgCMDkCMTkUKwNFZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZGQCCQ8PFgIeBFRleHQFG+iAg+eUn+acquWPguWKoOivpeasoeiAg+ivlWRkGAEFHl9fQ29udHJvbHNSZXF1aXJlUG9zdEJhY2tLZXlfXxYBBQlidG5SZXN1bHQ2zECOqa4oVtU74bPyBJojY+zCag=="
ev = "/wEWSwKW3t3wBgLN2fyaBALT2cSaBALQ2fCaBALQ2ZyZBALT2ZyZBALT2ciaBALN2dSaBALX2cCaBALF2ZCZBALR2fCaBALQ2cCaBALU2ZyZBALU2fiaBALK2cyaBALN2ciaBALN2fiaBALT2cCaBALT2cyaBALW2ZyZBALX2fCaBALX2fyaBALQ2ciaBALQ2cSaBALX2dSaBALK2ciaBALX2cSaBALX2fiaBALT2fiaBALK2fyaBALN2cyaBALQ2ZCZBALQ2fiaBALU2fCaBALX2ZCZBALQ2cyaBALT2fyaBALK2fiaBALQ2dSaBALS2ZCZBALQ2fyaBALK2ZyZBALK2dSaBALF2fyaBALW2cCaBALU2fyaBALW2dSaBALF2dSaBALK2ZCZBALU2cSaBALU2cCaBALU2cyaBALU2ciaBALU2dSaBALU2ZCZBALS2fyaBALS2fiaBALS2cSaBALS2cCaBALS2cyaBALS2ciaBALS2dSaBALN2cSaBALN2cCaBALN2ZCZBALW2fiaBALT2dSaBALT2ZCZBALN2ZyZBALS2ZyZBALEhISFCwK58bqdCQKzkqT7CQKskqT7CQKK2+T+DIXaGIuQcyJqDaiGXZEkpaHrSEPF"

#name = "安强"
#idnum = "530181196901274039"

data = xlrd.open_workbook('data.xls')
table = data.sheets()[0]

for i in range(1, table.nrows):
    name = table.cell(i, 0).value.encode("gbk")
    idnum = str(table.cell(i, 1).value)
    phone = str(table.cell(i, 8).value)


    formData = urllib.urlencode({'__VIEWSTATE' : vs,
                                 'dplExam' : "62",
                                 'txtName' : name,
                                 'txtIDNum' : idnum,
                                 'dplRegType' : "0",
                                 'btnResult.x' : "53",
                                 'btnResult.y' : "11",
                                 "__EVENTVALIDATION" : ev
                                 })

    print "get:" + name
    p = get_page(url, formData)

    #print p

    if "成绩表" in p:
        print name, idnum, phone, "========================================"


input()
