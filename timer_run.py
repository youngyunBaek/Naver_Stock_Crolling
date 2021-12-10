import datetime

class Timer:
    def __int__(self):
        return

    def time_read(self):
        now = datetime.datetime.now()
        month = now.month
        day = now.day
        week = now.weekday()
        hour = now.hour
        min = now.minute
        sec = now.second
        return month, day, week, hour, min, sec

    def time_check(self):
        now = datetime.datetime.now()
        week = now.weekday()
        hour = now.hour
        min = now.minute
        sec = now.second
        check = 3
        # retrun 0: time[30min] 1: day
        if sec in [1]:
            if week in [0, 1, 2, 3, 4, 5]:
                if min in [0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55]:
                    if hour in [8, 9, 10, 11, 12, 13, 14, 15, 16] and week in [0, 1, 2, 3, 4]:
                        if min in [0, 5]:
                            time = str(hour)+"0"+str(min)
                        else:
                            time = str(hour)+str(min)
                        time = int(time)
                        if (time <= 855) or (time >= 1545):
                            check = 3
                        else:
                            check = 0
                        if (time > 915) and (time < 1605):
                            check = check + 1
                    elif hour in [22, 23, 0, 1, 2, 3, 4, 5, 6]:
                        if min in [0, 5]:
                            time = str(hour)+"0"+str(min)
                        else:
                            time = str(hour)+str(min)
                        time = int(time)
                        if (time >= 2330) or (time <= 620):
                            if (time >= 2330) and week in [0, 1, 2, 3, 4]:
                                check = 5
                            if (time <= 630) and week in [1, 2, 3, 4, 5]:
                                check = 5
                    elif hour in [17] and week in [0, 1, 2, 3, 4]:
                        if min in [0, 5]:
                            time = str(hour)+"0"+str(min)
                        else:
                            time = str(hour)+str(min)
                        time = int(time)
                        if time == 1700:
                            check = 6
        return check
