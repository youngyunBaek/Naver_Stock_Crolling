import crolling_run
import excel_run
import timer_run
import threading
import time
from multiprocessing import Process, Queue

pre_check = 2

path_driver = "C:/chromedriver/chromedriver"
# path_driver = "D:/chromedriver/chromedriver"

# path = "D:/OneDrive/Python_private/"
# path_old = "D:/OneDrive/Python_private/"
# path_result = "D:/OneDrive/Python_private/"

# path = "F:/OneDrive/Python_private/"
# path_old = "F:/OneDrive/Python_private/"
# path_result = "F:/OneDrive/Python_private/"

path = "C:/OneDrive/Python_private/"
path_old = "C:/OneDrive/Python_private/"
path_result = "C:/OneDrive/Python_private/"

cr = crolling_run.Crolling()
xr = excel_run.ExcelCompare()
tr = timer_run.Timer()

#cr.data_down(path_driver, path)
#cr.data_save(path, path_old)

while xr.data_check(path) == 1:
    print("data false")
    cr.data_down(path_driver, path)

while 1:
    check = tr.time_check()
    month, day, week, hour, min, sec = tr.time_read()
    if len(str(hour)) == 1:
        hour = "0" + str(hour)
    if len(str(min)) == 1:
        min = "0" + str(min)
    if len(str(day)) == 1:
        day = "0" + str(day)
    dtime = str(month)+str(day)
    ttime = str(hour)+str(min)
    if pre_check != check:
        if check in [0,1,4,5,6]:
            print("")
            print(check, " ", dtime, " ", ttime)

            if check in [0, 1]:
                t1 = threading.Thread(target=cr.data_down, args=(path_driver, path))
                t1.start()
                threading.Thread(target=cr.KOSPI_data_down, args=(path_driver, path)).start()
                threading.Thread(target=cr.KOSDAQ_data_down, args=(path_driver, path)).start()
                threading.Thread(target=cr.KOSPI200_data_down, args=(path_driver, path)).start()
            if check in [1, 4]:
                threading.Thread(target=cr.Future_data_down, args=(path_driver, path)).start()
            if check in [5]:
                threading.Thread(target=cr.ETC2_data_down, args=(path_driver, path)).start()

            if check in [0, 1]:
                t1.join()

            while xr.data_check(path) == 1:
                print("data false")
                cr.data_down(path_driver, path)

            if check in [0, 1]:
                xr.excel_compare(path, path_old, path_result)

            time.sleep(10)

            if check in [6]:
                xr.compact_kospi(path)
                xr.compact_kodaq(path)
                xr.compact_kospi200(path)
                xr.compact_etc2(path)
                xr.compact_future(path)
                xr.compact_data(path)
                xr.day_stock_excel_compare(path, path_old, path_result)
                cr.data_copy(dtime, path)
            else:
                cr.data_save_time(ttime, path, path_old, check)

        pre_check = check
