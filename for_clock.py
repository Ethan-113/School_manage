from asyncio import start_server
import threading
from datetime import datetime, timedelta
import time
import numpy as np
import logging
from pywebio.input import input
from pywebio.output import put_text, put_buttons, use_scope, put_markdown, toast

from Course_Activity import course_init
from Course_Activity import identity
from Course_Activity import get_destination
from MAP import map_init
from TemporaryEvents import temp_event_init
from TemporaryEvents import temporary_forevent
from TemporaryEvents import temporary_deleteAll
from Course_Activity import clock_ring
from Course_Activity import get_todaydata

log_file = '..\\Logs\\'  # 日志文件路径

# 全局变量，用于存储上一次的小时数和日期
previous_hour = None
previous_date = None

# 用来获取临时事件中需要进行提醒的事件
clock_list_temporary = []

# 进行闹钟时存储该小时内所有的事件的关键位置
way_clock_temporary = list()
way_clock_class = list()

# 用于检测是否是第一次打印日期
if_today = True

# 全局变量，用于存储当前的时间比，如果是60的话就代表现实时间过1s，模拟时间过60s
time_ratio = 60


# 操作线程
def get_user_input():
    time.sleep(1)
    put_text("请选择操作：")
    put_buttons(['课程管理', '临时事务管理', '时间加速比修改'], onclick=button_clicked)


def button_clicked(btn):
    if btn == '课程管理':
        logging.info('用户进入课程管理')
        course_init(current_time, identity_data)
    elif btn == '临时事务管理':
        logging.info('用户进入临时事务管理')
        temp_event_init(stu_num, num_class, month_today, day_today, week_today)
    elif btn == '时间加速比修改':
        speed_change()


def speed_change():
    global time_ratio
    speed = int(input("输入系统时间与现实时间的比值，如10代表现实时间每1s系统时间过10s，输入0暂停时间"))
    time_ratio = speed


if __name__ == '__main__':
    global month_today, day_today, week_today

    stu_num = 0

    identity_data = identity()

    if identity_data['if_stu'] == 1:
        # 获得学生相关信息
        stu_num = identity_data['stu_num']
        stu_class = identity_data['class']
        num_class = stu_class % 10

        # 个人日志文件路径
        log_file = log_file + "Log-" + str(stu_num) + ".txt"
        logging.basicConfig(filename=log_file, level=logging.DEBUG, format='%(asctime)s -- %(message)s')
        # 记录日志
        logging.info('学生登录系统')

        # 获取临时事件中有闹钟的事件
        clock_list_temporary = temporary_forevent(stu_num)

        current_time = datetime.now()  # 当前时间
        program_time = timedelta()  # 程序时间初始为0秒
        month_today = current_time.month
        day_today = current_time.day
        week_today = int(current_time.strftime("%w"))

        get_todaydata(current_time, identity_data)

        put_markdown('# 35组学生日程管理系统')

        # 检查用户输入，调用其他模块的线程
        thread_input = threading.Thread(target=get_user_input)
        thread_input.start()

        # 此处时间必须写在主线程里，因为涉及GUI库调用
        while True:
            # 刷新临时事务闹钟
            clock_list_temporary = temporary_forevent(stu_num)

            # 模拟程序时间流逝
            program_time += timedelta(seconds=1)  # 每次循环加1秒
            current_time += timedelta(seconds=time_ratio)  # 根据时间比更新当前时间

            # 检测日期变化，调用相关功能
            if current_time.date() != previous_date:

                previous_date = current_time.date()
                # 更新月和日和星期的全局变量
                month_today = current_time.month
                day_today = current_time.day
                week_today = int(current_time.strftime("%w"))

                # 打印当前日期
                with use_scope('day', clear=True):
                    put_markdown('## 当前日期' + current_time.strftime("%Y-%m-%d"))
                # 防止首次进入程序打印日期时清空临时事务表
                if if_today == False:
                    temporary_deleteAll(stu_num)
                if_today = False

            # 检测小时变化，调用相关功能
            if current_time.hour != previous_hour:
                previous_hour = current_time.hour
                # 用于存储地址的两个list先进行清空
                way_clock_temporary.clear()
                way_clock_class.clear()
                # 打印当前小时和分钟
                with use_scope('time', clear=True):
                    put_markdown('## 当前时间' + current_time.strftime("%H:%M"))
                # 获得当前小时内课程的地址
                temp = get_destination(current_time)
                for i in range(0, len(temp)):
                    now_class = np.array((temp[i][0], temp[i][1]))
                    way_clock_class.append(now_class)

                # 获得当前小时内临时事务的地址
                with use_scope('this_hour', clear=True):
                    put_markdown('#### 当前小时的临时事务')
                    for i in range(len(clock_list_temporary)):
                        if clock_list_temporary[i].getHour() == current_time.hour:
                            put_text(
                                "\n在本小时内需要去" + clock_list_temporary[i].getName() + "," + clock_list_temporary[
                                    i].getDescription())
                            now_temporary = np.array(
                                (clock_list_temporary[i].getPlace()[0], clock_list_temporary[i].getPlace()[1]))
                            way_clock_temporary.append(now_temporary)

                if len(way_clock_temporary) != 0:
                    map_init(way_clock_temporary, '临时事务路线与地图')

                if temp != [[0, 0]]:
                    map_init(way_clock_class, '课程路线与地图')

                # 调用第二天提醒课程功能，满足条件则调用
                clock_ring(current_time, identity_data)

            # 等待1秒
            time.sleep(1)

    # 依然是没有通用变量名
    elif identity_data['if_stu'] == 0:
        toast("检测为管理员账号，进入课程管理")
        current_time = datetime.now()  # 当前时间
        while True:
            course_init(current_time, identity_data)