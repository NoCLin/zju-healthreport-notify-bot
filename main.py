# -*- coding: UTF-8 -*-
__author__ = "ZJUCST"
__copyright__ = "copyleft"
__license__ = "WTFPL"
__version__ = "1.0"

# 代码写得很随意 能用就行，懂得都懂
# By ZJUCST

import datetime
import json
import os
import shutil
import sys
from collections import defaultdict
from http.cookiejar import MozillaCookieJar

import requests
import xlrd
from apscheduler.schedulers.blocking import BlockingScheduler

import zju_login

CONFIG_FILE = os.path.join("data", "config.json")
EXCLUDES_FILE = os.path.join("data", "excludes.txt")
COOKIES_FILE = os.path.join("data", "cookies.txt")
RECORDS_DIR = os.path.join("data", "records")

# 请将账号密码写入 config.json

with open(CONFIG_FILE, "rt", encoding="utf-8") as f:
    config = json.loads(f.read())


# 简单的下载函数，未做retry
def download_file(sess, url, out_file):
    print('outputting to file: %s' % out_file)

    response = sess.get(url, stream=True)
    with open(out_file, 'wb') as f:
        response.raw.decode_content = True
        shutil.copyfileobj(response.raw, f)
    return out_file


def get_date_str():
    return datetime.datetime.now().strftime('%Y-%m-%d')


def get_datetime_str():
    return datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')


# 批量提醒
def send_ding_msg(person_list, ding_robot_url):
    print(person_list, ding_robot_url)

    # 提取需要@的手机号
    at_mobiles = [p["mobile"] for p in person_list]
    # 构造消息
    message = "以下同学请尽快健康打卡：" + "、".join([p["name"] for p in person_list])

    # 详见钉钉自定义群机器人文档
    post_json = {
        "msgtype": "text",
        "text": {
            "content": message,
        },
        "at": {
            "atMobiles": at_mobiles,
            "isAtAll": False
        }
    }

    r = requests.post(ding_robot_url, json=post_json)

    print("钉钉消息发送结果", r.json())

    return r.json()


def download_and_notify():
    print("开始提醒")

    # 年级和钉钉机器人URL的映射
    # 请在钉钉群里添加钉钉机器人，由于安全考虑，需要设置关键词，可以设置为每条消息都有的 "健康打卡"几个字
    grade_group_robot_mapping = config["grade_group_robot_mapping"]

    with open(EXCLUDES_FILE, "rt") as f:
        exclude_sid_list = f.read().strip().splitlines()

    print("正在下载文件")

    # 持久化cookies
    sess = requests.Session()
    cookies = MozillaCookieJar(COOKIES_FILE)
    if os.path.exists(COOKIES_FILE):
        cookies.load(ignore_discard=True, ignore_expires=True)
    sess.cookies = cookies

    # url = "https://healthreport.zju.edu.cn/ncov/wap/zju/export-download?group_id=1&group_type=1&type=weishangbao&date=2020-10-09"

    # 未上报报表下载url
    url = f"https://healthreport.zju.edu.cn/ncov/wap/zju/export-download?group_id=1&group_type=1&type=weishangbao&date={get_date_str()}"
    if not os.path.exists(RECORDS_DIR):
        os.mkdir(RECORDS_DIR)

    dest_file = os.path.join(RECORDS_DIR, f"{get_datetime_str()}.xlsx")

    # 下载文件
    record_file = download_file(sess, url, dest_file)
    print("文件下载完成", record_file)

    try:
        wb = xlrd.open_workbook(record_file)
    except:
        # 这里偷懒处理了.....打不开excel应该是没登录
        print("可能未登录")
        zju_login.login(sess, username=config["username"], password=config["password"])
        print("登录成功")
        print("重新下载")
        # 此处可能出错，如果出错需要重新执行一次
        record_file = download_file(sess, url, dest_file)
        wb = xlrd.open_workbook(record_file)

    # 读取下载好的excel信息
    sheet = wb.sheet_by_index(0)
    # 构造表头
    headers = dict((i, sheet.cell_value(0, i)) for i in range(sheet.ncols))
    # 提取信息
    values = list(
        (
            dict((headers[j], sheet.cell_value(i, j))
                 for j in headers)
            for i in range(1, sheet.nrows)
        )
    )

    group_by_grade = defaultdict(list)

    for row in values:
        name = row["姓名"]
        sid = row["学工号"]
        mobile = row["手机号码"]
        # 软件学院的需求是按年级发给不同的群，其他请院根据实际需求改
        # 本科生博士生年级提取规则也是[1:3]
        grade = sid[1:3]  # 从学号中提取年级（只适配了研究生的学号，如 22051001 20为年级）

        if sid in exclude_sid_list:
            print("跳过", name, sid, "(在排除列表中)")
            continue
        if not mobile:
            print("跳过", name, sid, "没有手机号信息")
            continue

        # 先把需要提醒的收集起来
        group_by_grade[grade].append({"name": name, "mobile": mobile})

    print(group_by_grade)

    # 开发者注：根据实际需求自己写一点点定制代码

    # for 软院 需求： 1718一个群 19 20单独群
    print("通知17 18级学生")
    # 这里是由于 17、18都是同一个群，就直接把人员合并了
    send_ding_msg(group_by_grade["17"] + group_by_grade["18"], grade_group_robot_mapping["18"])
    print("通知19级学生")
    send_ding_msg(group_by_grade["19"], grade_group_robot_mapping["19"])
    print("通知20级学生")
    send_ding_msg(group_by_grade["20"], grade_group_robot_mapping["20"])

    # for 某院 需求：全部学生在一个群
    # all_students = []
    # for k, v in group_by_grade.items():
    #     all_students += v
    #
    # print(all_students)
    # send_ding_msg(all_students, grade_group_robot_mapping["某院研究生"])


if __name__ == '__main__':
    if len(sys.argv) > 1:
        print("提醒一次")
        download_and_notify()
        exit()

    scheduler = BlockingScheduler()
    print("添加任务，每天15、18、21点提醒")

    # 添加cronjob

    scheduler.add_job(download_and_notify, 'cron', hour=15, minute=0)
    scheduler.add_job(download_and_notify, 'cron', hour=18, minute=0)
    scheduler.add_job(download_and_notify, 'cron', hour=21, minute=0)

    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        pass
