# -*- coding: utf-8 -*-
import re

import requests


def _rsa_encrypt(password_str, e_str, M_str):
    password_bytes = bytes(password_str, 'ascii')
    password_int = int.from_bytes(password_bytes, 'big')
    e_int = int(e_str, 16)
    M_int = int(M_str, 16)
    result_int = pow(password_int, e_int, M_int)
    return hex(result_int)[2:].rjust(128, '0')


def login(sess: requests.Session, username, password, ):
    login_url = "https://zjuam.zju.edu.cn/cas/login?service=https%3A%2F%2Fhealthreport.zju.edu.cn%2Fa_zju%2Fapi%2Fsso%2Findex%3Fredirect%3Dhttps%253A%252F%252Fhealthreport.zju.edu.cn%252Fncov%252Fwap%252Fdefault%252Findex"

    res = sess.get(login_url)
    execution = re.search('name="execution" value="(.*?)"', res.text).group(1)
    res = sess.get(url='https://zjuam.zju.edu.cn/cas/v2/getPubKey').json()
    n, e = res['modulus'], res['exponent']
    encrypt_password = _rsa_encrypt(password, e, n)

    data = {
        'username': username,
        'password': encrypt_password,
        'execution': execution,
        '_eventId': 'submit'
    }
    res = sess.post(url=login_url, data=data)

    # FIXME: naive实现
    if '统一身份认证' in res.content.decode():
        raise Exception('登录失败，请核实账号密码重新登录')
    sess.cookies.save()
    return sess
