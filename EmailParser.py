#!/usr/bin/env python
#-*- encoding: utf-8 -*-


import sys
import locale
import poplib
import string
import email

host = 'pop.163.com'
username = 'rayshining12@163.com'
password = 'Action1995120_'

# POPS3 over over an SSL encrypted socket, the standard port is 995
pop_con = poplib.POP3_SSL(host)
pop_con.user(username)
pop_con.pass_(password)

print(pop_con.list(), messages)

pop_con.quit()

# messages = [pop_con.retr(i) for i in range(1, len(pop_con.list()[1]) + 1)]
# print(pop_con.list(), messages)
# print('-----------------------------------------------------------------')
# messages = ['\n'.join(msg[1]) for msg in messages]
# print(messages)
# print('-----------------------------------------------------------------')

# messages = [email.parser.Parser().parsestr(msg) for msg in messages]

# i = 0
# for index in range(0, len(messages)):
#   message = messages[index]
#   i = i + 1
#   subject = message.get('subject')
#   header =
