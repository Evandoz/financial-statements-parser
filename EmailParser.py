#!/usr/bin/env python
#-*- encoding: utf-8 -*-


import sys
import locale
import poplib
import string
import email

import tkinter
import tkinter.filedialog
import tkinter.messagebox

from tkinter import ttk


class ParseEmail(object):
  def __init__(self):
    self.root = tkinter.Tk()
    self.root.resizable(0, 0)
    self.center_main_window(640, 420)
    self.root.title('邮箱附件批量下载工具')

    self.top = ttk.Frame()
    s = ttk.Style()
    s.configure('E.TEntry', borderwidth=10, insertwidth=1, relief=tkinter.FLAT)
    self.email_entry = ttk.Entry(self.top, width=62, style='E.TEntry')
    self.parse_button = ttk.Button(self.top, text='登录并下载', width=10)

  def get_screen_size(self):
    return self.root.winfo_screenwidth(), self.root.winfo_screenheight()

  # 获取主窗口的大小，用于子窗口的定位

  def get_window_size(self):
    return self.root.winfo_reqwidth(), self.root.winfo_reqheight()

  # 获取主窗口的位置，用于子窗口的定位

  def get_window_pos(self):
    self.root.update()
    return self.root.winfo_x(), self.root.winfo_y()

  # 主窗口中心点位置

  def main_window_pos(self):
    window_x, window_y = self.get_window_pos()
    window_w, window_h = self.get_window_size()
    # main_center_x, main_center_y = window_x+window_w/2, window_y+window_h/2
    # 为了统一使用center_window函数，这里将main_center_x、main_center_y的结果同时乘以2
    return window_x*2+window_w, window_y*2+window_h

  # 窗口居中显示
  # 居中定位函数可能会被多个组件调用，因此引入window做参数，而不是self.root将其写死

  def center_window(self, window, width, height, screenwidth, screenheight):
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
    window.geometry(size)

  # 主窗口居中
  def center_main_window(self, width, height):
    screenwidth, screenheight = self.get_screen_size()
    self.center_window(self.root, width, height, screenwidth, screenheight)

  # 子窗口居中
  # 居中定位函数可能会被多个子窗口调用，因此引入window做参数，而不是self.root将其写死

  def center_child_window(self, window, width, height):
    screenwidth, screenheight = self.main_window_pos()
    self.center_window(window, width, height, screenwidth, screenheight)

  def show(self):
    self.top.pack(padx=12, pady=12)
    self.email_entry.pack(expand=tkinter.YES, side=tkinter.LEFT, fill=tkinter.Y, pady=1)
    self.parse_button.pack(expand=tkinter.YES, side=tkinter.LEFT, fill=tkinter.Y, padx=8, ipady=1)

    self.root.mainloop()


# host = 'pop.126.com'
# username = 'evando@126.com'
# password = 'Crist0207'

# # POPS3 over over an SSL encrypted socket, the standard port is 995
# pop_con = poplib.POP3_SSL(host)
# pop_con.user(username)
# pop_con.pass_(password)

# print(pop_con.list(), messages)

# pop_con.quit()

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


def main():
  PE = ParseEmail()
  PE.show()


if __name__ == '__main__':
  main()
