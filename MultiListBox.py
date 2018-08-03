#-*-coding:utf-8-*-

import tkinter
from tkinter import Frame, Label, Listbox, Scrollbar

MOVE_LINES = 0
MOVE_PAGES = 1
MOVE_TOEND = 2

class MultiListBox(Frame):
  def __init__(self, master, lists, **kw):
    Frame.__init__(self, master)
    self.lists = []
    # 多列ListBox并置
    h = kw['height']
    for l,w in lists:
        frame = Frame(self)
        frame.pack(side=tkinter.LEFT, expand=tkinter.YES, fill=tkinter.BOTH)
        Label(frame, text=l, width=w, borderwidth=0, background='#fff',
              relief=tkinter.FLAT).pack(expand=tkinter.YES, padx=1, ipady=4, fill=tkinter.BOTH)
        lb = Listbox(frame, width=w, height=h, borderwidth=0, selectborderwidth=0,
                     relief=tkinter.FLAT, exportselection=tkinter.FALSE)
        lb.pack(expand=tkinter.YES, ipadx=5, fill=tkinter.BOTH)
        self.lists.append(lb)
        lb.bind('<B1-Motion>', lambda e, s=self: s._select(e.y))
        lb.bind('<Button-1>', lambda e, s=self: s._select(e.y))
        lb.bind('<Leave>', lambda e: 'break')
        lb.bind('<B2-Motion>', lambda e, s=self: s._b2motion(e.x, e.y))
        lb.bind('<Button-2>', lambda e, s=self: s._button2(e.x, e.y))

    # self.bind('<up>',    lambda e, s=self: s._move (-1, MOVE_LINES))
    # self.bind('<down>',  lambda e, s=self: s._move (+1, MOVE_LINES))
    # self.bind('<prior>', lambda e, s=self: s._move (-1, MOVE_PAGES))
    # self.bind('<next>',  lambda e, s=self: s._move (+1, MOVE_PAGES))
    # self.bind('<home>',  lambda e, s=self: s._move (-1, MOVE_TOEND))
    # self.bind('<end>',   lambda e, s=self: s._move (+1, MOVE_TOEND))

    frame = Frame(self)
    frame.pack(side=tkinter.LEFT, fill=tkinter.Y)
    # Label(frame, borderwidth=0, relief=tkinter.FLAT).pack(padx=1, fill=tkinter.X)
    sb = Scrollbar(frame, orient=tkinter.VERTICAL, command=self._scroll)
    sb.pack(expand=tkinter.YES, fill=tkinter.Y)
    self.lists[0]['yscrollcommand']=sb.set


  def _move (self, lines, relative=0):
    """
    Move the selection a specified number of lines or pages up or
    down the list.  Used by keyboard navigation.
    """
    selected = self.lists[0].curselection()
    try:
        selected = list(map(int, selected))
    except ValueError:
        pass

    try:
        sel = selected[0]
    except IndexError:
        sel = 0

    old  = sel
    size = self.lists [0].size()

    if relative == MOVE_LINES:
        sel = sel + lines
    elif relative == MOVE_PAGES:
        sel = sel + (lines * int (self.lists [0]['height']))
    elif relative == MOVE_TOEND:
        if lines < 0:
            sel = 0
        elif lines > 0:
            sel = size - 1
    else:
        print("MultiListbox._move: Unknown move type!")

    if sel < 0:
        sel = 0
    elif sel >= size:
        sel = size - 1

    self.selection_clear (old, old)
    self.see (sel)
    self.selection_set (sel)
    return 'break'

  def _select(self, y):
    row = self.lists[0].nearest(y)
    self.selection_clear(0, tkinter.END)
    self.selection_set(row)
    self.focus_force()
    return 'break'

  def _button2(self, x, y):
    for l in self.lists:
        l.scan_mark(x, y)
    return 'break'

  def _b2motion(self, x, y):
    for l in self.lists:
        l.scan_dragto(x, y)
    return 'break'

  def _scroll(self, *args):
    for l in self.lists:
        l.yview(*args)
    return 'break'

  def curselection(self):
    return self.lists[0].curselection()

  def itemconfigure(self, row_index, col_index, cnf=None, **kw):
    lb = self.lists[col_index]
    return lb.itemconfigure(row_index, cnf, **kw)

  def rowconfigure(self, row_index, cnf={}, **kw):
    for lb in self.lists:
        lb.itemconfigure(row_index, cnf, **kw)

  def delete(self, first, last=None):
    for l in self.lists:
        l.delete(first, last)

  def get(self, first, last=None):
    result = []
    for l in self.lists:
        result.append(l.get(first,last))
    #if last:
    #    return map(*([None] + result))
    return result

  def index(self, index):
    self.lists[0].index(index)

  def insert(self, index, *elements):
    for e in elements:
        i = 0
        for l in self.lists:
            l.insert(index, e[i])
            i = i + 1

  def size(self):
    return self.lists[0].size()

  def see(self, index):
    for l in self.lists:
        l.see(index)

  def selection_anchor(self, index):
    for l in self.lists:
      l.selection_anchor(index)

  def selection_clear(self, first, last=None):
    for l in self.lists:
      l.selection_clear(first, last)

  def selection_includes(self, index):
    return self.lists[0].selection_includes(index)

  def selection_set(self, first, last=None):
    for l in self.lists:
      l.selection_set(first, last)

  def yview_scroll(self, *args, **kwargs):
    for lb in self.lists:
      lb.yview_scroll(*args, **kwargs)
