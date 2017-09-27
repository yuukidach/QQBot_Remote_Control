#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime
from course_schedule import get_next_course

username = ''

def onQQMessage(bot, contact, memeber, content):
    global username
    if username == '':
        if content == 'dash':
            bot.SendTo(contact, 'At your service!')
            username = 'dash'
        else:
            bot.SendTo(contact, 'Sorry. Only work for my master.   :(')

    elif username != '':
        if content == "下节课":
            nntime = datetime.datetime.now()
            nnweek = nntime.isoweekday()
            result = get_next_course(nnweek, nntime, '/home/dash/Documents/dash_courses.xlsx')
            bot.SendTo(contact, result)
        elif content == "好的":
            username = ''
            bot.SendTo(contact, 'Glad to serve you!   :D')
