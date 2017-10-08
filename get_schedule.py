#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime, time
from course_schedule import get_next_course, get_day_courses

username = ''

localxlsx = '/home/dash/Documents/dash_courses.xlsx'
remotexlsx = '/root/qqbot/course_schedule/dash_courses.xlsx'

nowxlsx = remotexlsx

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
			result = get_next_course(nnweek, nntime, nowxlsx)
			bot.SendTo(contact, result)
		elif content == "今天的课":
			nntime = datetime.datetime.now()
			nnweek = nntime.isoweekday()
			results = get_day_courses(nnweek, nowxlsx)
			if isinstance(results, str):
				bot.SendTo(contact, results)
			else:
				for result in results:
					bot.SendTo(contact, result.value)
					time.sleep(0.1)
		elif content == "明天的课":
			nntime = datetime.datetime.now()
			nnweek = nntime.isoweekday()
			results = get_day_courses(nnweek+1, nowxlsx)
			if isinstance(results, str):
				bot.SendTo(contact, results)
			else:
				for result in results:
					bot.SendTo(contact, result.value)
					time.sleep(0.1)
		elif content == "好的":
			username = ''
			bot.SendTo(contact, 'Glad to serve you!   :D')
