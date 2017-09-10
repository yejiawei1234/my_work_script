# -*- coding: utf-8 -*-

import getopt
import os
import datetime
import pandas as pd
import glob
import sys
import time

__version__ = '0.1'
__author__ = 'Yejiawei'


_options = ['dir=', 'timespan=', 'prename=']

_short_options = 'd:t:p:'
try:
    opts, args = getopt.getopt(sys.argv[1:], _short_options, _options)
except getopt.GetoptError as e:
    print('you need help')

if not opts and not args:
    print('you need help')
else:
    conf = {}
    for opt, arg in opts:
        if opt in ('-d', '--dir'):
            conf['dir'] = arg
        elif opt in ('-t', '--timespan'):
            conf['timespan'] = arg
        elif opt in ('-p', '--prename'):
            conf['prename'] = arg





def findworkdir(dir):
    _mypath = os.path.expanduser('~/Desktop/{}'.format(dir))

    if os.path.isabs(dir) and os.path.isdir(dir):
        print('{} is a abs path'.format(dir))
        return dir
    elif os.path.exists(_mypath):

        if os.path.isdir(_mypath):
            print('{} is a path'.format(_mypath))
            return _mypath
        else:
            raise ValueError('{} is not a dir'.format(_mypath))
    else:
        raise ValueError('{} does not exist'.format(_mypath))


def default_time():
    _mydate = datetime.date.today()
    weekday = _mydate.weekday()
    start_delta = datetime.timedelta(days=weekday, weeks=1)
    start_date = _mydate - start_delta
    sevendays = datetime.timedelta(days=7)
    end_date = start_date + sevendays
    f_start_date = '{}{:0>2}{:0>2}'.format(start_date.year, start_date.month, start_date.day)
    f_end_date = '{}{:0>2}{:0>2}'.format(end_date.year, end_date.month, end_date.day)
    timespan = '{}-{}'.format(f_start_date, f_end_date)
    return timespan


def checktime(timespan):
    if timespan:
        print('you input a date {}'.format(timespan))
        return timespan
    else:
        timespan1 = default_time()
        print('{} default date'.format(timespan1))
        return timespan1
def checkprename(prename):
    if prename:
        print('your prename is {}'.format(prename))
        return prename
    else:
        _prename = 'weekly'
        print('your prename is default {}'.format(_prename))
        return _prename



def timeit(func):
    def wapper(*args, **kwargs):
        start_time = time.time()
        p = func(*args, **kwargs)
        end_time = time.time()
        span = end_time - start_time
        print("It takes {:04.2f}s".format(span))
        return p
    return wapper


@timeit
def changefilename(timespan, dirpath, prename):
    filelist = glob.glob('{}/*.xlsx'.format(dirpath))
    newnamelist = []
    for oldname in filelist:
        partname = os.path.dirname(oldname)
        ext = os.path.splitext(oldname)[1]
        basename = pd.read_excel(oldname, sheetname="周报表汇总", header=None).iloc[0, 0]
        name = prename + " " + basename + " " + timespan + ext
        newname = os.path.join(partname, name)
        if newname not in newnamelist:
            newnamelist.append(newname)
            os.rename(oldname, newname)
        else:
            print("{0} is already in file list，old name is {1}".format(os.path.basename(newname), os.path.basename(oldname)))


if __name__ == "__main__":
    dirpath = findworkdir(conf.get('dir'))
    timespan = checktime(conf.get('timespan'))
    prename = checkprename(conf.get('prename'))
    changefilename(timespan=timespan, dirpath=dirpath, prename=prename)

