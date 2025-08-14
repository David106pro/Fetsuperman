#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@author: mxh @time:2020/9/27 14:35
"""
import logging
from logging import handlers
import time
import os
import json


def init_log(log_name, _json=False, _console=False, task_id=""):
    """
    初始化日志，按天分割日志
    :return:
    """
    log_name += "%s" % task_id
    logging.debug("01-01 init_process_log method start ...")

    log_content_format = '%(asctime)s %(levelname)s %(module)s.%(funcName)s Line.%(lineno)d %(message)s'
    log_json_format = json.dumps({
        "time": "%(asctime)s",
        "levelname": "%(levelname)s",
        "module_funName_lineno": "%(module)s.%(funcName)s Line.%(lineno)d",
        "message": "%(message)s"
    })
    if _json:
        log_content_format = log_json_format
    log_format = logging.Formatter(log_content_format)

    rotate_handler = SafeRotatingFileHandler(log_name, 'MIDNIGHT', 1, 0)
    # 设置日志后缀
    rotate_handler.suffix = '%Y%m%d'
    rotate_handler.setFormatter(log_format)
    if _console:
        console_handler = logging.StreamHandler()  # 输出到控制台的handler
        console_handler.setFormatter(log_format)
        logging.getLogger().addHandler(console_handler)
    # log = logging.getLogger()
    # 设置日志打印详细信息
    logging.getLogger().addHandler(rotate_handler)

    # 设置日志级别
    logging.getLogger().setLevel(logging.INFO)
    logging.getLogger().debug("01-01 init_process_log method end ...")


class SafeRotatingFileHandler(logging.handlers.TimedRotatingFileHandler):

    def __init__(self, filename, when='h', interval=1, backupCount=0, encoding=None, delay=False, utc=False):
        logging.handlers.TimedRotatingFileHandler.__init__(self, filename, when, interval, backupCount, encoding, delay,
                                                           utc)

    def do_rollover(self):
        try:
            """
            do a rollover; in this case, a date/time stamp is appended to the filename
            when the rollover happens.  However, you want the file to be named for the
            start of the interval, not the current time.  If there is a backup count,
            then we have to get a list of matching filenames, sort them and remove
            the one with the oldest suffix.
            """
            if self.stream:
                self.stream.close()
                self.stream = None
            # get the time that this sequence started at and make it a TimeTuple
            current_time = int(time.time())
            dst_now = time.localtime(current_time)[-1]
            t = self.rolloverAt - self.interval
            if self.utc:
                time_tuple = time.gmtime(t)
            else:
                time_tuple = time.localtime(t)
                dst_then = time_tuple[-1]
                if dst_now != dst_then:
                    if dst_now:
                        addend = 3600
                    else:
                        addend = -3600
                    time_tuple = time.localtime(t + addend)
            dfn = self.baseFilename + "." + time.strftime(self.suffix, time_tuple)
            # if os.path.exists(dfn):
            #    os.remove(dfn)
            # Issue 18940: A file may not have been created if delay is True.
            # if os.path.exists(self.baseFilename):
            if not os.path.exists(dfn) and os.path.exists(self.baseFilename):
                current_file = open(self.baseFilename)
                current_file.close()
                os.rename(self.baseFilename, dfn)
            if self.backupCount > 0:
                for s in self.getFilesToDelete():
                    current_file = open(s)
                    current_file.close()
                    os.remove(s)
            if not self.delay:
                self.stream = self._open()
            new_rollover_at = self.computeRollover(current_time)
            while new_rollover_at <= current_time:
                new_rollover_at = new_rollover_at + self.interval
            # If DST changes and midnight or weekly rollover, adjust for this.
            if (self.when == 'MIDNIGHT' or self.when.startswith('W')) and not self.utc:
                dst_at_rollover = time.localtime(new_rollover_at)[-1]
                if dst_now != dst_at_rollover:
                    if not dst_now:  # DST kicks in before next rollover, so we need to deduct an hour
                        addend = -3600
                    else:  # DST bows out before next rollover, so we need to add an hour
                        addend = 3600
                    new_rollover_at += addend
            self.rolloverAt = new_rollover_at
        except Exception as e:
            logging.error('###############', e)


def init_log_name(log_dir, file_):
    """
    :param log_dir:
    :param file_: 通常调用处传入 __file__ 即可
    :return:
    """
    file_name = os.path.basename(file_)
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    logging_file = os.path.join(log_dir, file_name.replace(".py", ".log"))
    logging.info(logging_file)
    return logging_file
