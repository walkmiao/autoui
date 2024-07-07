import logging
import os
from logging.handlers import TimedRotatingFileHandler

import psutil


def get_pid(name):
    """
    作用：根据进程名获取进程pid
    @param name: 进程名
    @return: pid
    """
    processes = psutil.process_iter()
    for process in processes:
        if process.name() == name:
            return process.pid


def create_logger(name, level='DEBUG'):
    logger = logging.getLogger(__name__)
    new_formatter = '[{}]-[%(levelname)s]%(asctime)s:%(msecs)s.%(process)d,%(thread)d#>[%(funcName)s]:%(lineno)s  %(' \
                    'message)s '.format(name)
    formatter = logging.Formatter(new_formatter)
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        # 如果目录不存在，则创建目录
        os.makedirs(log_dir)
    handler = TimedRotatingFileHandler('{}/{}.log'.format(log_dir,name), when='D', encoding='utf-8', backupCount=10)
    handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.addHandler(stream_handler)
    logger.setLevel(level)
    return logger
