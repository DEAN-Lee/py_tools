import os
import configparser
import sys


def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def readDBConf():
    config = configparser.ConfigParser()
    config.read(resource_path(os.path.join('conf.ini')), encoding="utf8")
    return config.get("database", "host"), config.get("database", "user"), config.get(
        "database", "passwd"), config.get("database", "charset"), config.get("database", "db"), config.getint(
        "database", "port")

