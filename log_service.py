import builtins
from datetime import datetime


class LogService:
    __log: builtins

    def __init__(self):
        file = "C:\\Logs\\FileListener{0}.log".format(str(datetime.now().date()))
        self.__log = open(file, 'a')

    def line(self, text):
        self.__log.write("{0} -> {1}\n".format(str(datetime.now()), text))
        self.__log.flush()

    def close(self):
        if not self.__log.closed:
            self.__log.close()
