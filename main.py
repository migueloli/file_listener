import json
import smtplib
import ssl
from pathlib import Path

import socket
import sys
import traceback

import servicemanager
import win32event
import win32service
from win32serviceutil import ServiceFramework, HandleCommandLine

from log_service import LogService


class Inovacell(ServiceFramework):
    _svc_name_ = "FileListener"
    _svc_display_name_ = "FileListenerService"
    _svc_description_ = "Serviço para escutar arquivos."

    def __init__(self, args):
        ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

        log = LogService()
        log.line("Stopping service...")
        win32event.SetEvent(self.hWaitStop)
        log.close()

        self.stop_requested = True

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))

        log = LogService()
        log.line("Starting service...")
        log.close()

        self.main()

    def main(self):
        rc = None

        while rc != win32event.WAIT_OBJECT_0:
            log = None
            try:
                log = LogService()

                log.line("Abrindo conexão com o banco...")

                # st_atime e st_mtime alterados ao salvar.
                with open('C:\\file_listener.json') as json_file:
                    data = json.load(json_file)

                for file in data["files"]:
                    file_stats = Path(file['path']).stat()
                    if file['last_update'] >= file_stats.st_mtime:
                        receivers = data["receivers"]
                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL(receivers["server"], receivers["port"], context=context) as server:
                            server.login(receivers["user"], receivers["password"])
                            for email in data["receivers"]:
                                server.sendmail(
                                    receivers["user"],
                                    email,
                                    "Arquivo alterado {}, verifique.".format(file)
                                )

            except Exception as e:
                print(e)
                log.line(str(traceback.format_exc()))
            finally:
                log.line("Fechando conexões...")
                log.line("Concluido...\n")
                if log:
                    log.close()

            rc = win32event.WaitForSingleObject(self.hWaitStop, 300000)
        pass


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(Inovacell)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        HandleCommandLine(Inovacell)

