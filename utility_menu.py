from __future__ import annotations
import sys
import pythoncom
from os import path, unlink, listdir, mkdir, rename, chmod, environ
import os
from stat import S_IWRITE
import pywintypes
import wmi
from wmi import WMI, x_wmi
from winreg import HKEY_USERS, KEY_ALL_ACCESS, DeleteKey, DeleteValue, QueryValueEx, REG_DWORD, KEY_WOW64_64KEY, REG_SZ
from winreg import OpenKey, QueryInfoKey, EnumKey, ConnectRegistry, HKEY_LOCAL_MACHINE, KEY_SET_VALUE, SetValueEx
from shutil import rmtree, copy
from subprocess import run, CREATE_NO_WINDOW
from time import sleep, time
from getpass import getuser
from win32net import NetShareEnum
from random import random
from datetime import datetime, timedelta
from json import load, loads, dump
from logging import getLogger, basicConfig, exception
from threading import Thread
from concurrent.futures import TimeoutError, ThreadPoolExecutor
from functools import wraps
from psutil import disk_usage, process_iter, ZombieProcess, AccessDenied, NoSuchProcess
from pyad import adquery, aduser, adgroup, pyadutils, pyadexceptions
from datetime import timedelta
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import QObject, pyqtSignal, QRunnable, pyqtSlot, QThreadPool, Qt, QCoreApplication
from PyQt6.QtWidgets import QMessageBox, QMainWindow, QColorDialog
import traceback
import pyodbc
from typing import Callable


class Objects:
    """The objects which will be passed via signals"""
    first = True
    objects = dict()
    texts = dict()
    buttons = dict()


class PassSignals:
    """The signals for the QT threads directly linked to emit method to shorten code"""
    def __init__(
            self, print_: pyqtSignal, print_success_: pyqtSignal, print_error_: pyqtSignal, ask_yes_no_: pyqtSignal,
            show_info_: pyqtSignal, call_pb: pyqtSignal, progress_callback: pyqtSignal, clean_pb: pyqtSignal,
            update_, update_error_, update_success_, copy_: pyqtSignal, enable_1_: pyqtSignal, disable_1_: pyqtSignal,
            clear_all_: pyqtSignal, task_pb: pyqtSignal, task_: pyqtSignal, del_users_: pyqtSignal,
            run_without_waiting: pyqtSignal, zoom: pyqtSignal
    ):
        self.print = print_.emit
        self.print_success = print_success_.emit
        self.print_error = print_error_.emit
        self.ask_yes_no = ask_yes_no_.emit
        self.show_info = show_info_.emit
        self.call_pb = call_pb.emit
        self.progress = progress_callback.emit
        self.clean_pb = clean_pb.emit
        self.update = update_.emit
        self.update_error = update_error_.emit
        self.update_success = update_success_.emit
        self.copy = copy_.emit
        self.enable_1 = enable_1_.emit
        self.disable_1 = disable_1_.emit
        self.clear_all = clear_all_.emit
        self.task_pb = task_pb.emit
        self.task = task_.emit
        self.del_users = del_users_.emit
        self.run_without_waiting = run_without_waiting.emit
        self.zoom = zoom.emit


class WorkerSignals(QObject):
    """The signals for each worker"""
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal(int)
    print_success_ = pyqtSignal(str, str)
    print_ = pyqtSignal(str)
    print_error_ = pyqtSignal(str, str)
    ask_yes_no = pyqtSignal(str, str)
    yes_no = (False, False)
    show_info_ = pyqtSignal(str, str)
    call_pb = pyqtSignal(str)
    clean_pb = pyqtSignal(str)
    update = pyqtSignal(str, str)
    update_error = pyqtSignal(str, str)
    update_success = pyqtSignal(str, str)
    copy = pyqtSignal(str)
    enable_1 = pyqtSignal(str)
    disable_1 = pyqtSignal(str)
    clear_all = pyqtSignal()
    task = pyqtSignal(object, str)
    task_pb = pyqtSignal(object, str)
    del_users = pyqtSignal(list)
    run_without_waiting = pyqtSignal(object, list)
    zoom = pyqtSignal()

    @staticmethod
    def yes_or_no():
        return WorkerSignals.yes_no == (True, True)


class Worker(QRunnable):
    """The worker it self which will be called from the main thread, to avoid any freezes in the UI"""
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self.to_pass = PassSignals(
            self.signals.print_, self.signals.print_success_, self.signals.print_error_,
            self.signals.ask_yes_no, self.signals.show_info_, self.signals.call_pb, self.signals.progress,
            self.signals.clean_pb, self.signals.update, self.signals.update_error, self.signals.update_success,
            self.signals.copy, self.signals.enable_1, self.signals.disable_1, self.signals.clear_all,
            self.signals.task_pb, self.signals.task, self.signals.del_users, self.signals.run_without_waiting,
            self.signals.zoom
        )
        self.kwargs["signals"] = self.to_pass

    @pyqtSlot()
    def run(self) -> None:
        try:
            self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            log()
        finally:
            self.signals.finished.emit()


class ProgressBar:
    """Progress bar for long tasks"""
    def __init__(self, total_items: int, title: str, end: str, signals: PassSignals):
        """"initial configuration of the progressbar"""
        self.total_items = total_items + 1
        self.current_item = 0
        self.signals = signals
        self.end = end
        signals.call_pb(title)

    def __call__(self):
        """updates the progressbar when its being called"""
        self.current_item += 1
        if self.current_item <= self.total_items:
            progress = (self.current_item / self.total_items) * 100
            self.signals.progress(int(progress))

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.signals.clean_pb(self.end)


class AskYesNo:
    """A simple yes or no dialog"""
    def __init__(self, title: str, messege: str, signals: PassSignals):
        """"initial configuration of the progressbar"""
        WorkerSignals.yes_no = (False, False)
        signals.ask_yes_no(title, messege)
        while WorkerSignals.yes_no == (False, False):
            sleep(0.3)

    @staticmethod
    def yes_no():
        return WorkerSignals.yes_no == (True, True)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass


class FatalError(QMainWindow):
    """Fatal Error dialog"""
    def __init__(self, title: str = "Error!", prompt_: str = "Fatal error occured"):
        app_ = QtWidgets.QApplication(sys.argv)
        super().__init__()
        QMessageBox.critical(None, title, prompt_, buttons=QMessageBox.StandardButton.Ok)
        sys.exit(1)


class Pointers:
    """Cursor pointers"""
    @staticmethod
    def hand():
        return QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor)

    @staticmethod
    def normal():
        return QtGui.QCursor(QtCore.Qt.CursorShape.ArrowCursor)


class Formats:
    """Colors for the 'console'"""
    error = '<span style="color:rgb(255, 0, 4);">{}</span>'
    warning = '<span style="color:orange;">{}</span>'
    success = '<span style="color:rgb(0, 255, 0);">{}</span>'
    normal = '<span style="color:black;">{}</span>'


class Fonts:
    """Fonts which will be used for the UI"""
    ariel_12 = QtGui.QFont()
    ariel_12.setFamily("Arial")
    ariel_12.setPointSize(12)
    ariel_12_bold = QtGui.QFont()
    ariel_12_bold.setFamily("Arial")
    ariel_12_bold.setPointSize(12)
    ariel_12_bold.setBold(True)
    ariel_11 = QtGui.QFont()
    ariel_11.setFamily("Arial")
    ariel_11.setPointSize(11)
    ariel_11_bold = QtGui.QFont()
    ariel_11_bold.setFamily("Arial")
    ariel_11_bold.setPointSize(11)
    ariel_11_bold.setBold(True)


def redirect(*output: str) -> None:
    """redirects all output to the console Text object"""
    if not str(*output) == "\n":
        ui.console.append(Formats.normal.format(str(*output)))
        refresh()


def my_exception_hook(exctype, value, traceback_):
    """An exception hook which will display uncaught exceptions to the 'console'"""
    print(traceback.format_exception(exctype, value, traceback_))
    log()


def ask_yes_no(title: str, text: str) -> None:
    """A function which initialize and manage the output of the YesNo instance"""
    WorkerSignals.yes_no = (False, False)
    yes_no = QMessageBox.question(main_window_, title, text,
                                  QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    if yes_no == QMessageBox.StandardButton.Yes:
        WorkerSignals.yes_no = (True, True)
    else:
        WorkerSignals.yes_no = (False, True)
    sleep(1)
    WorkerSignals.yes_no = (False, False)


def clear_all(first: bool = False) -> None:
    """Clears all the UI display elements"""
    """"clears all objects in the gui window, and sets their default text"""
    for obj in ((ui.pc_display, "Computer: "), (ui.pc_status, "Computer status: "),
                (ui.user_display, "Current user: "), (ui.uptime_display, "Uptime: "),
                (ui.c_space_display, "Space in C disk: "), (ui.d_space_display, "Space in D disk: "),
                (ui.ram_display, "Total RAM: "), (ui.ie_display, "Internet explorer: "),
                (ui.cpt_status, "Cockpit printer: "), (ui.user_status, "User status: ")):
        if first:
            obj[0].setFont(Fonts.ariel_11_bold)
            obj[0].setLineWrapMode(QtWidgets.QTextBrowser.LineWrapMode.NoWrap)
        obj[0].clear()
        obj[0].append(obj[1])
        ui.console.clear()
    refresh()
    if Objects.texts:
        return
    Objects.texts = {ui.pc_display: "Computer: ", ui.pc_status: "Computer status: ", ui.user_display: "Current user: ",
                     ui.uptime_display: "Uptime: ", ui.c_space_display: "Space in C disk: ",
                     ui.d_space_display: "Space in D disk: ", ui.ram_display: "Total RAM: ",
                     ui.ie_display: "Internet explorer: ", ui.cpt_status: "Cockpit printer: ",
                     ui.user_status: "User status: "}


def print_error(obj: QtWidgets | str, output: str) -> None:
    """prints an error to an object, will be red colored"""
    if isinstance(obj, str):
        obj = Objects.objects[obj]
    obj.append(Formats.error.format(output))
    refresh()


def print_success(obj: QtWidgets | str, output: str) -> None:
    """print green text to an object"""
    if isinstance(obj, str):
        obj = Objects.objects[obj]
    obj.append(Formats.success.format(output))
    refresh()


def update(obj: QtWidgets | str, statement: str) -> None:
    """updates an object with new text"""
    if isinstance(obj, str):
        obj = Objects.objects[obj]
    obj.clear()
    obj.append(Objects.texts[obj] + statement)
    refresh()


def update_error(obj: QtWidgets | str, statement: str | timedelta) -> None:
    """Updates a display element with text-colored red"""
    if isinstance(obj, str):
        obj = Objects.objects[obj]
    obj.clear()
    obj.append(Objects.texts[obj] + Formats.error.format(statement))
    refresh()


def update_success(obj: QtWidgets | str, statement: str) -> None:
    """Updates a display element with text-colored green"""
    if isinstance(obj, str):
        obj = Objects.objects[obj]
    obj.clear()
    obj.append(Objects.texts[obj] + Formats.success.format(statement))
    refresh()


def disable(disable_submit: bool = False) -> None:
    """"disables all the buttons, so they aren't clickable while a function is still executing, also disables submitting
    by pressing the enter key"""
    for obj in (ui.export_btn, ui.fix_ie_btn, ui.fix_cpt_btn, ui.close_outlook_btn, ui.del_zoom_btn,
                ui.printers_btn, ui.clear_space_btn, ui.del_users_btn, ui.del_teams_btn, ui.sample_btn, ui.del_ost_btn,
                ui.reset_spool_btn, ui.restart_pc_btn, ui.fix_3_lang_btn):
        obj.setEnabled(False)

    refresh()
    if disable_submit:
        ui.submit_btn.setEnabled(False)
    else:
        ui.submit_btn.setEnabled(True)
    if not config.current_computer:
        ui.copy_btn.setEnabled(False)
    refresh()


def enable() -> None:
    """"enables the buttons back, also enables submitting via pressing enter"""
    if Objects.first:
        Objects.first = False
        for obj in (ui.export_btn, ui.copy_btn, ui.fix_ie_btn, ui.close_outlook_btn, ui.del_zoom_btn, ui.printers_btn,
                    ui.clear_space_btn, ui.del_users_btn, ui.reset_spool_btn, ui.submit_btn, ui.restart_pc_btn,
                    ui.fix_3_lang_btn, ui.del_ost_btn, ui.fix_cpt_btn, ui.del_teams_btn, ui.sample_btn):
            obj.setCursor(Pointers.hand())
        refresh()
    for obj in (ui.export_btn, ui.copy_btn, ui.fix_ie_btn, ui.close_outlook_btn, ui.del_zoom_btn, ui.printers_btn,
                ui.clear_space_btn, ui.del_users_btn, ui.reset_spool_btn, ui.restart_pc_btn,
                ui.fix_3_lang_btn, sample_btn):
        obj.setEnabled(True)
    refresh()
    if not config.current_user or config.disable_user_depends:
        ui.del_ost_btn.setEnabled(False)
        ui.fix_cpt_btn.setEnabled(False)
        ui.del_teams_btn.setEnabled(False)
        refresh()
        return
    ui.del_ost_btn.setEnabled(True)
    ui.fix_cpt_btn.setEnabled(True)
    ui.del_teams_btn.setEnabled(True)
    refresh()


def enable_1(obj: str | QtWidgets) -> None:
    """Enables one button"""
    if isinstance(obj, str):
        obj = Objects.buttons[obj]
    obj.setEnabled(True)


def disable_1(obj: str | QtWidgets) -> None:
    """Disables one button"""
    if isinstance(obj, str):
        obj = Objects.buttons[obj]
    obj.setEnabled(False)


def copy_clip(to_copy: str) -> None:
    """"copies the computer name to the user's clipboard"""
    ui.clipboard.setText(to_copy)


def asset(filename: str) -> str:
    """return the full path to images - assets that the script uses"""
    return fr"{config.assets}\{filename}"


def fix_ie_func(signals: PassSignals) -> None:
    """fixes the Internet Explorer application via deleting appropriate registry keys, as well as disabling
    compatibility mode"""
    pc = config.current_computer
    with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg:
        for key_name in (
                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{"
                r"1FD49718-1D00-4B19-AF5F-070AF6D5D54C}",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{"
                r"1FD49718-1D00-4B19-AF5F-070AF6D5D54C"):
            try:
                with OpenKey(reg, key_name, 0, KEY_ALL_ACCESS) as key:
                    DeleteKey(key, "")
            except FileNotFoundError:
                pass
            except:
                signals.print_error(Objects.console, "Unable to fix internet explorer")
                log()
                return

    key = fr"{get_sid()}\Software\Microsoft\Internet Explorer\BrowserEmulation"
    try:
        with ConnectRegistry(pc, HKEY_USERS) as reg:
            with OpenKey(reg, key, 0, KEY_SET_VALUE) as key:
                SetValueEx(key, "IntranetCompatibilityMode", 0, REG_DWORD, 0)
                SetValueEx(key, "MSCompatibilityMode", 0, REG_DWORD, 1)
    except FileNotFoundError:
        pass
    except:
        log()

    signals.update_success(Objects.ie_display, "Fixed")
    signals.print_success(Objects.console, "Fixed internet explorer")


def fix_cpt_func(signals: PassSignals) -> None:
    """"fixes cockpit printer via deleting the appropriate registry keys"""
    sid_ = config.current_sid
    if not sid_:
        return
    with ConnectRegistry(config.current_computer, HKEY_USERS) as reg:
        try:
            with OpenKey(reg, fr"{sid_}\SOFTWARE\Jetro Platforms\JDsClient\PrintPlugIn", 0, KEY_ALL_ACCESS) as key:
                DeleteValue(key, "PrintClientPath")
        except FileNotFoundError:
            pass
        except:
            log()
            signals.print_error(Objects.console, "Failed to fix cpt printer")
            return
    signals.print_success(Objects.console, "Fixed cpt printer")
    signals.update_success(Objects.cpt_status, "Fixed")


def fix_3_languages(signals: PassSignals) -> None:
    """fixes 3 languages bug via deleting the appropriate registry keys, need to restart PC after"""
    with ConnectRegistry(config.current_computer, HKEY_USERS) as reg:
        try:
            with OpenKey(reg, r".DEFAULT\Keyboard Layout\Preload", 0, KEY_ALL_ACCESS) as key:
                DeleteKey(key, "")
        except FileNotFoundError:
            pass
        except:
            signals.print_error(Objects.console, "Failed to fix 3 languages bug")
            log()
            return
    signals.print_success(Objects.console, "Fixed 3 languages bug")


def reset_spooler(signals: PassSignals) -> None:
    """resets the print spooler via WMI"""
    try:
        pythoncom.CoInitialize()
        connection = WMI(computer=config.current_computer)
        service = connection.Win32_Service(name="Spooler")
        service[0].StopService()
        sleep(1)
        service[0].StartService()
        signals.print_success(Objects.console, "Successfully restarted the spooler")
    except:
        signals.print_error(Objects.console, "Failed to restart the spooler")
        log()
    try:
        with ConnectRegistry(config.current_computer, HKEY_LOCAL_MACHINE) as reg:
            with OpenKey(reg, r"SYSTEM\CurrentControlSet\Services\Spooler", 0, KEY_ALL_ACCESS | KEY_WOW64_64KEY) as key:
                SetValueEx(key, 'DependOnService', 0, REG_SZ, "RPCSS")
    except (PermissionError, FileNotFoundError):
        pass
    except:
        log()


def delete_the_ost(signals: PassSignals) -> None:
    """renames the ost file to .old with random digits to avoid conflict with duplicate ost filenames
    handles shutting down outlook and skype on the remote computer so the OST could be renamed"""
    user_ = config.current_user
    pc = config.current_computer
    pythoncom.CoInitialize()
    with AskYesNo("OST deletion", f"Are you sure you want to delete the ost of {user_name_translation(user_)}",
                  signals):
        pass
    if WorkerSignals.yes_no == (False, True):
        signals.print("Canceled OST deletion")
        return
    host = WMI(computer=pc)
    for procs in ("lync.exe", "outlook.exe", "UcMapi.exe"):
        for proc in host.Win32_Process(name=procs):
            if proc:
                try:
                    proc.Terminate()
                except:
                    log()
    if not path.exists(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook"):
        signals.print_error(Objects.console, "Could not find an OST file")
        return

    ost = listdir(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook")
    for file___ in ost:
        if file___.endswith("ost"):
            ost = fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook\{file___}"
            try:
                sleep(1)
                rename(ost, f"{ost}{random():.3f}.old")
                signals.print_success(Objects.console, "Successfully removed the ost file")
                return
            except FileExistsError:
                try:
                    rename(ost, f"{ost}{random():.3f}.old")
                    signals.print_success(Objects.console, "Successfully removed the ost file")
                    return
                except:
                    log()
                    signals.print_error(Objects.console, f"Could not Delete the OST file")
                    return
            except:
                signals.print_error(Objects.console, f"Could not Delete the OST file")
                log()
                return
    else:
        signals.print_error(Objects.console, f"Could not find an OST file")


def my_rm(file_: str, bar: ProgressBar) -> None:
    """removes readonly files via changing the file permissions"""
    try:
        if path.isfile(file_) or path.islink(file_):
            unlink(file_)
        elif path.isdir(file_):
            rmtree(file_, ignore_errors=True)
    except (PermissionError, FileNotFoundError):
        pass
    except:
        log()
    bar()


def rmtree_recreate(dir_: str) -> None:
    """removes the entire dir with its contents and then recreate it"""
    try:
        rmtree(dir_, ignore_errors=True)
        mkdir(dir_)
    except (FileExistsError, PermissionError):
        pass
    except:
        log()


def my_rmtree(dir_: str, bar_: ProgressBar | None = None) -> None:
    """delete folders and their contents, then calls the bar object"""
    if path.isdir(dir_):
        rmtree(dir_, onerror=on_rm_error)
    if bar_ is not None:
        bar_()


def clear_space_func(signals: PassSignals) -> None:
    """clears spaces from the remote computer, paths, and other configurations as for which files to delete
    can be configured via the config file. using multithreading to delete the files faster"""
    pc = config.current_computer
    users_dirs = listdir(fr"\\{pc}\c$\users")
    pythoncom.CoInitialize()

    space_init = get_space(pc)
    flag = False

    edb_file = fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
    if path.exists(edb_file) and config.delete_edb:
        try:
            connection = WMI(computer=pc)
            connection.Win32_Process.Create(
                CommandLine='cmd.exe /c powercfg.exe /hibernate off"'
            )
            service = connection.Win32_Service(name="WSearch")
            service[0].StopService()
            sleep(0.6)
            unlink(fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb")
            service[0].StartService()
            flag = True
        except (PermissionError, pywintypes.com_error, FileNotFoundError):
            pass
        except:
            log()

    if config.c_paths_with_msg:
        for path_msg in config.c_paths_with_msg:
            if len(path_msg[0]) < 3:
                continue
            if path.exists(fr"\\{pc}\c$\{path_msg[0]}"):
                files = [fr"\\{pc}\c$\{path_msg[0]}\{file___}" for file___ in listdir(fr"\\{pc}\c$\{path_msg[0]}")]
                with ProgressBar(len(files), path_msg[1], path_msg[-1], signals) as bar:
                    with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                        jobs = [executor.submit(my_rm, file___, bar) for file___ in files]
                        while not all([result.done() for result in jobs]):
                            sleep(0.1)

    if config.delete_user_temp:
        with ProgressBar(len(users_dirs), f"Deleting temps of {len(users_dirs)} users",
                         f"Deleted temps of {len(users_dirs)} users", signals) as bar:
            dirs = [fr"\\{pc}\c$\users\{dir_}\AppData\Local\Temp" for dir_ in users_dirs if
                    (dir_.lower().strip() != config.user.lower().strip() and config.current_computer.lower()
                     != config.host.lower().strip())]
            with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                jobs = [executor.submit(my_rm, dir_, bar) for dir_ in dirs]
                while not all([result.done() for result in jobs]):
                    sleep(0.1)

    if config.u_paths_with_msg:
        for path_msg in config.u_paths_with_msg:
            if len(path_msg[0]) < 3:
                continue
            msg_ = path_msg[1].replace("users_amount", str(len(users_dirs)))
            with ProgressBar(len(users_dirs), msg_, path_msg[-1].replace(str(len(users_dirs))), signals) as bar:
                for user in users_dirs:
                    if path.exists(fr"\\{pc}\c$\users\{user}\{path_msg[0]}"):
                        files = listdir(fr"\\{pc}\c$\users\{user}\{path_msg[0]}")
                        with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                            jobs = [executor.submit(my_rm, file___, bar) for file___ in files]
                            while not all([result.done() for result in jobs]):
                                sleep(0.1)
                    bar()

    if config.u_paths_without_msg or config.c_paths_without_msg:
        with ProgressBar(len(config.u_paths_without_msg) + len(config.c_paths_without_msg), "Deleting additional files",
                         "Deleted additional files", signals) as bar:
            for path_msg in config.c_paths_without_msg:
                if len(path_msg[0]) < 3:
                    continue
                if path.exists(fr"\\{pc}\c$\{path_msg[0]}"):
                    files = listdir(fr"\\{pc}\c$\{path_msg[0]}")
                    with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                        jobs = [executor.submit(my_rm, file___, bar) for file___ in files]
                        while not all([result.done() for result in jobs]):
                            sleep(0.1)
                bar()

            for path_msg in config.u_paths_without_msg:
                if len(path_msg[0]) < 3:
                    continue
                for user in users_dirs:
                    if path.exists(fr"\\{pc}\c$\users\{user}\{path_msg[0]}"):
                        files = listdir(fr"\\{pc}\c$\users\{user}\{path_msg[0]}")
                        with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                            jobs = [executor.submit(my_rm, file___, bar) for file___ in files]
                            while not all([result.done() for result in jobs]):
                                sleep(0.1)
                bar()

    if not flag and config.delete_edb and path.exists(edb_file):
        try:
            connection = WMI(computer=pc)
            service = connection.Win32_Service(name="WSearch")
            service[0].StopService()
            sleep(0.8)
            unlink(fr"\\{pc}\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb")
            service[0].StartService()
            flag = True
        except (PermissionError, FileNotFoundError):
            pass
        except:
            log()
    if flag and config.delete_edb:
        signals.print(f"Deleted the search.edb file")
    else:
        if config.delete_edb:
            signals.print_error(Objects.console, "Failed to remove search.edb file")
    space_final = get_space(pc)
    signals.print_success(Objects.console, f"Cleared {abs((space_final - space_init)):.1f} GB from the disk")

    try:
        space = get_space(pc)
        if space <= 5:
            signals.update_error(Objects.c_space_display, f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
        else:
            signals.update(Objects.c_space_display,
                           f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
    except:
        log()
        signals.update_error(Objects.c_space_display, "ERROR")


def sample_function(signals: PassSignals):
    """I had my own function here for my organization, cant share due to privacy"""
    pass


def del_ost(signals: PassSignals):
    """renames the ost file to .old with random digits to avoid conflict with duplicate ost filenames
        handles shutting down outlook and skype on the remote computer so the OST could be renamed"""
    user_ = config.current_user
    pc = config.current_computer
    pythoncom.CoInitialize()
    with AskYesNo("OST deletion",
                  f"Are you sure you want to delete the ost of {user_name_translation(user_)}?", signals):
        pass
    if WorkerSignals.yes_no == (False, True):
        signals.print("Canceled OST deletion")
        return
    host = WMI(computer=pc)
    for procs in ("lync.exe", "outlook.exe", "UcMapi.exe"):
        for proc in host.Win32_Process(name=procs):
            if proc:
                try:
                    proc.Terminate()
                except:
                    log()
    if not path.exists(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook"):
        signals.print_error(Objects.console, f"Could not find an OST file")
        return

    ost = listdir(fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook")
    for file___ in ost:
        if file___.endswith("ost"):
            ost = fr"\\{pc}\c$\Users\{user_}\AppData\Local\Microsoft\Outlook\{file___}"
            try:
                sleep(1)
                rename(ost, f"{ost}{random():.3f}.old")
                signals.print_success(Objects.console, "Successfully removed the ost file")
                return
            except FileExistsError:
                try:
                    rename(ost, f"{ost}{random():.3f}.old")
                    signals.print_success(Objects.console, "Successfully removed the ost file")
                    return
                except:
                    log()
                    signals.print_error(Objects.console, "Could not Delete the OST file")
                    return
            except:
                signals.print_error(Objects.console, "Could not Delete the OST file")
                log()
                return
    else:
        signals.print_error(Objects.console, "Could not Delete the OST file")


def del_users(signals: PassSignals) -> None:
    """gives you the option to choose which users folders to delete as well as multithreading or WMI deletion of folders
    will exclude the current user of the remote PC if found one. 
    users to exclude could be configured in the config file"""
    pythoncom.CoInitialize()
    config.yes_no = False
    pc = config.current_computer
    users_to_choose_delete = {}
    to_pass = []
    for dir_ in listdir(fr"\\{pc}\c$\Users"):
        if str(config.current_user).lower().strip() in dir_.lower() or dir_.lower() in config.exclude or \
                any([dir_.lower().startswith(exc_lude) for exc_lude in config.startwith_exclude]) \
                or not path.isdir(fr"\\{pc}\c$\users\{dir_}"):
            continue
        translated = user_name_translation(dir_)
        users_to_choose_delete[translated] = dir_
        to_pass.append(translated)
    if not users_to_choose_delete:
        signals.print("No users were found to delete")
        return
    will_delete = [users_to_choose_delete[user_] for user_ in deletion_selection(to_pass, signals)]
    if not will_delete:
        signals.print("No users were found to delete")
        return
    with AskYesNo("Users deletion", f"Are you sure you want to delete the following users?\n\n"
                                    + '.\n'.join([user_name_translation(user_) for user_ in will_delete]),
                  signals):
        pass
    if WorkerSignals.yes_no == (False, True):
        signals.print("Canceled users deletion")
        return
    space_init = get_space(pc)
    if config.lr == 1:
        with ProgressBar(len(will_delete), f"Deleting {len(will_delete)} folders",
                         f"Deleted {len(will_delete)} users", signals) as bar:
            will_delete = [fr"\\{pc}\c$\users\{dir_}" for dir_ in will_delete]
            with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
                jobs = [executor.submit(my_rmtree, dir_, bar) for dir_ in will_delete]
                while not all([result.done() for result in jobs]):
                    sleep(0.3)
        space_final = get_space(pc)
        signals.print(f"Cleared {abs((space_final - space_init)):.1f} GB from the disk")
        try:
            space = get_space(pc)
            if space <= 5:
                signals.update_error(Objects.c_space_display, f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
            else:
                signals.update(Objects.c_space_display, f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
        except:
            log()
            signals.update_error(Objects.c_space_display, "ERROR")

    else:
        try:
            pc_con = WMI(computer=pc)
        except wmi.x_wmi:
            signals.print_error(Objects.console, "Couldn't connect to WMI")
            return
        for dir_ in will_delete:
            pc_con.Win32_Process.Create(
                CommandLine=fr'cmd.exe /c rd /s /q "C:/Users/{dir_}" & echo done > "c:/{dir_}_done.txt"'
            )
        signals.print_success(Objects.console, f"started deletion of {len(will_delete)} users using WMI")
        signals.run_without_waiting(wait_del, [pc, will_delete, space_init])


def wait_del(pc: str, users_l: list, space_: float, signals: PassSignals) -> None:
    """Awaits for the WMI deletion"""
    pythoncom.CoInitialize()
    to_check = users_l.copy()
    fails = 0
    sleep(1)
    while any([path.exists(fr"\\{pc}\c$\users\{dir_}") for dir_ in users_l]) and to_check:
        for user_ok in to_check.copy():
            if not path.exists(fr"\\{pc}\c$"):
                fails += 1
                sleep(2)
                if fails >= 3:
                    signals.show_info("Failure to delete users", f"The computer {pc} is not available anymore"
                                                                 f"the status of the deletion of "
                                                                 f"{len(users_l)} users is unknown")
                    return
            if path.exists(fr"\\{pc}\c$\{user_ok}_done.txt"):
                if path.exists(fr"\\{pc}\c$\{user_ok}"):
                    my_rmtree(fr"\\{pc}\c$\{user_ok}")
                    try:
                        unlink(fr"\\{pc}\c$\{user_ok}_done.txt")
                    except (FileNotFoundError, PermissionError):
                        pass
                else:
                    try:
                        unlink(fr"\\{pc}\c$\{user_ok}_done.txt")
                    except (FileNotFoundError, PermissionError):
                        pass
                to_check.remove(user_ok)
        if not any([path.exists(fr"\\{pc}\c$\users\{dir_}") for dir_ in to_check]) or not to_check:
            break
        sleep(0.75)
    sleep(1)
    try:
        final = get_space(pc)
        signals.show_info(
            f"Done deleting users on {pc}", f"Done deleting {len(users_l)} users on {pc}"
                                            f"Cleared {abs((final - space_)):.1f} GB from the disk there's now "
                                            f"{final:.1f}GB free out of {get_total_space(pc):.1f}GB"
        )
        if config.current_computer == pc:
            space = get_space(pc)
            if space <= 5:
                signals.update_error(Objects.c_space_display, f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
            else:
                signals.update(Objects.c_space_display, f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
    except:
        log()
        signals.show_info(
            "User deletion has ended",
            f"user deletion on {pc} has ended, could not retrieve"f"disk space"
        )


def get_printers_func(signals: PassSignals) -> None:
    """retrieves all network printers installed on the remote computer
     achieves that via querying the appropriate registry keys"""
    found_any = False
    pc = config.current_computer
    with ConnectRegistry(pc, HKEY_USERS) as reg:
        users_dict = {}
        sid_list = []
        with OpenKey(reg, "") as users:
            users_len = QueryInfoKey(users)[0]
            for i in range(users_len):
                try:
                    sid_list.append(EnumKey(users, i))
                except FileNotFoundError:
                    pass
                except:
                    log()

        with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as users_path:
            pythoncom.CoInitialize()
            for sid in set(sid_list):
                try:
                    with OpenKey(users_path,
                                 fr"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{sid}") as profiles:
                        username = QueryValueEx(profiles, "ProfileImagePath")
                        if username[0].startswith("C:\\"):
                            username = username[0].split("\\")[-1]
                            users_dict[sid] = user_name_translation(username)
                except FileNotFoundError:
                    pass
                except:
                    log()

        flag = False
        for sid in sid_list:
            try:
                with OpenKey(reg, fr"{sid}\Printers\Connections") as printer_path:
                    printers_len = QueryInfoKey(printer_path)[0]
                    for i in range(printers_len):
                        try:
                            printer = EnumKey(printer_path, i).replace(",", "\\").strip()
                            p = f"{printer} was found on user {users_dict[sid]}"
                            if not flag:
                                signals.print(f"\n\n{' Network printers ':-^118}")
                                flag = True
                            signals.print(p)
                            found_any = True
                        except:
                            log()
            except FileNotFoundError:
                pass
            except:
                log()
    flag = False
    with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg:
        with OpenKey(reg, r'SYSTEM\CurrentControlSet\Control\Print\Printers') as printers:
            found = []
            printers_len = QueryInfoKey(printers)[0]
            for i in range(printers_len):
                with OpenKey(printers, EnumKey(printers, i)) as printer:
                    try:
                        prnt = QueryValueEx(printer, "Port")[0]
                        if "_" in prnt:
                            prnt = prnt.split("_")[0].strip()
                        if prnt in found:
                            continue
                        found.append(prnt)
                        int(prnt.split(".")[0])
                        if not flag:
                            signals.print(f"\n\n{' TCP/IP printers ':-^119}")
                            flag = True
                        signals.print(
                            f"TCP/IP Printer with an IP of {prnt} is located at {config.ip_printers[prnt.strip()]}" if
                            prnt in config.ip_printers else
                            f"Printer with an IP of {prnt} is not on any of the servers")
                        found_any = True
                    except (FileNotFoundError, ValueError):
                        pass
                    except:
                        log()
        flag = False
        with OpenKey(reg, r"SYSTEM\CurrentControlSet\Enum\SWD\PRINTENUM") as printers:
            found = []
            printers_len = QueryInfoKey(printers)[0]
            for i in range(printers_len):
                with OpenKey(printers, EnumKey(printers, i)) as printer:
                    try:
                        prnt = QueryValueEx(printer, "LocationInformation")[0].split("/")[2].split(":")[0]
                        if prnt in found:
                            continue
                        if "_" in prnt:
                            prnt = prnt.split("_")[0]
                        found.append(prnt)
                        int(prnt.split(".")[0])
                        if not flag:
                            signals.print(f"\n\n{' WSD printers ':-^119}")
                            flag = True
                        signals.print(
                            f"WSD printer with an IP of {prnt.strip()} is located at "
                            f"{config.ip_printers[prnt.strip()]}" if prnt.strip() in config.ip_printers else
                            f"WSD printer with an IP of {prnt} is not on any of the servers")
                        found_any = True
                    except (FileNotFoundError, ValueError, IndexError):
                        pass
                    except:
                        log()
    if not found_any:
        signals.print_error(Objects.console, "No printers were found")


def rem_teams(signals: PassSignals) -> None:
    """removes the teams related files for fresh install"""
    with AskYesNo("Delete teams files?", "are you sure you want to delete teams?", signals):
        pass
    if WorkerSignals.yes_no == (False, True):
        signals.print("Canceled teams deletion")
        return
    pc = config.current_computer
    pythoncom.CoInitialize()
    host = WMI(computer=pc)
    for procs in ("lync.exe", "outlook.exe", "UcMapi.exe", "msedgewebview2.exe", "teams.exe", "ms-teams.exe"):
        for proc in host.Win32_Process(name=procs):
            if proc:
                try:
                    proc.Terminate()
                except x_wmi:
                    pass
                except:
                    log()
    user_ = config.current_user
    pc_ = config.current_computer
    new_teams = fr"\\{pc_}\C$\Program Files\WindowsApps\MSTeams_23285.3604.2469.4152_x64__8wekyb3d8bbwe"
    for dirname in listdir(fr"\\{pc_}\c$\Program Files\WindowsApps"):
        if dirname.lower().startswith("msteams"):
            new_teams = fr"\\{pc_}\C$\Program Files\WindowsApps\{dirname}"
            break
    base = fr"\\{pc_}\c$\users\{user_}\appdata"
    dirs = (fr"\\{pc_}\C$\Program Files (x86)\Teams Installer", path.join(base, r"Roaming\Teams"),
            path.join(base, r"Roaming\Microsoft\Teams"), path.join(base, r"Local\Microsoft\TeamsMeetingAddin"),
            path.join(base, r"Local\Microsoft\Teams"), path.join(base, r"Local\Microsoft\TeamsPresenceAddin"),
            new_teams)
    with ProgressBar(len(dirs), "Deleting Teams files", "Deleted teams files", signals) as bar:
        with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
            jobs = [executor.submit(my_rmtree, dir_, bar) for dir_ in dirs]
            while not all([result.done() for result in jobs]):
                sleep(0.3)
        sleep(10)
    sleep(0.3)


def close_outlook(signals: PassSignals):
    """closes the outlook and other skype related processes"""
    pc = config.current_computer
    pythoncom.CoInitialize()
    host = WMI(computer=pc)
    for procs in ("lync.exe", "outlook.exe", "UcMapi.exe"):
        for proc in host.Win32_Process(name=procs):
            if proc:
                try:
                    proc.Terminate()
                except:
                    log()
    signals.print_success(Objects.console, "Shut down outlook Successfully")


def export(signals: PassSignals):
    """exports the mapped network drives and network printers to a txt file and .bat file to auto install"""
    to_export = config.current_computer
    with ConnectRegistry(to_export, HKEY_USERS) as reg:
        svr = []
        tcp = []
        wsd = []
        sid_list = []
        with OpenKey(reg, "") as users:
            users_len = QueryInfoKey(users)[0]
            for i in range(users_len):
                try:
                    sid_list.append(EnumKey(users, i))
                except FileNotFoundError:
                    pass
                except:
                    log()

        for sid in sid_list:
            try:
                with OpenKey(reg, fr"{sid}\Printers\Connections") as printer_path:
                    printers_len = QueryInfoKey(printer_path)[0]
                    for i in range(printers_len):
                        try:
                            printer = EnumKey(printer_path, i).replace(",", "\\").strip()
                            svr.append(printer)
                        except FileNotFoundError:
                            pass
            except FileNotFoundError:
                pass
            except:
                log()

        found = []
        with ConnectRegistry(to_export, HKEY_LOCAL_MACHINE) as reg_:
            with OpenKey(reg_, r'SYSTEM\CurrentControlSet\Control\Print\Printers') as printers:
                printers_len = QueryInfoKey(printers)[0]
                for i in range(printers_len):
                    with OpenKey(printers, EnumKey(printers, i)) as printer:
                        try:
                            prnt = QueryValueEx(printer, "Port")[0]
                            driver = QueryValueEx(printer, "Printer Driver")[0]
                            name = QueryValueEx(printer, "Name")[0]
                            if "_" in prnt:
                                prnt = prnt.split("_")[0].strip()
                            if prnt in found:
                                continue
                            found.append(prnt)
                            int(prnt.split(".")[0])
                            tcp.append((prnt.strip(), name, driver))
                        except (FileNotFoundError, ValueError):
                            pass
                        except:
                            log()

            with OpenKey(reg_, r"SYSTEM\CurrentControlSet\Enum\SWD\PRINTENUM") as printers:
                printers_len = QueryInfoKey(printers)[0]
                for i in range(printers_len):
                    with OpenKey(printers, EnumKey(printers, i)) as printer:
                        try:
                            prnt = QueryValueEx(printer, "LocationInformation")[0].split("/")[2].split(":")[0]
                            if prnt in found:
                                continue
                            if "_" in prnt:
                                prnt = prnt.split("_")[0]
                            found.append(prnt)
                            int(prnt.split(".")[0])
                            wsd.append(prnt.strip())
                        except (FileNotFoundError, ValueError, IndexError):
                            pass
                        except:
                            pass
        drives = []
        if config.current_user:
            with OpenKey(reg, fr"{get_sid(config.current_user)}\Network") as reg_:
                drives_len = QueryInfoKey(reg_)[0]
                for drive in range(drives_len):
                    try:
                        drive = EnumKey(reg_, drive)
                        with OpenKey(reg_, drive) as key_reg:
                            cur_drive = QueryValueEx(key_reg, "RemotePath")
                            drives.append((drive, cur_drive[0]))
                    except (FileNotFoundError, ValueError, IndexError):
                        pass
                    except:
                        log()
        svr_inst = svr
        tcp_manual = ["\n\nTCP/IP printers - NEED TO INSTALL MANUALLY\n"]
        wsd_manual = ["\n\nWSD printers - NEED TO INSTALL MANUALLY\n"]
        for tcp_prnt in tcp:
            if tcp_prnt[0] in config.ip_printers:
                svr_inst.append(config.ip_printers[tcp_prnt[0]])
            else:
                tcp_manual.append(tcp_prnt)
                signals.print(f"TCP/IP printer {tcp_prnt[0]} needs to be installed manually")
        for wsd_prnt in wsd:
            if wsd_prnt in config.ip_printers:
                svr_inst.append(config.ip_printers[wsd_prnt])
            else:
                wsd_manual.append(wsd_prnt)
                signals.print(f"WSD printer {wsd_prnt} needs to be installed manually")
        path_ = fr"C:\users\{config.user}\desktop\{config.current_computer}_backup"
        if not any(len(lst) > 1 for lst in (wsd, tcp, svr, drives)):
            signals.print_error(Objects.console, "No printers or network drives were found")
            return
        if not path.exists(path_):
            mkdir(path_)
        newline = "\n"
        with open(fr"{path_}\{config.current_computer}_log.txt", "w", encoding="utf-8") as txt:
            if svr:
                txt.write(f"Network printers\n\n{newline.join(svr)}")
            if tcp:
                txt.write("\n\nTCP/IP printers \n")
                for prnt in tcp:
                    if prnt[0] in config.ip_printers:
                        txt.write(f"\n{prnt[0]} - {config.ip_printers[prnt[0]]} - {prnt[1]} - {prnt[2]}")
                    else:
                        txt.write(f"\n{prnt[0]} - {prnt[1]} - {prnt[2]}")
            if wsd:
                txt.write("\n\nWSD printers\n")
                for prnt in wsd:
                    if prnt in config.ip_printers:
                        txt.write(f"\n{prnt} - {config.ip_printers[prnt]}")
                    else:
                        txt.write(str(prnt))
            if drives:
                txt.write("\n\nNetwork drives\n")
                for drive in drives:
                    txt.write(f"\n{drive[0]} - {drive[1]}")
            if len(tcp_manual) > 1:
                txt.write(tcp_manual[0])
                for prnt in tcp_manual[1:]:
                    txt.write(f"\n{prnt[0]} - {prnt[1]} - {prnt[2]}")
            if len(wsd_manual) > 1:
                txt.write(wsd_manual[0])
                for prnt in wsd_manual[1:]:
                    txt.write(prnt)
        with open(fr"{path_}\{config.current_computer}_run.bat", "w", encoding="utf-8") as bat:
            bat.write("""
@echo off
echo choose from the following
echo 1. Install network printers
echo 2. Map Network drives
echo 3. Both
set /p opt=:
if %opt%==1 (goto printers)
if %opt%==3 (goto printers)
if %opt%==2 (goto drives)
echo Invalid option %opt%
pause
exit
:printers\n""")
            if not svr_inst:
                bat.write("echo No printers were found to install\ngoto finish")
            else:
                for prnt in svr_inst:
                    bat.write(f"rundll32 printui.dll,PrintUIEntry /in /n {prnt}\n"
                              f"if %errorlevel% NEQ 0 (echo ERROR could not install {prnt}) "
                              f"else (echo installed {prnt})\n")
            bat.write("if not %opt%==3 (echo Done installing printers && pause && exit)\n")
            bat.write("goto drives\n")
            bat.write(":drives\n")
            if drives:
                for drive in drives:
                    bat.write(f"net use {drive[0]}: {drive[1]} /persistent:yes > nul 2>&1\n")
                    bat.write(f"if %errorlevel% NEQ 0 (echo Failed to map {drive[1]} to {drive[0]} check that "
                              f"{drive[0]} is not already in use) else "
                              f"(echo {drive[1]} was mapped to {drive[0]})\n")
            else:
                bat.write("echo no drives were found to map\n")
            bat.write("goto finish\n")
            bat.write(":finish\npause && exit")
            signals.print_success(Objects.console, f"Exported to {path_}")


def update_user(user_: str, signals: PassSignals) -> None:
    """updates the user status to the user_active Text box"""
    try:
        user_s = query_user(user_, signals)
        if user_s == 0:
            signals.update_success(Objects.user_status, "Active")
            config.copy = False
        elif user_s == 1:
            signals.update_error(Objects.user_status, "Disabled")
        elif user_s == 3:
            signals.update_error(Objects.user_status, "Expired")
        elif user_s == 4:
            signals.update_error(Objects.user_status, "Password expired")
        elif user_s == 5:
            signals.update_error(Objects.user_status, "No logon hours")
        else:
            signals.update_error(Objects.user_status, "ERROR")
    except:
        signals.update_error(Objects.user_status, "ERROR")
        log()


def on_submit(signals: PassSignals, pc: str = None, passed_user: str = None) -> None:
    """checks if the passed string is a computer in the domain
    if it is, it checks if its online, if it is online it then proceed to display information on the computer
    if the passed string is not a computer in the domain it looks for a file with the same name and txt extension
    in the preconfigured path via the config file, if it finds any it treats the contents of the file as the computer
    name and rerun on_submit with the computer as the arg
    if the string is neither a username nor a computer name it checks if it's a printer - TCP/IP or installed via
    print server"""
    signals.clear_all()
    config.comp_online = False
    config.copy = True
    config.current_user = None
    config.current_computer = None
    checked_user = False
    config.users_equal = False
    config.disable_user_depends = False
    t = None
    if not pc:
        pc = ui.computer_entry.text().strip()
        config.prev_index = -2
    if not pc:
        config.disable = True
        return
    pythoncom.CoInitialize()
    for is_pc in (pc, f"m{pc}-w10", f"m{pc}", f"{pc}-w10"):
        if not pc_in_domain(is_pc):
            continue
        if passed_user:
            t = Thread(target=update_user, args=[passed_user, signals], daemon=True)
            t.start()
            checked_user = True
        pc = is_pc
        signals.enable_1(Objects.copy_btn)
        if config.copy:
            signals.copy(pc)
        config.current_computer = pc
        signals.update(Objects.pc_display, pc)
        if not check_pc_active(pc):
            signals.update_error(Objects.pc_status, "OFFLINE")
            config.disable = True
            if passed_user:
                config.users_equal = passed_user.lower().strip()
            return
        config.current_computer = pc
        if not wmi_connectable():
            signals.print_error(Objects.console, "Could not connect to computer's WMI")
            signals.enable_1(Objects.submit_btn)
            config.disable = True
            if passed_user:
                config.users_equal = passed_user.lower().strip()
            return
        if not reg_connect():
            signals.print_error(Objects.console, "Could not connect to computer's registry")
            signals.enable_1(Objects.submit_btn)
            config.disable = True
            if passed_user:
                config.users_equal = passed_user.lower().strip()
            return

        signals.update_success(Objects.pc_status, "ONLINE")
        user_ = get_username(pc)
        if passed_user:
            if not user_ or passed_user != user_:
                config.users_equal = passed_user.lower().strip()
            if str(passed_user).lower().strip() != str(user_).lower().strip():
                config.disable_user_depends = True
        if user_:
            config.current_user = user_
            if not passed_user or str(passed_user).lower() == str(user_).lower():
                signals.update(Objects.user_display, f"{user_name_translation(user_)[:30]}")
            else:
                signals.update_error(Objects.user_display, user_)
        else:
            signals.update_error(Objects.user_display, "No user")
            if passed_user:
                config.users_equal = str(passed_user).lower().strip()
        try:
            r_pc = WMI(pc)
            for k in r_pc.Win32_OperatingSystem():
                last_boot_time = datetime.strptime(k.LastBootUpTime.split('.')[0], '%Y%m%d%H%M%S')
                current_time = datetime.strptime(k.LocalDateTime.split('.')[0], '%Y%m%d%H%M%S')
                uptime_ = current_time - last_boot_time
                if uptime_ > timedelta(days=7):
                    signals.update_error(Objects.uptime_display, str(uptime_).strip())
                else:
                    signals.update(Objects.uptime_display, str(uptime_).strip())
                break
        except Exception as er:
            signals.update_error(Objects.uptime_display, "ERROR")
            if er not in (AttributeError, pywintypes.com_error):
                log()

        try:
            space = get_space(pc)
            if space <= 5:
                signals.update_error(Objects.c_space_display, f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
            else:
                signals.update(Objects.c_space_display, f"{space:.1f}GB free out of {get_total_space(pc):.1f}GB")
        except:
            log()
            signals.update_error(Objects.c_space_display, "ERROR")
        sleep(0.1)

        if path.exists(fr"\\{pc}\d$"):
            try:
                spaced = get_space(pc, disk="d")
                if spaced <= 5:
                    signals.update_error(Objects.d_space_display,
                                         f"{spaced:.1f}GB free out of {get_total_space(pc, disk='d'):.1f}GB")
                else:
                    signals.update(Objects.d_space_display,
                                   f"{spaced:.1f}GB free out of {get_total_space(pc, disk='d'):.1f}GB")
            except:
                log()
                signals.update_error(Objects.d_space_display, f"ERROR")
        else:
            signals.update_error(Objects.d_space_display, f"Does not exist")
        try:
            try:
                r_pc
            except NameError:
                r_pc = WMI(pc)
            for ram_ in r_pc.Win32_ComputerSystem():
                total_ram = int(ram_.TotalPhysicalMemory) / (1024 ** 3)
                if total_ram < 7:
                    signals.update_error(Objects.ram_display, f"{round(total_ram)}GB")
                else:
                    signals.update(Objects.ram_display, f"{round(total_ram)}GB")
        except Exception as er:
            signals.update_error(Objects.ram_display, "ERROR")
            if er not in (AttributeError, pywintypes.com_error):
                log()
        if is_ie_fixed(pc):
            signals.update(Objects.ie_display, "Fixed")
        else:
            signals.update_error(Objects.ie_display, "Not fixed")

        if is_cpt_fixed(pc):
            signals.update(Objects.cpt_status, "Fixed")
        else:
            signals.update_error(Objects.cpt_status, "Not fixed")

        if (user_ or passed_user) and not checked_user:
            if passed_user:
                user_ = passed_user
            update_user(user_, signals)
        else:
            if not checked_user:
                signals.update_error(Objects.user_status, "No user")
        if not config.current_computer:
            config.current_computer = pc
        if t:
            if t.is_alive():
                t.join()
        config.disable = False
        config.comp_online = True
        return

    else:
        if passed_user:
            signals.print_error(Objects.console, "The computer is not in the domain anymore")
            signals.update(Objects.pc_display, pc)
            signals.update_error(Objects.pc_status, "NOT IN DOMAIN")
            signals.copy(pc[1:-4])
            update_user(passed_user, signals)
            config.disable = True
            return
        try:
            with open(f"{config.users_txt}\\{pc}.txt") as pc_file:
                user_ = pc
                pc = pc_file.read().strip()
                on_submit(pc=pc, passed_user=user_, signals=signals)
                return
        except (FileNotFoundError, AttributeError):
            config.disable = True
            if user_exists(pc):
                signals.print_error(
                    Objects.console, f"Could not locate the current or last omputer {pc} has logged on to"
                )
                update_user(pc, signals)
                config.users_equal = pc.lower().strip()
            else:
                if any([pc.lower() in config.ip_printers, pc.lower() in config.svr_printers]):
                    pc = pc.lower()
                    if pc in config.ip_printers:
                        pr = pc
                        signals.print(f"Printer with an IP of {pc} is at {config.ip_printers[pr]}")
                        pc = config.ip_printers[pc]
                    elif pc in config.svr_printers:
                        signals.print(f"Printer {pc} has an ip of {config.svr_printers[pc]}")
                        pc = config.svr_printers[pc]
                    signals.copy(pc)
                else:
                    if r"\\" in pc:
                        pr = pc
                        signals.print_error(Objects.console, f"Could not locate printer {pr}")
                    elif pc.count(".") > 2:
                        signals.print_error(Objects.console, f"Could not locate TCP/IP printer with ip of {pc}")
                    else:
                        signals.print_error(Objects.console, f"No such user or computer in the domain {pc}")
            return
        except OSError:
            config.disable = True
            signals.print_error(Objects.console, "Invalid computer name")
            return


class SetConfig:
    """"sets the basic config for the script to run on, these configs are needed in order for the script to run"""

    def __init__(self, json_config_file: dict) -> None:
        self.config_file = json_config_file
        self.ip_printers = {}
        self.svr_printers = {}
        self.log = self.config_file["log"]
        self.host = environ["COMPUTERNAME"]
        self.user = getuser().lower()
        self.domain = self.config_file["domain"]
        self.delete_edb = self.config_file["delete_edb"]
        self.delete_user_temp = self.config_file["delete_user_temp"]
        self.c_paths_with_msg = [path_with_msg for path_with_msg in self.config_file["to_delete"] if
                                 len(path_with_msg) > 1]
        self.c_paths_without_msg = [path_without_msg for path_without_msg in self.config_file["to_delete"] if
                                    len(path_without_msg) == 1]
        self.u_paths_with_msg = [path_with_msg for path_with_msg in self.config_file["user_specific_delete"] if
                                 len(path_with_msg) > 1]
        self.u_paths_without_msg = [path_without_msg for path_without_msg in self.config_file["user_specific_delete"] if
                                    len(path_without_msg) == 1]
        self.mx_w = self.config_file["max_workers"]
        self.current_computer = None
        self.current_user = None
        self.exclude = self.config_file["do_not_delete"]
        self.startwith_exclude = self.config_file["start_with_exclude"]
        self.will_delete = []
        self.yes_no = False
        self.users_equal = False
        self.wmi_connectable = False
        self.reg_connectable = False
        self.comp_online = False
        self.assets = self.config_file["assets"].replace("/", "\\")
        self.disable = False
        self.first_time = 1
        self.current_sid = None
        self.copy = True
        self.zoom = 3
        self.temp = environ["temp"]
        self.lr = 0
        self.interaction_done = False
        self.disable_user_depends = False
        self.title = self.config_file["title"]


def update_prnts():
    """A scheduled update of the printers in memory against the print servers"""
    while True:
        for svr in config.config_file["print_servers"]:
            try:
                ip_svr = {prnt_name[2].strip().lower(): fr"{svr.lower()}\{prnt_name[0].strip()}" for prnt_name in
                          NetShareEnum(svr)}
                prnt_svr = {fr"{svr.lower()}\{prnt_name[0].strip().lower()}": prnt_name[2].strip() for prnt_name in
                            NetShareEnum(svr)}
                config.ip_printers.update(ip_svr)
                config.svr_printers.update(prnt_svr)
            except:
                pass
        sleep(1200)


try:
    with open("GUI_config.json", encoding="utf8") as config_file:
        config = SetConfig(load(config_file))
        Thread(target=update_prnts, daemon=True).start()
except FileNotFoundError:
    try:
        FatalError(title="No GUI_config.json", prompt_="No GUI_config.json was found, exiting!")
        sys.exit(1)
    except Exception:
        sys.exit(69)

if config.log and not path.isfile(config.log):
    try:
        with open(config.log, "w") as _:
            pass
    except:
        config.log = ""

if config.log:
    basicConfig(filename=config.log, filemode="a", format="%(message)s")
    logger = getLogger("logfile")


def log() -> None:
    """logs the exceptions as well as date, host, and username to the logfile"""
    if not config.log:
        return
    err_log = f"""{'_' * 145}\nat {datetime.now().strftime('%Y-%m-%d %H:%M')} an error occurred on {config.host}\
 - {config.user}\n"""
    exception(err_log)


if not config.log:
    basicConfig(filename="FATAL_errors.log", filemode="w", format="%(message)s")
    logger = getLogger("fatal exceptions")


class TimeoutException(Exception):
    pass


def Timeout(timeout: int | float) -> bool | TimeoutException | callable:
    """"run a function with timeout limit via threading"""

    def deco(func: callable):
        @wraps(func)
        def wrapper(*args, **kwargs):
            res = [TimeoutException('function [%s] timeout [%s seconds] exceeded!' % (func.__name__, timeout))]

            def newFunc():
                try:
                    res[0] = func(*args, **kwargs)
                except Exception as xo:
                    res[0] = xo

            t = Thread(target=newFunc, daemon=True)
            try:
                t.start()
                t.join(timeout)
            except TimeoutError:
                raise TimeoutException('Timeout occurred!')
            except Exception as je:
                log()
                raise je
            ret = res[0]
            if isinstance(ret, BaseException):
                raise ret
            return ret

        return wrapper

    return deco


# noinspection PyCallingNonCallable
def wmi_connectable() -> bool:
    """timed out test to check that the computer is connectable via WMI"""
    x = Timeout(timeout=1.5)(WMI_connectable_actual)
    try:
        y = x()
    except (TimeoutException, x_wmi):
        config.wmi_connectable = False
        return False
    except:
        log()
        config.wmi_connectable = False
        return False
    config.reg_connectable = True if y else False
    return y


def WMI_connectable_actual() -> bool:
    """"the actual WMI_connectable check"""
    pc = config.current_computer
    try:
        pythoncom.CoInitialize()
        WMI(computer=pc)
        return True
    except (pywintypes.com_error, x_wmi):
        return False
    except:
        log()
        return False


def get_space(pc: str, disk: str = "c") -> float:
    """"returns the free space in the disk in GB"""
    return disk_usage(fr"\\{pc}\{disk}$").free / (1024 ** 3)


def get_total_space(pc: str, disk: str = "c") -> float:
    """Return the total disk size"""
    return disk_usage(fr"\\{pc}\{disk}$").total / (1024 ** 3)


def pc_in_domain(pc: str) -> str | bool | None:
    """"query if the computer is in the domain"""
    ad = adquery.ADQuery()
    try:
        ad.execute_query(
            attributes=["name"],
            where_clause=f"name = '{pc}'",
            base_dn=config.domain
        )
    except pywintypes.com_error:
        return None
    except:
        log()
    try:
        result = ad.get_results()
    except pywintypes.com_error:
        return False
    is_pc = None
    try:
        for p in result:
            is_pc = p["name"]
        return is_pc
    except pywintypes.com_error:
        pass
    return is_pc


def user_exists(username_: str) -> bool:
    """checks if a user exists in the domain"""
    ad = adquery.ADQuery()
    try:
        ad.execute_query(
            attributes=["sAMAccountName"],
            where_clause=f"sAMAccountName='{username_}'",
            base_dn=config.domain
        )
        result = ad.get_results()
        for _ in result:
            return True
    except pywintypes.com_error:
        pass
    return False


def add_member(username_: str, group_name: str) -> None:
    """Adds a member to an AD group. there's no usage for it in the publicly published script"""
    try:
        group_name = adgroup.ADGroup.from_cn(group_name)
        user_cn = None
        ad = adquery.ADQuery()
        ad.execute_query(
            attributes=["cn"],
            where_clause=f"sAMAccountName='{username_}'",
            base_dn=config.domain
        )

        result = ad.get_results()
        for u in result:
            user_cn = u["cn"]
        user_cn = aduser.ADUser.from_cn(user_cn)
        if user_cn:
            group_name.add_members([user_cn])
    except pywintypes.com_error:
        pass
    except:
        log()


def query_user(username_: str, signals: PassSignals) -> int:
    """Checks the user status against the AD DC"""
    pythoncom.CoInitialize()
    username_ = username_.strip()
    ad = adquery.ADQuery()
    ad.execute_query(
        attributes=["accountExpires", "logonHours", "lockoutTime", "userAccountControl", "pwdLastSet"],
        where_clause=f"sAMAccountName='{username_}'",
        base_dn=config.domain
    )
    try:
        uso = list(ad.get_results())
    except pyadexceptions.invalidResults:
        return 6
    generic = "1970-01-01 07:00:00"
    now = datetime.now()
    try:
        expiration = pyadutils.convert_datetime(uso[0]["accountExpires"])
    except OSError:
        expiration = generic
    lock_time = pyadutils.convert_datetime(uso[0]["lockoutTime"]) if uso[0]["lockoutTime"] is not None else None
    passwd = pyadutils.convert_datetime(uso[0]["pwdLastSet"])

    if uso[0]['userAccountControl'] == 514:
        return 1

    if not expiration.__str__() == generic:
        if now > expiration:
            return 3

    if uso[0]["logonHours"] is not None:
        logon_hours = []
        for shift_1, shift_2, shift_3 in zip(*[iter(uso[0]["logonHours"].tobytes())] * 3):
            logon_hours.append(format(shift_1, '08b') + format(shift_2, '08b') + format(shift_3, '08b'))
        if sum("0" in day for day in logon_hours) > 2:
            return 5

    if passwd.__str__() == generic or passwd > now:
        return 4

    if lock_time and lock_time.__str__() != generic:
        Thread(target=lambda: run(["powershell", "-Command", f"Unlock-ADAccount -Identity {username_}"],
                                  shell=True, creationflags=CREATE_NO_WINDOW), daemon=True).start()
        signals.show_info("Unlocked user", f"User {username_} was unlocked")
        return 0
    return 0


def copy_pc():
    """Copy the ComputerName to clipboard"""
    try:
        copy_clip(config.current_computer.strip()[1:8])
    except:
        log()


def check_pc_active_actual(pc: str) -> bool:
    """checks if the computer is reachable via UNC pathing"""
    return path.exists(fr"\\{pc}\c$")


def check_pc_active(pc: str = None) -> bool:
    """"timed out check if the computer is online and reachable"""
    # noinspection PyCallingNonCallable
    x = Timeout(timeout=3)(check_pc_active_actual)
    try:
        y = x(pc=pc)
    except TimeoutException:
        return False
    except:
        log()
        return False

    return y


def get_username(pc: str) -> str | None:
    """retrieves the active user on the remote computer"""
    try:
        con = WMI(computer=pc)
        rec = con.query("SELECT * FROM Win32_ComputerSystem")
        for user_ in rec:
            try:
                user_ = user_.UserName.split("\\")[1]
                return user_
            except (AttributeError, pywintypes.com_error):
                pass
            except:
                log()
        try:
            processes = con.query("SELECT * FROM Win32_Process WHERE Name='explorer.exe'")
            for process in processes:
                _, _, user_ = process.GetOwner()
                return user_
        except:
            log()
    except pywintypes.com_error:
        return None
    except:
        log()


def get_sid(user_: str = None) -> str | bool | None:
    """Extract the sid of a user via its username"""
    if not user_:
        user_ = config.current_user
    if not user_:
        return

    with ConnectRegistry(config.current_computer, HKEY_USERS) as reg:
        sid_list = []
        with OpenKey(reg, "") as users:
            users_len = QueryInfoKey(users)[0]
            for i in range(users_len):
                try:
                    sid_list.append(EnumKey(users, i))
                except FileNotFoundError:
                    pass

    with ConnectRegistry(config.current_computer, HKEY_LOCAL_MACHINE) as users_path:
        for sid in set(sid_list):
            try:
                with OpenKey(users_path,
                             fr"SOFTWARE\Microsoft\Windows NT\\CurrentVersion\ProfileList\{sid}") as profiles:
                    username = QueryValueEx(profiles, "ProfileImagePath")
                    if username[0].startswith("C:\\"):
                        username = username[0].split("\\")[-1]
                        if not username:
                            continue
                        if user_.lower() == username.lower():
                            config.current_sid = sid
                            return sid
            except FileNotFoundError:
                pass
        return False


def user_name_translation(username_: str) -> str | None:
    """returns the display name of a user in the domain"""
    ad = adquery.ADQuery()
    ad.execute_query(
        attributes=["displayName"],
        where_clause=f"sAMAccountName='{username_}'",
        base_dn=config.domain
    )
    result = ad.get_results()
    for u in result:
        return u["displayName"]
    return username_


def date_is_older(date_string: datetime.strptime) -> bool:
    """checks if the initial date has passed"""
    provided_date = datetime.strptime(date_string, "%d/%m/%Y %H:%M:%S")
    return provided_date < datetime.now()


def on_rm_error(_: callable, path_: str, error: tuple) -> None:
    """deletes readonly files via changing permission to the file or directory"""
    if error[0] != PermissionError:
        return
    try:
        chmod(path_, S_IWRITE)
        if path.isfile(path_):
            unlink(path_)
        elif path.isdir(path_):
            rmtree(path_, ignore_errors=True)
    except (PermissionError, FileNotFoundError):
        pass
    except:
        log()


def reg_connect() -> bool:
    """timed check if the registry is connectable"""
    x = Timeout(timeout=1.5)(is_reg)
    try:
        y = x()
    except TimeoutException:
        config.reg_connectable = False
        return False
    except:
        config.reg_connectable = False
        log()
        return False
    config.reg_connectable = True if y else False
    return y


def is_reg(pc: str = None) -> bool:
    """checks if the remote registry is connectable"""
    if not pc:
        pc = config.current_computer
    try:
        with ConnectRegistry(pc, HKEY_USERS) as _:
            return True
    except FileNotFoundError:
        pass
    except (PermissionError, OSError):
        return False
    except:
        log()
    return False


def is_ie_fixed(pc: str) -> bool:
    """checks if Internet Explorer is fixed via querying the registry keys"""
    try:
        with ConnectRegistry(pc, HKEY_LOCAL_MACHINE) as reg_:
            try:
                with OpenKey(reg_, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper "
                                   r"Objects\{1FD49718-1D00-4B19-AF5F-070AF6D5D54C}", 0, KEY_ALL_ACCESS) as _:
                    return False
            except FileNotFoundError:
                pass
            except:
                log()
            return True
    except FileNotFoundError:
        pass
    except:
        log()
    return False


def show_info(title: str, message: str):
    """A function which initializes the ShowInfo class"""
    ShowInfo(title, message).exec()
    config.interaction_done = True


def del_users_selection_actual(users_list: list):
    """A function which initializes the UserDeletion class"""
    UserDeletion(users_list).exec()


def deletion_selection(users_list: list, signals: PassSignals):
    """A function which waits for the UserDeletion class"""
    config.wll_delete = []
    config.interaction_done = False
    signals.del_users(users_list)
    while not config.interaction_done:
        sleep(0.3)
    return config.will_delete


def settings_select():
    """A function which initializes the SettingsSelect class"""
    SettingsSelection().exec()


def is_cpt_fixed(pc: str) -> bool:
    """checks if cockpit printer is fixed via querying the registry keys"""
    sid_ = get_sid()
    if not sid_:
        return True
    try:
        with ConnectRegistry(pc, HKEY_USERS) as reg_:
            try:
                with OpenKey(reg_, fr"{sid_}\SOFTWARE\Jetro Platforms\JDsClient\PrintPlugIn", 0,
                             KEY_ALL_ACCESS) as key_:
                    QueryValueEx(key_, "PrintClientPath")
                    return False
            except FileNotFoundError:
                return True
            except:
                log()
    except FileNotFoundError:
        pass
    except:
        log()
    return False


def zoom_select(signals: PassSignals):
    """A function which initializes the zoom class"""
    config.zoom = 3
    config.interaction_done = False
    signals.zoom()
    while not config.interaction_done:
        sleep(0.3)
    return config.zoom


def delete_zoom(signals: PassSignals) -> None:
    """Remove zoom according to selection"""
    prmt = "Removing Zoom"
    prns = ""
    q_rm = []
    opt = zoom_select(signals)
    if opt == 3:
        signals.print("Canceled Zoom deletion")
        return
    elif opt == 0:
        q_rm = [rem_addon, rem_zoom_dir, rem_zoom_64, rem_zoom_32]
        prns = "Successfully Deleted Zoom and Addon"
    elif opt == 1:
        q_rm = [rem_addon]
        prmt = "Removing addon"
        prns = "Successfully Deleted Zoom addon"
    elif opt == 2:
        q_rm = [rem_zoom_dir, rem_zoom_64, rem_zoom_32]
        prns = "Successfully Deleted Zoom"
    signals.print("this will take a few minutes")
    with ProgressBar(len(q_rm), prmt, "", signals) as bar_:
        with ThreadPoolExecutor(max_workers=config.mx_w) as executor:
            jobs = [executor.submit(func_, bar_) for func_ in q_rm]
            while not all([result.done() for result in jobs]):
                sleep(0.3)
    signals.print_success(Objects.console, prns)


def rem_zoom_64(bar_: ProgressBar) -> None:
    """Attempts to remove Zoom(64) bit machine wide installation"""
    pythoncom.CoInitialize()
    con = WMI(computer=config.current_computer)
    pg = con.Win32_Product(name="Zoom(64bit)")
    for pog in pg:
        try:
            pog.Uninstall()
        except:
            log()
    bar_()


def rem_zoom_32(bar_: ProgressBar) -> None:
    """Attempts to remove Zoom(32) bit machine wide installation"""
    pythoncom.CoInitialize()
    con = WMI(computer=config.current_computer)
    pg = con.Win32_Product(name="Zoom(32bit)")
    for pog in pg:
        try:
            pog.Uninstall()
        except:
            log()
    bar_()


def rem_addon(bar_: ProgressBar) -> None:
    """Attempts to remove Zoom Outlook addon"""
    pythoncom.CoInitialize()
    rem_reg_addon()
    bar_()
    con = WMI(computer=config.current_computer)
    pg = con.Win32_Product(name="Zoom Outlook Plugin")
    for pog in pg:
        try:
            pog.Uninstall()
        except:
            log()
    bar_()


def rem_zoom_dir(bar_: ProgressBar) -> None:
    """Removes the directory of user-specific zoom installation"""
    pythoncom.CoInitialize()
    rem_reg_zoom()
    con = WMI(computer=config.current_computer)
    for exe in ("outlook.exe", "Zoom.exe"):
        for process in con.Win32_Process(name=exe):
            try:
                process.Terminate()
            except:
                log()
    zoom = rf"\\{config.current_computer}\c$\Users\{config.current_user}\AppData\Roaming\Zoom"
    sleep(1)
    rmtree(zoom, ignore_errors=True)
    bar_()


def del_sub(root_key, key1, key2=None) -> None:
    """Deletes sub-keys in the registry"""
    if key2 is None:
        c_key = key1
    else:
        c_key = fr"{key1}\{key2}"
    with OpenKey(root_key, c_key, 0, KEY_ALL_ACCESS) as keys:
        ikey = QueryInfoKey(keys)
        for x in range(0, ikey[0]):
            s_key = EnumKey(keys, 0)
            try:
                DeleteKey(keys, s_key)
            except (PermissionError, FileNotFoundError):
                del_sub(root_key, c_key, s_key)
            except:
                log()
        DeleteKey(keys, "")


def rem_reg_addon() -> None:
    """Removes the zoom addon from the registry"""
    del_it = []
    with ConnectRegistry(config.current_computer, HKEY_LOCAL_MACHINE) as reg:
        with OpenKey(reg, r"SOFTWARE\Classes\Installer\Products") as products:
            prod_len = QueryInfoKey(products)[0]
            for i in range(prod_len):
                try:
                    with OpenKey(reg, fr"SOFTWARE\Classes\Installer\Products\{EnumKey(products, i)}") as to_del:
                        try:
                            if QueryValueEx(to_del, "ProductName")[0].lower() == "zoom outlook plugin":
                                del_it.append(fr"SOFTWARE\Classes\Installer\Products\{EnumKey(products, i)}")
                        except FileNotFoundError:
                            pass
                except (FileNotFoundError, PermissionError):
                    pass
                except:
                    log()
        if del_it:
            for z_key in del_it:
                try:
                    del_sub(reg, z_key)
                except:
                    log()


def rem_reg_zoom():
    """Removes the zoom application from the registry"""
    del_it = []
    with ConnectRegistry(config.current_computer, HKEY_LOCAL_MACHINE) as reg:
        with OpenKey(reg, r"SOFTWARE\Classes\Installer\Products") as products:
            prod_len = QueryInfoKey(products)[0]
            for i in range(prod_len):
                try:
                    with OpenKey(reg, fr"SOFTWARE\Classes\Installer\Products\{EnumKey(products, i)}") as to_del:
                        try:
                            if QueryValueEx(to_del, "ProductName")[0].lower() in ("zoom",
                                                                                  "zoom (64-bit)", "zoom (32-bit)",
                                                                                  "zoom(64-bit)", "zoom(32-bit)"):
                                del_it.append(fr"SOFTWARE\Classes\Installer\Products\{EnumKey(products, i)}")
                        except FileNotFoundError:
                            pass
                except (FileNotFoundError, PermissionError):
                    pass
                except:
                    log()
        if del_it:
            for z_key in del_it:
                try:
                    del_sub(reg, z_key)
                except:
                    log()


def restart(signals: PassSignals) -> None:
    """Restarts the remote computer"""
    pythoncom.CoInitialize()
    with AskYesNo("Restart computer", f"Are you sure you want to restart {config.current_computer}",
                  signals):
        pass
    if WorkerSignals.yes_no == (False, True):
        signals.print(f"Canceled the restart of {config.current_computer}")
        return
    WMI(config.current_computer).Win32_OperatingSystem()[0].Reboot()
    signals.print_success(Objects.console, f"Attempted to restart {config.current_computer}")
    config.disable = True


write_colors = True
if path.isfile(fr"{environ['USERPROFILE']}\AppData\Local\UIV2_config\color.json"):
    """Checks if all colors are in place"""
    try:
        with open(fr"{environ['USERPROFILE']}\AppData\Local\UIV2_config\color.json") as color_file_:
            config.colors = load(color_file_)
        assert "background" in config.colors
        assert "buttons" in config.colors
        assert "text_area" in config.colors
        assert "text_color" in config.colors
        assert "dark" in config.colors
        write_colors = False
    except AssertionError:
        pass
    except:
        log()
if write_colors:
    print("writing new config file")
    try:
        if not path.isdir(fr"{environ['USERPROFILE']}\AppData\Local\UIV2_config"):
            mkdir(fr"{environ['USERPROFILE']}\AppData\Local\UIV2_config")
        with open(fr"{environ['USERPROFILE']}\AppData\Local\UIV2_config\color.json", "w") as colors_file:
            light_mode = {"background": "(238, 238, 245)", "buttons": "a(249, 248, 245, 243)", "text_area":
                "(255, 255, 255)", "text_color": "(0, 0, 0)", "dark": False}
            config.colors = light_mode
            dump(light_mode, colors_file)
    except:
        log()


class SettingsSelection(QtWidgets.QDialog):
    """The UI for the settings"""
    def __init__(self):
        super().__init__()
        self.setWindowIcon(ui.icon_light)
        self.setObjectName("settings_selection")
        self.setFixedSize(411, 210)
        self.icon = ui.icon_dark if config.colors["dark"] else ui.icon_light
        self.setModal(True)
        self.temp = config.colors.copy()
        self.setFocusPolicy(QtCore.Qt.FocusPolicy.TabFocus)
        self.buttons_choose = QtWidgets.QPushButton(parent=self)
        self.buttons_choose.setGeometry(QtCore.QRect(160, 110, 101, 31))
        self.buttons_choose.setFont(Fonts.ariel_12_bold)
        self.buttons_choose.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.buttons_choose.setMouseTracking(False)
        self.buttons_choose.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.buttons_choose.setAutoFillBackground(False)
        self.buttons_choose.setObjectName("deselect_all")
        self.buttons_choose.clicked.connect(self.change_buttons_bg)
        self.buttons_display = QtWidgets.QFrame(parent=self)
        self.buttons_display.setGeometry(QtCore.QRect(160, 10, 101, 81))
        self.buttons_display.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.buttons_display.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.buttons_display.setObjectName("buttons_display")
        self.text_choose = QtWidgets.QPushButton(parent=self)
        self.text_choose.setGeometry(QtCore.QRect(290, 110, 101, 31))
        self.text_choose.setFont(Fonts.ariel_12_bold)
        self.text_choose.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.text_choose.setMouseTracking(True)
        self.text_choose.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.text_choose.setAutoFillBackground(False)
        self.text_choose.setObjectName("text_choose")
        self.text_choose.clicked.connect(self.change_text_bg)
        self.text_display = QtWidgets.QFrame(parent=self)
        self.text_display.setGeometry(QtCore.QRect(290, 10, 101, 81))
        self.text_display.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.text_display.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.text_display.setObjectName("text_display")
        self.cancel = QtWidgets.QPushButton(parent=self)
        self.cancel.setGeometry(QtCore.QRect(290, 160, 101, 31))
        self.cancel.setFont(Fonts.ariel_12_bold)
        self.cancel.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.cancel.setMouseTracking(True)
        self.cancel.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.cancel.setAutoFillBackground(False)
        self.cancel.setObjectName("cancel")
        self.cancel.clicked.connect(self.done)
        self.save = QtWidgets.QPushButton(parent=self)
        self.save.setGeometry(QtCore.QRect(160, 160, 101, 31))
        self.save.setFont(Fonts.ariel_12_bold)
        self.save.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.save.setMouseTracking(True)
        self.save.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.save.setAutoFillBackground(False)
        self.save.setObjectName("save")
        self.save.clicked.connect(self.save_)
        self.light_mode = QtWidgets.QPushButton(parent=self)
        self.light_mode.setGeometry(QtCore.QRect(30, 110, 101, 31))
        self.light_mode.setFont(Fonts.ariel_12_bold)
        self.light_mode.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.light_mode.setMouseTracking(True)
        self.light_mode.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.light_mode.setAutoFillBackground(False)
        self.light_mode.setObjectName("light_mode")
        self.light_mode.clicked.connect(self.set_light_mode)
        self.dark_mode = QtWidgets.QPushButton(parent=self)
        self.dark_mode.setGeometry(QtCore.QRect(30, 160, 101, 31))
        self.dark_mode.setFont(Fonts.ariel_12_bold)
        self.dark_mode.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.dark_mode.setMouseTracking(True)
        self.dark_mode.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.dark_mode.setAutoFillBackground(False)
        self.dark_mode.setObjectName("dark_mode")
        self.dark_mode.clicked.connect(self.set_dark_mode)
        self.background_choose = QtWidgets.QPushButton(parent=self)
        self.background_choose.setGeometry(QtCore.QRect(30, 60, 101, 31))
        self.background_choose.setFont(Fonts.ariel_12_bold)
        self.background_choose.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.background_choose.setMouseTracking(True)
        self.background_choose.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.background_choose.setAutoFillBackground(False)
        self.background_choose.setObjectName("background_choose")
        self.background_choose.clicked.connect(self.change_background)
        self.fonts_color = QtWidgets.QPushButton(parent=self)
        self.fonts_color.setGeometry(QtCore.QRect(30, 10, 101, 31))
        self.fonts_color.setFont(Fonts.ariel_12_bold)
        self.fonts_color.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.fonts_color.setMouseTracking(True)
        self.fonts_color.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.fonts_color.setAutoFillBackground(False)
        self.fonts_color.setObjectName("fonts_color")
        self.fonts_color.clicked.connect(self.text_color)
        self.set_color()
        self.translate_ui(self)
        QtCore.QMetaObject.connectSlotsByName(self)
        self.show()

    @staticmethod
    def get_color():
        """Color picker"""
        color = QColorDialog.getColor()
        if color.isValid():
            return "a" + str(color.getRgb())

    def text_color(self):
        color = self.get_color()
        if not color:
            return
        self.temp["text_color"] = color
        self.set_color()

    def set_light_mode(self):
        self.temp = {
            "background": "(238, 238, 245)", "buttons": "a(249, 248, 245, 243)", "text_area": "(255, 255, 255)",
            "text_color": "(0, 0, 0)", "dark": False
        }
        self.icon = ui.icon_light
        self.set_color()

    def set_dark_mode(self):
        self.temp = {
            "background": "(32, 32, 32)", "buttons": "(38, 41, 55)", "text_area": "(51, 44, 61)",
            "text_color": "a(249, 248, 245, 243)", "dark": True
        }
        self.icon = ui.icon_dark
        self.set_color()

    def change_buttons_bg(self):
        color = self.get_color()
        if not color:
            return
        self.temp["buttons"] = color
        self.set_color()

    def change_text_bg(self):
        color = self.get_color()
        if not color:
            return
        self.temp["text_area"] = color
        self.set_color()

    def change_background(self):
        color = self.get_color()
        if not color:
            return
        self.temp["background"] = color
        self.set_color()

    def set_color(self):
        self.fonts_color.setStyleSheet("QPushButton {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 1px;\n"
                                       "    border-radius: 15px;\n"
                                       "    padding-left: 2px;\n"
                                       "    padding-right: 2px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:pressed {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 2px;\n"
                                       "    border-radius: 15px;\n"
                                       "    padding-left: 2px;\n"
                                       "    padding-right: 2px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    border-style: inset;\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "".format(bc=self.temp["buttons"], tc=self.temp["text_color"]))
        self.background_choose.setStyleSheet("QPushButton {{\n"
                                             "    border: 2px solid rgb(138, 138, 138);\n"
                                             "    border-radius: 15px;\n"
                                             "    padding-left: 2px;\n"
                                             "    padding-right: 2px;\n"
                                             "    background-color: rgb{bc};\n"
                                             "    color: rgb{tc};\n"
                                             "}}\n"
                                             "QPushButton:pressed {{\n"
                                             "    border: 2px solid rgb(138, 138, 138);\n"
                                             "    border-radius: 15px;\n"
                                             "    padding-left: 2px;\n"
                                             "    padding-right: 2px;\n"
                                             "    background-color: rgb{bc};\n"
                                             "    border-style: inset;\n"
                                             "    color: rgb{tc};\n"
                                             "}}\n"
                                             "".format(bc=self.temp["buttons"], tc=self.temp["text_color"]))
        self.dark_mode.setStyleSheet("QPushButton {\n"
                                     "    border: 2px solid rgb(138, 138, 138);\n"
                                     "    border-width: 1px;\n"
                                     "    border-radius: 15px;\n"
                                     "    padding-left: 2px;\n"
                                     "    padding-right: 2px;\n"
                                     "    background-color: rgb(32, 32, 32);\n"
                                     "    color: rgb(255, 255, 255);\n"
                                     "}\n"
                                     "QPushButton:pressed {\n"
                                     "    border: 2px solid rgb(138, 138, 138);\n"
                                     "    border-width: 2px;\n"
                                     "    border-radius: 15px;\n"
                                     "    padding-left: 2px;\n"
                                     "    padding-right: 2px;\n"
                                     "    background-color:  rgb(32, 32, 32);\n"
                                     "    border-style: inset;\n"
                                     "}\n"
                                     "")
        self.light_mode.setStyleSheet("QPushButton {\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 15px;\n"
                                      "    padding-left: 2px;\n"
                                      "    padding-right: 2px;\n"
                                      "    background-color: rgba(249, 248, 245, 243);\n"
                                      "}\n"
                                      "QPushButton:pressed {\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 2px;\n"
                                      "    border-radius: 15px;\n"
                                      "    padding-left: 2px;\n"
                                      "    padding-right: 2px;\n"
                                      "    background-color: rgba(249, 248, 245, 243);\n"
                                      "    border-style: inset;\n"
                                      "}\n"
                                      "")
        self.save.setStyleSheet("QPushButton {{\n"
                                "    border: 2px solid rgb(138, 138, 138);\n"
                                "    border-width: 1px;\n"
                                "    border-radius: 15px;\n"
                                "    padding-left: 2px;\n"
                                "    padding-right: 2px;\n"
                                "    background-color: rgb{bc};\n"
                                "    color: rgb{tc};\n"
                                "}}\n"
                                "QPushButton:pressed {{\n"
                                "    border: 2px solid rgb(138, 138, 138);\n"
                                "    border-width: 2px;\n"
                                "    border-radius: 15px;\n"
                                "    padding-left: 2px;\n"
                                "    padding-right: 2px;\n"
                                "    background-color: rgb{bc};\n"
                                "    border-style: inset;\n"
                                "    color: rgb{tc};\n"
                                "}}\n"
                                "".format(bc=self.temp["buttons"], tc=self.temp["text_color"]))
        self.cancel.setStyleSheet("QPushButton {{\n"
                                  "    border: 2px solid rgb(138, 138, 138);\n"
                                  "    border-width: 1px;\n"
                                  "    border-radius: 15px;\n"
                                  "    padding-left: 2px;\n"
                                  "    padding-right: 2px;\n"
                                  "    background-color: rgb{bc};\n"
                                  "    color: rgb{tc};\n"
                                  "}}\n"
                                  "QPushButton:pressed {{\n"
                                  "    border: 2px solid rgb(138, 138, 138);\n"
                                  "    border-width: 2px;\n"
                                  "    border-radius: 15px;\n"
                                  "    padding-left: 2px;\n"
                                  "    padding-right: 2px;\n"
                                  "    background-color: rgb{bc};\n"
                                  "    border-style: inset;\n"
                                  "    color: rgb{tc};\n"
                                  "}}\n"
                                  "".format(bc=self.temp["buttons"], tc=self.temp["text_color"]))
        self.text_display.setStyleSheet("QFrame {{ \n"
                                        "    background-color: rgb{bc};\n"
                                        "    border: 2px solid rgb(138, 138, 138);\n"
                                        "    border-radius: 15px\n"
                                        "}}".format(bc=self.temp["text_area"]))
        self.text_choose.setStyleSheet("QPushButton {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 1px;\n"
                                       "    border-radius: 15px;\n"
                                       "    padding-left: 2px;\n"
                                       "    padding-right: 2px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:pressed {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 2px;\n"
                                       "    border-radius: 15px;\n"
                                       "    padding-left: 2px;\n"
                                       "    padding-right: 2px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    border-style: inset;\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "".format(bc=self.temp["buttons"], tc=self.temp["text_color"]))
        self.buttons_display.setStyleSheet("QFrame {{\n"
                                           "    background-color: rgb{bc};\n"
                                           "    border: 2px solid rgb(138, 138, 138);\n"
                                           "    border-radius: 15px;\n"
                                           "}}".format(bc=self.temp["buttons"]))
        self.buttons_choose.setStyleSheet("QPushButton {{\n"
                                          "    border: 2px solid rgb(138, 138, 138);\n"
                                          "    border-width: 1px;\n"
                                          "    border-radius: 15px;\n"
                                          "    padding-left: 2px;\n"
                                          "    padding-right: 2px;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    color: rgb{tc};\n"
                                          "}}\n"
                                          "QPushButton:pressed {{\n"
                                          "    border: 2px solid rgb(138, 138, 138);\n"
                                          "    border-width: 2px;\n"
                                          "    border-radius: 15px;\n"
                                          "    padding-left: 2px;\n"
                                          "    padding-right: 2px;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    border-style: inset;\n"
                                          "    color: rgb{tc};\n"
                                          "}}\n"
                                          "".format(bc=self.temp["buttons"], tc=self.temp["text_color"]))
        self.setStyleSheet("QDialog {{\n"
                           "    background-color: rgb{bc};\n"
                           "}}".format(bc=self.temp["background"]))

    def save_(self):
        config.colors = self.temp.copy()
        ui.settings_btn.setIcon(self.icon)
        Formats.normal = '<span style="color:rgb%%;">{}</span>'.replace("%%", config.colors["text_color"])
        with open(fr"{environ['USERPROFILE']}\AppData\Local\UIV2_config\color.json", "w") as colors_file_:
            dump(config.colors, colors_file_)
        ui.set_colors()
        self.done()

    def done(self, code: int = 0):
        self.close()
        self.deleteLater()

    def translate_ui(self, settings_selection):
        _translate = QtCore.QCoreApplication.translate
        settings_selection.setWindowTitle(_translate("settings_selection", "Colors selection"))
        self.buttons_choose.setText(_translate("settings_selection", "Buttons"))
        self.text_choose.setText(_translate("settings_selection", "Text areas"))
        self.cancel.setText(_translate("settings_selection", "Cancel"))
        self.save.setText(_translate("settings_selection", "Save"))
        self.light_mode.setText(_translate("settings_selection", "Light Mode"))
        self.dark_mode.setText(_translate("settings_selection", "Dark Mode"))
        self.background_choose.setText(_translate("settings_selection", "Background"))
        self.fonts_color.setText(_translate("settings_selection", "Fonts color"))


class ZoomDeletion(QtWidgets.QDialog):
    """Zoom deletion dialog class"""
    def __init__(self):
        super().__init__()
        self.setWindowIcon(ui.warning_icon)
        self.setModal(True)
        self.setFixedSize(418, 123)
        size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed)
        size_policy.setHorizontalStretch(0)
        size_policy.setVerticalStretch(0)
        size_policy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(size_policy)
        self.setFocusPolicy(QtCore.Qt.FocusPolicy.StrongFocus)
        self.setStyleSheet("QDialog {{\n"
                           "    background-color: rgb{bc};\n"
                           "    border-color: rgb(0, 0, 0);\n"
                           "    border-width: 1;\n"
                           "}}\n"
                           "".format(bc=config.colors["background"]))
        self.both = QtWidgets.QPushButton(parent=self)
        self.both.setGeometry(QtCore.QRect(20, 80, 81, 31))
        self.both.setFont(Fonts.ariel_12_bold)
        self.both.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.both.setMouseTracking(False)
        self.both.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.both.setAutoFillBackground(False)
        self.both.setStyleSheet("QPushButton {{\n"
                                "    border: 2px solid rgb(138, 138, 138);\n"
                                "    border-width: 1px;\n"
                                "    border-radius: 15px;\n"
                                "    padding-left: 2px;\n"
                                "    padding-right: 2px;\n"
                                "    background-color: rgb{bc};\n"
                                "    color: rgb{tc};\n"
                                "}}\n"
                                "QPushButton:pressed {{\n"
                                "    border: 2px solid rgb(138, 138, 138);\n"
                                "    border-width: 2px;\n"
                                "    border-radius: 15px;\n"
                                "    padding-left: 2px;\n"
                                "    padding-right: 2px;\n"
                                "    background-color: rgb{bc};\n"
                                "    border-style: inset;\n"
                                "    color: rgb{tc};\n"
                                "}}\n"
                                "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.both.setObjectName("both")
        self.both.clicked.connect(lambda: self.finished(0))
        self.addon = QtWidgets.QPushButton(parent=self)
        self.addon.setGeometry(QtCore.QRect(120, 80, 81, 31))
        self.addon.setFont(Fonts.ariel_12_bold)
        self.addon.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.addon.setMouseTracking(False)
        self.addon.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.addon.setAutoFillBackground(False)
        self.addon.setStyleSheet("QPushButton {{\n"
                                 "    border: 2px solid rgb(138, 138, 138);\n"
                                 "    border-width: 1px;\n"
                                 "    border-radius: 15px;\n"
                                 "    padding-left: 2px;\n"
                                 "    padding-right: 2px;\n"
                                 "    background-color: rgb{bc};\n"
                                 "    color: rgb{tc};\n"
                                 "}}\n"
                                 "QPushButton:pressed {{\n"
                                 "    border: 2px solid rgb(138, 138, 138);\n"
                                 "    border-width: 2px;\n"
                                 "    border-radius: 15px;\n"
                                 "    padding-left: 2px;\n"
                                 "    padding-right: 2px;\n"
                                 "    background-color: rgb{bc};\n"
                                 "    border-style: inset;\n"
                                 "    color: rgb{tc};\n"
                                 "}}\n"
                                 "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.addon.setObjectName("addon")
        self.addon.clicked.connect(lambda: self.finished(1))
        self.zoom = QtWidgets.QPushButton(parent=self)
        self.zoom.setGeometry(QtCore.QRect(220, 80, 81, 31))
        self.zoom.setFont(Fonts.ariel_12_bold)
        self.zoom.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.zoom.setMouseTracking(False)
        self.zoom.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.zoom.setAutoFillBackground(False)
        self.zoom.setStyleSheet("QPushButton {{\n"
                                "    border: 2px solid rgb(138, 138, 138);\n"
                                "    border-width: 1px;\n"
                                "    border-radius: 15px;\n"
                                "    padding-left: 2px;\n"
                                "    padding-right: 2px;\n"
                                "    background-color: rgb{bc};\n"
                                "    color: rgb{tc};\n"
                                "}}\n"
                                "QPushButton:pressed {{\n"
                                "    border: 2px solid rgb(138, 138, 138);\n"
                                "    border-width: 2px;\n"
                                "    border-radius: 15px;\n"
                                "    padding-left: 2px;\n"
                                "    padding-right: 2px;\n"
                                "    background-color: rgb{bc};\n"
                                "    border-style: inset;\n"
                                "    color: rgb{tc};\n"
                                "}}\n"
                                "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.zoom.setObjectName("zoom")
        self.zoom.clicked.connect(lambda: self.finished(2))
        self.cancel = QtWidgets.QPushButton(parent=self)
        self.cancel.setGeometry(QtCore.QRect(320, 80, 81, 31))
        self.cancel.setFont(Fonts.ariel_12_bold)
        self.cancel.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.cancel.setMouseTracking(False)
        self.cancel.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.cancel.setAutoFillBackground(False)
        self.cancel.setStyleSheet("QPushButton {{\n"
                                  "    border: 2px solid rgb(138, 138, 138);\n"
                                  "    border-width: 1px;\n"
                                  "    border-radius: 15px;\n"
                                  "    padding-left: 2px;\n"
                                  "    padding-right: 2px;\n"
                                  "    background-color: rgb{bc};\n"
                                  "    color: rgb{tc};\n"
                                  "}}\n"
                                  "QPushButton:pressed {{\n"
                                  "    border: 2px solid rgb(138, 138, 138);\n"
                                  "    border-width: 2px;\n"
                                  "    border-radius: 15px;\n"
                                  "    padding-left: 2px;\n"
                                  "    padding-right: 2px;\n"
                                  "    background-color: rgb{bc};\n"
                                  "    border-style: inset;\n"
                                  "    color: rgb{tc};\n"
                                  "}}\n"
                                  "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.cancel.setObjectName("cancel")
        self.cancel.clicked.connect(lambda: self.finished(3))
        self.label = QtWidgets.QLabel(parent=self)
        self.label.setGeometry(QtCore.QRect(10, 10, 401, 61))
        self.label.setFont(Fonts.ariel_12_bold)
        self.label.setStyleSheet("QLabel {{\n"
                                 "    background-color: rgb{bc};\n"
                                 "    border: 0px solid rgb(238, 238, 245);\n"
                                 "    border-width: 0px;\n"
                                 "    color: rgb{tc};\n"
                                 "}}".format(bc=config.colors["background"], tc=config.colors["text_color"]))
        self.label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.label.setObjectName("label")

        self.translate_ui()
        QtCore.QMetaObject.connectSlotsByName(self)
        self.show()
        self.activateWindow()
        self.raise_()

    def translate_ui(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("self", "Zoom Deletion"))
        self.both.setText(_translate("self", "Both"))
        self.addon.setText(_translate("self", "Addon"))
        self.zoom.setText(_translate("self", "Zoom"))
        self.cancel.setText(_translate("self", "Cancel"))
        self.label.setText(_translate("self", "Select an option from the following options:"))

    def finished(self, code: int):
        config.zoom = code
        config.interaction_done = True
        self.close()
        self.deleteLater()

    def done(self, code):
        config.interaction_done = True
        self.close()
        self.deleteLater()


class YesNo(QtWidgets.QDialog):
    """Yes no dialog class"""
    def __init__(self, title: str, question: str):
        super().__init__()
        config.yes_no = (False, False)
        self.setWindowIcon(ui.warning_icon)
        self.setModal(True)
        self.setObjectName("self")
        self.resize(477, 163)
        self.setFixedWidth(477)
        self.setMinimumSize(477, 163)
        self.setWindowTitle(title)
        self.ret = False
        self.setStyleSheet("QDialog {{\n"
                           "    background-color: rgb{bc};\n"
                           "}}".format(bc=config.colors["background"]))
        self.text_edit = QtWidgets.QTextEdit(parent=self)
        self.text_edit.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.text_edit.setReadOnly(True)
        self.text_edit.setFont(Fonts.ariel_11_bold)
        self.text_edit.setStyleSheet("QTextEdit {{\n"
                                     "    background-color: rgb{bc};\n"
                                     "    border: none;"
                                     "    color: rgb{tc};\n"
                                     "    text-align: center;\n"
                                     "}}".format(bc=config.colors["background"], tc=config.colors["text_color"]))
        self.text_edit.insertPlainText(question)

        self.no = QtWidgets.QPushButton("NO", parent=self)
        self.no.setGeometry(QtCore.QRect(270, 120, 101, 31))
        self.no.setFixedSize(101, 31)
        self.no.setStyleSheet("QPushButton {{\n"
                              "    border: 2px solid rgb(138, 138, 138);\n"
                              "    border-width: 1px;\n"
                              "    border-radius: 15px;\n"
                              "    padding-left: 2px;\n"
                              "    padding-right: 2px;\n"
                              "    background-color: rgb{bc};\n"
                              "    color: rgb{tc};\n"
                              "}}\n"
                              "QPushButton:pressed {{\n"
                              "    border: 2px solid rgb(138, 138, 138);\n"
                              "    border-width: 2px;\n"
                              "    border-radius: 15px;\n"
                              "    padding-left: 2px;\n"
                              "    padding-right: 2px;\n"
                              "    background-color: rgb{bc};\n"
                              "    border-style: inset;\n"
                              "    color: rgb{tc};\n"
                              "}}\n"
                              "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.no.setFont(Fonts.ariel_12_bold)
        self.no.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.no.setMouseTracking(False)
        self.no.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.no.setAutoFillBackground(False)
        self.no.clicked.connect(lambda: self.done(False))
        self.yes = QtWidgets.QPushButton("Yes", parent=self)
        self.yes.setGeometry(QtCore.QRect(110, 120, 101, 31))
        self.yes.setFixedSize(101, 31)
        self.yes.setFont(Fonts.ariel_12_bold)
        self.yes.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.yes.setMouseTracking(False)
        self.yes.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.yes.setAutoFillBackground(False)
        self.yes.setStyleSheet("QPushButton {{\n"
                               "    border: 2px solid rgb(138, 138, 138);\n"
                               "    border-width: 1px;\n"
                               "    border-radius: 15px;\n"
                               "    padding-left: 2px;\n"
                               "    padding-right: 2px;\n"
                               "    background-color: rgb{bc};\n"
                               "    color: rgb{tc};\n"
                               "}}\n"
                               "QPushButton:pressed {{\n"
                               "    border: 2px solid rgb(138, 138, 138);\n"
                               "    border-width: 2px;\n"
                               "    border-radius: 15px;\n"
                               "    padding-left: 2px;\n"
                               "    padding-right: 2px;\n"
                               "    background-color: rgb{bc};\n"
                               "    border-style: inset;\n"
                               "    color: rgb{tc};\n"
                               "}}\n"
                               "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.yes.clicked.connect(lambda: self.done(True))
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.no)
        button_layout.addWidget(self.yes)
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.text_edit)
        layout.addLayout(button_layout)
        self.setLayout(layout)

        desktop = app.primaryScreen()
        screen_height = desktop.availableGeometry().height()
        max_height = int(0.9 * screen_height) - 100
        current = self.height()
        height = len(question)
        if "\n" in question:
            amount = question.count("\n")
            while amount and current < max_height:
                current += 15
                amount -= 1
            current += 10
        while height > current and current < max_height:
            current += 14
            height -= 40
        self.setFixedHeight(current - 15)
        self.adjustSize()
        self.translate_ui(title)
        self.show()
        self.activateWindow()
        self.raise_()

    def translate_ui(self, title: str):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("self", title))
        self.no.setText(_translate("self", "NO"))
        self.yes.setText(_translate("self", "Yes"))

    def done(self, return_: bool | int = 0):
        if type(return_) is bool:
            self.ret = return_
        WorkerSignals.yes_no = (self.ret, True)
        self.close()
        self.deleteLater()


class ShowInfo(QtWidgets.QDialog):
    """Shows info to the user via a prompt"""
    def __init__(self, title: str, message: str):
        super().__init__()
        self.setWindowIcon(ui.info_icon)
        self.setModal(True)
        self.setObjectName("self")
        self.resize(477, 163)
        self.setFixedWidth(477)
        self.setMinimumSize(477, 163)
        self.setWindowTitle(title)
        self.setStyleSheet("QDialog {{\n"
                           "    background-color: rgb{bc};\n"
                           "}}".format(bc=config.colors["background"]))
        self.text_edit = QtWidgets.QTextEdit(parent=self)
        self.text_edit.setReadOnly(True)
        self.text_edit.setFont(Fonts.ariel_11_bold)
        self.text_edit.setText(message)
        self.text_edit.setStyleSheet("QTextEdit {{\n"
                                     "    background-color: rgb{bc};\n"
                                     "    border: none;"
                                     "    color: rgb{tc};\n"
                                     "}}".format(bc=config.colors["background"], tc=config.colors["text_color"]))
        self.text_edit.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        self.ok = QtWidgets.QPushButton("OK", parent=self)
        self.ok.setGeometry(QtCore.QRect(270, 120, 101, 31))
        self.ok.setFixedSize(101, 31)
        self.ok.setStyleSheet("QPushButton {{\n"
                              "    border: 2px solid rgb(138, 138, 138);\n"
                              "    border-width: 1px;\n"
                              "    border-radius: 15px;\n"
                              "    padding-left: 2px;\n"
                              "    padding-right: 2px;\n"
                              "    background-color: rgb{bc};\n"
                              "    color: rgb{tc};\n"
                              "}}\n"
                              "QPushButton:pressed {{\n"
                              "    border: 2px solid rgb(138, 138, 138);\n"
                              "    border-width: 2px;\n"
                              "    border-radius: 15px;\n"
                              "    padding-left: 2px;\n"
                              "    padding-right: 2px;\n"
                              "    background-color: rgb{bc};\n"
                              "    border-style: inset;\n"
                              "    color: rgb{tc};\n"
                              "}}\n"
                              "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.ok.setFont(Fonts.ariel_12_bold)
        self.ok.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.ok.setMouseTracking(False)
        self.ok.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.ok.setAutoFillBackground(False)
        self.ok.clicked.connect(self.done)
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ok)
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.text_edit)
        layout.addLayout(button_layout)
        self.setLayout(layout)

        desktop = app.primaryScreen()
        screen_height = desktop.availableGeometry().height()
        max_height = int(0.9 * screen_height) - 100
        current = self.height()
        height = len(message)
        if "\n" in message:
            amount = message.count("\n")
            while amount and current < max_height:
                current += 15
                amount -= 1
            current += 10
        while height > current and current < max_height:
            current += 12
            height -= 40
        self.setFixedHeight(current)
        self.adjustSize()
        self.show()
        self.translate_ui(title)
        self.activateWindow()
        self.raise_()

    def translate_ui(self, title: str):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("self", title))
        self.ok.setText(_translate("self", "OK"))

    def done(self, return_: bool | int = 0):
        self.close()
        self.deleteLater()


class UserDeletion(QtWidgets.QDialog):
    """User deletion dialog class"""
    def __init__(self, users_lst: list):
        super().__init__()
        self.setWindowIcon(ui.warning_icon)
        config.will_delete = []
        self.setModal(True)
        self.setFixedSize(471, 348)
        size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed)
        size_policy.setHorizontalStretch(0)
        size_policy.setVerticalStretch(0)
        size_policy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(size_policy)
        self.setStyleSheet("QWidget {{\n"
                           "background-color: rgb{bc};\n"
                           "    border-color: rgb(0, 0, 0);\n"
                           "    border-width: 1;\n"
                           "}}\n"
                           "".format(bc=config.colors["background"]))
        self.users_lst = users_lst
        self.output_list = list()
        self.options_frame = QtWidgets.QFrame(parent=self)
        self.options_frame.setGeometry(QtCore.QRect(320, -10, 161, 361))
        self.options_frame.setStyleSheet("QFrame {{\n"
                                         "    background-color: rgb{bc};\n"
                                         "    border-color: rgb(0, 0, 0);\n"
                                         "    border-width: 1;\n"
                                         "}}\n"
                                         "".format(bc=config.colors["background"]))
        self.options_frame.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.options_frame.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.options_frame.setObjectName("options_frame")
        self.label = QtWidgets.QLabel(parent=self.options_frame)
        self.label.setGeometry(QtCore.QRect(10, 10, 131, 31))
        self.label.setFont(Fonts.ariel_12_bold)
        self.label.setStyleSheet("QLabel {{\n"
                                 "    background-color: rgb{bc};\n"
                                 "    border: 0px solid rgb(238, 238, 245);\n"
                                 "    border-width: 0px;\n"
                                 "    color: rgb{tc};\n"
                                 "}}\n"
                                 "".format(bc=config.colors["background"], tc=config.colors["text_color"]))
        self.label.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.label.setObjectName("label")
        self.wmi = QtWidgets.QRadioButton(parent=self.options_frame)
        self.wmi.setFont(Fonts.ariel_12_bold)
        self.wmi.setGeometry(QtCore.QRect(10, 50, 89, 20))
        self.wmi.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.wmi.setObjectName("wmi")
        self.wmi.toggled.connect(self.method)
        self.wmi.setStyleSheet("QRadioButton {{"
                               "    background-color: rgb{bc};\n"
                               "    color: rgb{tc};\n"
                               "}}".format(bc=config.colors["background"], tc=config.colors["text_color"]))
        self.threads = QtWidgets.QRadioButton(parent=self.options_frame)
        self.threads.setGeometry(QtCore.QRect(10, 80, 89, 20))
        self.threads.setFont(Fonts.ariel_12_bold)
        self.threads.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.threads.setObjectName("threads")
        self.threads.toggled.connect(self.method)
        self.threads.setStyleSheet("QRadioButton {{"
                                   "    background-color: rgb{bc};\n"
                                   "    color: rgb{tc};\n"
                                   "}}".format(bc=config.colors["background"], tc=config.colors["text_color"]))
        self.select_all = QtWidgets.QPushButton(parent=self.options_frame)
        self.select_all.setGeometry(QtCore.QRect(30, 230, 101, 31))
        self.select_all.setFont(Fonts.ariel_12_bold)
        self.select_all.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.select_all.setMouseTracking(False)
        self.select_all.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.select_all.setAutoFillBackground(False)
        self.select_all.setStyleSheet("QPushButton {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 15px;\n"
                                      "    padding-left: 2px;\n"
                                      "    padding-right: 2px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "QPushButton:pressed {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 2px;\n"
                                      "    border-radius: 15px;\n"
                                      "    padding-left: 2px;\n"
                                      "    padding-right: 2px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    border-style: inset;\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.select_all.setObjectName("select_all")
        self.select_all.clicked.connect(self.select_all_f)
        self.deselect_all = QtWidgets.QPushButton(parent=self.options_frame)
        self.deselect_all.setGeometry(QtCore.QRect(30, 270, 101, 31))
        self.deselect_all.setFont(Fonts.ariel_12_bold)
        self.deselect_all.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.deselect_all.setMouseTracking(False)
        self.deselect_all.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.deselect_all.setAutoFillBackground(False)
        self.deselect_all.setStyleSheet("QPushButton {{\n"
                                        "    border: 2px solid rgb(138, 138, 138);\n"
                                        "    border-width: 1px;\n"
                                        "    border-radius: 15px;\n"
                                        "    padding-left: 2px;\n"
                                        "    padding-right: 2px;\n"
                                        "    background-color: rgb{bc};\n"
                                        "    color: rgb{tc};\n"
                                        "}}\n"
                                        "QPushButton:pressed {{\n"
                                        "    border: 2px solid rgb(138, 138, 138);\n"
                                        "    border-width: 2px;\n"
                                        "    border-radius: 15px;\n"
                                        "    padding-left: 2px;\n"
                                        "    padding-right: 2px;\n"
                                        "    background-color: rgb{bc};\n"
                                        "    border-style: inset;\n"
                                        "    color: rgb{tc};\n"
                                        "}}\n"
                                        "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.deselect_all.setObjectName("deselect_all")
        self.deselect_all.clicked.connect(self.deselect_all_f)
        self.done_btn = QtWidgets.QPushButton(parent=self.options_frame)
        self.done_btn.setGeometry(QtCore.QRect(30, 310, 101, 31))
        self.done_btn.setFont(Fonts.ariel_12_bold)
        self.done_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.done_btn.setMouseTracking(False)
        self.done_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.done_btn.setAutoFillBackground(False)
        self.done_btn.setStyleSheet("QPushButton {{\n"
                                    "    border: 2px solid rgb(138, 138, 138);\n"
                                    "    border-width: 1px;\n"
                                    "    border-radius: 15px;\n"
                                    "    padding-left: 2px;\n"
                                    "    padding-right: 2px;\n"
                                    "    background-color: rgb{bc};\n"
                                    "    color: rgb{tc};\n"
                                    "}}\n"
                                    "QPushButton:pressed {{\n"
                                    "    border: 2px solid rgb(138, 138, 138);\n"
                                    "    border-width: 2px;\n"
                                    "    border-radius: 15px;\n"
                                    "    padding-left: 2px;\n"
                                    "    padding-right: 2px;\n"
                                    "    background-color: rgb{bc};\n"
                                    "    border-style: inset;\n"
                                    "    color: rgb{tc};\n"
                                    "}}\n"
                                    "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.done_btn.setObjectName("done")
        self.done_btn.clicked.connect(lambda: self.done(123))
        self.scrollArea = QtWidgets.QScrollArea(parent=self)
        self.scrollArea.setGeometry(QtCore.QRect(0, 0, 321, 351))
        self.scrollArea.setStyleSheet("""
        QScrollBar:horizontal {
            border: none;
            background: none;
            height: 18px;
            margin: 0px 13px 0 13px;
        }

        QScrollBar::handle:horizontal {
            background: rgb(93, 93, 93);
            min-width: 13px;
            background: rgb(93, 93, 93);
            min-height: 13px;
            border: 2px solid rgb(93, 93, 93);
            border-width: 2px;
            border-radius: 5px;
        }

        QScrollBar::add-line:horizontal {
            background: none;
            width: 13px;
            subcontrol-position: right;
            subcontrol-origin: margin;

        }

        QScrollBar::sub-line:horizontal {
            background: none;
            width: 13px;
            subcontrol-position: top left;
            subcontrol-origin: margin;
            position: absolute;
        }

        QScrollBar:left-arrow:horizontal, QScrollBar::right-arrow:horizontal {
            width: 13px;
            height: 13px;
            background: none;
            image: none;
        }

        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
            background: none;
        }

        QScrollBar:vertical {
            border: none;
            background: none;
            width: 13px;
            margin: 13px 0 18px 0;
        }

        QScrollBar::handle:vertical {
            background: rgb(93, 93, 93);
            min-width: 13px;
            background: rgb(93, 93, 93);
            min-height: 13px;
            border: 2px solid rgb(93, 93, 93);
            border-width: 2px;
            border-radius: 5px;
        }

        QScrollBar::add-line:vertical {
            background: none;
            height: 13px;
            subcontrol-position: bottom;
            subcontrol-origin: margin;
        }

        QScrollBar::sub-line:vertical {
            background: none;
            height: 13px;
            subcontrol-position: top left;
            subcontrol-origin: margin;
            position: absolute;
        }

        QScrollBar:up-arrow:vertical, QScrollBar::down-arrow:vertical {
            width: 13px;
            height: 13px;
            background: none;
            image: none;
        }

        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
            background: none;
        }

    """)
        self.scrollArea.setWidgetResizable(False)
        self.scrollArea.setObjectName("scrollArea")

        self.scroll_content_widget = QtWidgets.QWidget()
        self.scroll_content_layout = QtWidgets.QVBoxLayout(self.scroll_content_widget)

        long_ = 20
        for user in self.users_lst:
            user_select = QtWidgets.QCheckBox(parent=self.scroll_content_widget)
            user_select.setGeometry(QtCore.QRect(10, long_, 201, 20))
            user_select.setFont(Fonts.ariel_11_bold)
            user_select.setObjectName(user)
            user_select.setText(QtCore.QCoreApplication.translate("user_deletion", user))
            user_select.stateChanged.connect(self.add_to_output)
            user_select.setStyleSheet("QCheckBox {{\n"
                                      "   background-color: rgb{bc};\n"
                                      "   color: rgb{tc};\n"
                                      "}}\n"
                                      "QCheckBox::indicator:unchecked{{\n"
                                      "   background-color: grey;\n"
                                      "}}\n"
                                      "QCheckBox::indicator:checked {{\n"
                                      "   background-color: rgb{tc};\n"
                                      "}}".format(bc=config.colors["background"], tc=config.colors["text_color"]))
            long_ += 20

            self.scroll_content_layout.addWidget(user_select)

        self.scrollArea.setWidget(self.scroll_content_widget)

        self.translate_ui(self)
        self.wmi.click()
        QtCore.QMetaObject.connectSlotsByName(self)
        self.show()
        self.activateWindow()
        self.raise_()

    def translate_ui(self, user_deletion_):
        _translate = QtCore.QCoreApplication.translate
        user_deletion_.setWindowTitle(_translate("user_deletion", "User Deletion"))
        self.label.setText(_translate("user_deletion", "Deletion Method"))
        self.wmi.setText(_translate("user_deletion", "WMI"))
        self.threads.setText(_translate("user_deletion", "Threads"))
        self.select_all.setText(_translate("user_deletion", "Select all"))
        self.deselect_all.setText(_translate("user_deletion", "Deselet all"))
        self.done_btn.setText(_translate("user_deletion", "Done"))

    def method(self):
        choice = self.sender()
        if not choice.isChecked():
            return
        elif choice == self.wmi:
            config.lr = 2
        else:
            config.lr = 1
        pass

    def add_to_output(self):
        choice = self.sender()
        if choice.isChecked():
            self.output_list.append(choice.text())
        else:
            self.output_list.remove(choice.text())

    def select_all_f(self):
        for obj in self.scroll_content_widget.children():
            if isinstance(obj, QtWidgets.QCheckBox) and not obj.isChecked():
                obj.toggle()

    def deselect_all_f(self):
        for obj in self.scroll_content_widget.children():
            if isinstance(obj, QtWidgets.QCheckBox) and obj.isChecked():
                obj.toggle()

    def done(self, return_: bool | int = 0):
        if return_ == 123:
            config.will_delete = self.output_list
        config.interaction_done = True
        self.close()
        self.deleteLater()


def zoom_dialog():
    """A function which Initializes ZoomDeletion class"""
    ZoomDeletion().exec()


class GUI(QtWidgets.QMainWindow):
    """The main UI class"""
    def __init__(self):
        super().__init__()
        self.icon = QtGui.QIcon()
        self.icon.addPixmap(QtGui.QPixmap(asset("icon.ico")), QtGui.QIcon.Mode.Normal)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowType.MSWindowsFixedSizeDialogHint)
        self.setObjectName("self")
        self.setFixedSize(715, 798)
        self.setWindowIcon(self.icon)
        self.setToolTipDuration(1)
        self.setDocumentMode(False)
        self.setTabShape(QtWidgets.QTabWidget.TabShape.Rounded)
        self.threadpool = QThreadPool()
        self.threadpool.setMaxThreadCount(20)
        self.clipboard = app.clipboard()
        self.centralwidget = QtWidgets.QWidget(parent=self)
        self.centralwidget.setObjectName("centralwidget")
        self.computer_entry = QtWidgets.QLineEdit(parent=self.centralwidget)
        self.computer_entry.setGeometry(QtCore.QRect(220, 10, 281, 51))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(18)
        font.setBold(True)
        font.setUnderline(False)
        font.setStrikeOut(False)
        font.setKerning(False)
        self.computer_entry.setFont(font)
        self.computer_entry.setText("")
        self.computer_entry.setObjectName("computer_entry")
        self.progress_bar = QtWidgets.QProgressBar(parent=self.centralwidget)
        self.progress_bar.setGeometry(QtCore.QRect(370, 520, 321, 23))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setBold(True)
        self.progress_bar.setFont(font)
        self.progress_bar.setAcceptDrops(False)
        self.progress_bar.setStyleSheet("QProgressBar {\n"
                                        "    border: 2px solid rgb(6, 176, 37);\n"
                                        "    border-radius: 10px;\n"
                                        "    border-width: 3;\n"
                                        "    border-height: 0;\n"
                                        "    text-align: Center;\n"
                                        "    padding: -1;\n"
                                        "    padding-left: -1px;\n"
                                        "    padding-right: -11px;\n"
                                        "    margin: 0.5;\n"
                                        "    background-color: rgb(242, 242, 249);\n"
                                        "}\n"
                                        "QProgressBar::chunk {\n"
                                        "\n"
                                        "}")
        self.progress_bar.setProperty("value", 0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setInvertedAppearance(False)
        self.progress_bar.setTextDirection(QtWidgets.QProgressBar.Direction.TopToBottom)
        self.progress_bar.setObjectName("progress_bar")
        self.progress_bar.setHidden(True)
        self.console = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.console.setGeometry(QtCore.QRect(20, 550, 671, 221))
        self.console.setFont(Fonts.ariel_12_bold)
        self.console.setAcceptDrops(False)
        self.console.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.console.setLineWidth(1)
        self.console.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.console.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.console.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.console.setLineWrapMode(QtWidgets.QTextEdit.LineWrapMode.WidgetWidth)
        self.console.setTextInteractionFlags(
            QtCore.Qt.TextInteractionFlag.TextSelectableByKeyboard | QtCore.Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.console.setOpenLinks(False)
        self.console.setObjectName("console")
        self.pc_display = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.pc_display.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.pc_display.setGeometry(QtCore.QRect(20, 130, 241, 31))
        self.pc_display.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.pc_display.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.pc_display.setObjectName("pc_display")
        self.pc_status = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.pc_status.setGeometry(QtCore.QRect(20, 170, 331, 31))
        self.pc_status.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.pc_status.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.pc_status.setObjectName("pc_status")
        self.pc_status.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.user_display = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.user_display.setGeometry(QtCore.QRect(20, 210, 331, 31))
        self.user_display.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.user_display.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.user_display.setObjectName("user_display")
        self.user_display.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.c_space_display = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.c_space_display.setGeometry(QtCore.QRect(20, 290, 331, 31))
        self.c_space_display.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.c_space_display.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.c_space_display.setObjectName("c_space_display")
        self.c_space_display.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.uptime_display = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.uptime_display.setGeometry(QtCore.QRect(20, 250, 331, 31))
        self.uptime_display.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.uptime_display.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.uptime_display.setObjectName("uptime_display")
        self.uptime_display.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.icon_light = QtGui.QIcon()
        self.icon_dark = QtGui.QIcon()
        self.icon_dark.addPixmap(QtGui.QPixmap(asset("settings_dark.ico")),
                                 QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        self.icon_light.addPixmap(QtGui.QPixmap(asset("settings_light.ico")),
                                  QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        self.settings_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.settings_btn.setGeometry(QtCore.QRect(10, 10, 31, 31))
        self.settings_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.settings_btn.setText("")
        self.settings_btn.setIcon(self.icon_dark if config.colors["dark"] else self.icon_light)
        self.settings_btn.setIconSize(QtCore.QSize(30, 30))
        self.settings_btn.setDefault(False)
        self.settings_btn.setFlat(True)
        self.settings_btn.setObjectName("settings_btn")
        self.settings_btn.clicked.connect(settings_select)
        self.ie_display = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.ie_display.setGeometry(QtCore.QRect(20, 410, 331, 31))
        self.ie_display.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.ie_display.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.ie_display.setObjectName("ie_display")
        self.ie_display.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.ram_display = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.ram_display.setGeometry(QtCore.QRect(20, 370, 331, 31))
        self.ram_display.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.ram_display.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.ram_display.setObjectName("ram_display")
        self.ram_display.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.cpt_status = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.cpt_status.setGeometry(QtCore.QRect(20, 450, 331, 31))
        self.cpt_status.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.cpt_status.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.cpt_status.setObjectName("cpt_status")
        self.cpt_status.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.d_space_display = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.d_space_display.setGeometry(QtCore.QRect(20, 330, 331, 31))
        self.d_space_display.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.d_space_display.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.d_space_display.setObjectName("d_space_display")
        self.d_space_display.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.user_status = QtWidgets.QTextBrowser(parent=self.centralwidget)
        self.user_status.setGeometry(QtCore.QRect(20, 490, 331, 31))
        self.user_status.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.user_status.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.user_status.setObjectName("user_status")
        self.user_status.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.NoTextInteraction)
        self.copy_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.copy_btn.setGeometry(QtCore.QRect(270, 130, 81, 31))
        self.copy_btn.setFont(Fonts.ariel_12_bold)
        self.copy_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.copy_btn.setMouseTracking(True)
        self.copy_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.copy_btn.setAutoFillBackground(False)
        self.copy_btn.setObjectName("copy_btn")
        self.submit_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.submit_btn.setGeometry(QtCore.QRect(310, 70, 111, 31))
        self.submit_btn.setFont(Fonts.ariel_12_bold)
        self.submit_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.submit_btn.setMouseTracking(True)
        self.submit_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.submit_btn.setAutoFillBackground(False)
        self.submit_btn.setObjectName("submit_btn")
        self.label_for_pg = QtWidgets.QLabel(parent=self.centralwidget)
        self.label_for_pg.setGeometry(QtCore.QRect(360, 490, 331, 31))
        self.label_for_pg.setFont(Fonts.ariel_12_bold)
        self.label_for_pg.setText("")
        self.label_for_pg.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_for_pg.setObjectName("label_for_pg")
        self.reset_spool_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.reset_spool_btn.setGeometry(QtCore.QRect(550, 130, 141, 41))
        self.reset_spool_btn.setFont(Fonts.ariel_12_bold)
        self.reset_spool_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.reset_spool_btn.setMouseTracking(True)
        self.reset_spool_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.reset_spool_btn.setAutoFillBackground(False)
        self.reset_spool_btn.setObjectName("reset_spool_btn")
        self.close_outlook_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.close_outlook_btn.setGeometry(QtCore.QRect(550, 180, 141, 41))
        self.close_outlook_btn.setFont(Fonts.ariel_12_bold)
        self.close_outlook_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.close_outlook_btn.setMouseTracking(True)
        self.close_outlook_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.close_outlook_btn.setAutoFillBackground(False)
        self.close_outlook_btn.setObjectName("close_outlook_btn")
        self.del_ost_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.del_ost_btn.setGeometry(QtCore.QRect(550, 230, 141, 41))
        self.del_ost_btn.setFont(Fonts.ariel_12_bold)
        self.del_ost_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.del_ost_btn.setMouseTracking(True)
        self.del_ost_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.del_ost_btn.setAutoFillBackground(False)
        self.del_ost_btn.setObjectName("del_ost_btn")
        self.del_teams_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.del_teams_btn.setGeometry(QtCore.QRect(550, 280, 141, 41))
        self.del_teams_btn.setFont(Fonts.ariel_12_bold)
        self.del_teams_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.del_teams_btn.setMouseTracking(True)
        self.del_teams_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.del_teams_btn.setAutoFillBackground(False)
        self.del_teams_btn.setObjectName("del_teams_btn")
        self.sample_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.sample_btn.setGeometry(QtCore.QRect(550, 330, 141, 41))
        self.sample_btn.setFont(Fonts.ariel_12_bold)
        self.sample_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.sample_btn.setMouseTracking(True)
        self.sample_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.sample_btn.setAutoFillBackground(False)
        self.sample_btn.setObjectName("sample")
        self.clear_space_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.clear_space_btn.setGeometry(QtCore.QRect(390, 130, 141, 41))
        self.clear_space_btn.setFont(Fonts.ariel_12_bold)
        self.clear_space_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.clear_space_btn.setMouseTracking(True)
        self.clear_space_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.clear_space_btn.setAutoFillBackground(False)
        self.clear_space_btn.setObjectName("clear_space_btn")
        self.del_users_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.del_users_btn.setGeometry(QtCore.QRect(390, 180, 141, 41))
        self.del_users_btn.setFont(Fonts.ariel_12_bold)
        self.del_users_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.del_users_btn.setMouseTracking(True)
        self.del_users_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.del_users_btn.setAutoFillBackground(False)
        self.del_users_btn.setObjectName("del_users_btn")
        self.printers_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.printers_btn.setGeometry(QtCore.QRect(390, 230, 141, 41))
        self.printers_btn.setFont(Fonts.ariel_12_bold)
        self.printers_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.printers_btn.setMouseTracking(True)
        self.printers_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.printers_btn.setAutoFillBackground(False)
        self.printers_btn.setObjectName("printers_btn")
        self.del_zoom_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.del_zoom_btn.setGeometry(QtCore.QRect(390, 280, 141, 41))
        self.del_zoom_btn.setFont(Fonts.ariel_12_bold)
        self.del_zoom_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.del_zoom_btn.setMouseTracking(True)
        self.del_zoom_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.del_zoom_btn.setAutoFillBackground(False)
        self.del_zoom_btn.setObjectName("del_zoom_btn")
        self.export_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.export_btn.setGeometry(QtCore.QRect(390, 330, 141, 41))
        self.export_btn.setFont(Fonts.ariel_12_bold)
        self.export_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.export_btn.setMouseTracking(True)
        self.export_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.export_btn.setAutoFillBackground(False)
        self.export_btn.setObjectName("export_btn")
        self.fix_cpt_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.fix_cpt_btn.setGeometry(QtCore.QRect(390, 380, 141, 41))
        self.fix_cpt_btn.setFont(Fonts.ariel_12_bold)
        self.fix_cpt_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.fix_cpt_btn.setMouseTracking(True)
        self.fix_cpt_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.fix_cpt_btn.setAutoFillBackground(False)
        self.fix_cpt_btn.setObjectName("fix_cpt_btn")
        self.fix_ie_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.fix_ie_btn.setGeometry(QtCore.QRect(550, 380, 141, 41))
        self.fix_ie_btn.setFont(Fonts.ariel_12_bold)
        self.fix_ie_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.fix_ie_btn.setMouseTracking(True)
        self.fix_ie_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.fix_ie_btn.setAutoFillBackground(False)
        self.fix_ie_btn.setAutoDefault(False)
        self.fix_ie_btn.setDefault(False)
        self.fix_ie_btn.setObjectName("fix_ie_btn")
        self.restart_pc_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.restart_pc_btn.setGeometry(QtCore.QRect(390, 430, 141, 41))
        self.restart_pc_btn.setFont(Fonts.ariel_12_bold)
        self.restart_pc_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.restart_pc_btn.setMouseTracking(True)
        self.restart_pc_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.restart_pc_btn.setAutoFillBackground(False)
        self.restart_pc_btn.setObjectName("restart_pc")
        self.fix_3_lang_btn = QtWidgets.QPushButton(parent=self.centralwidget)
        self.fix_3_lang_btn.setGeometry(QtCore.QRect(550, 430, 141, 41))
        self.fix_3_lang_btn.setFont(Fonts.ariel_12_bold)
        self.fix_3_lang_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.fix_3_lang_btn.setMouseTracking(True)
        self.fix_3_lang_btn.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.fix_3_lang_btn.setAutoFillBackground(False)
        self.fix_3_lang_btn.setObjectName("del_teams_btn_2")
        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(parent=self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)
        self.translate_ui()
        QtCore.QMetaObject.connectSlotsByName(self)
        self.submit_btn.clicked.connect(lambda: self.runit(on_submit))
        self.restart_pc_btn.clicked.connect(lambda: self.runit(restart))
        self.close_outlook_btn.clicked.connect(lambda: self.runit(close_outlook))
        self.printers_btn.clicked.connect(lambda: self.runit(get_printers_func))
        self.del_teams_btn.clicked.connect(lambda: self.runit(rem_teams))
        self.clear_space_btn.clicked.connect(lambda: self.runit(clear_space_func))
        self.reset_spool_btn.clicked.connect(lambda: self.runit(reset_spooler))
        self.fix_cpt_btn.clicked.connect(lambda: self.runit(fix_cpt_func))
        self.fix_ie_btn.clicked.connect(lambda: self.runit(fix_ie_func))
        self.export_btn.clicked.connect(lambda: self.runit(export))
        self.copy_btn.clicked.connect(copy_pc)
        self.computer_entry.returnPressed.connect(lambda: self.runit(on_submit))
        self.del_users_btn.clicked.connect(lambda: print("Fuck off, this is not supported any more") if
                                           config.user.lower() != "c1490933" else self.runit(del_users))
        self.del_ost_btn.clicked.connect(lambda: self.runit(del_ost))
        self.del_zoom_btn.clicked.connect(lambda: self.runit(delete_zoom))
        self.fix_3_lang_btn.clicked.connect(lambda: self.runit(fix_3_languages))
        self.sample_btn.clicked.connect(lambda: self.runit(sample_function))
        self.set_colors()
        self.warning_icon = QtGui.QIcon()
        self.warning_icon.addPixmap(QtGui.QPixmap(asset("warning.ico")), QtGui.QIcon.Mode.Normal)
        self.info_icon = QtGui.QIcon()
        self.info_icon.addPixmap(QtGui.QPixmap(asset("info.ico")), QtGui.QIcon.Mode.Normal)
        self.show()

    def set_colors(self):
        """Sets the colors and style for the elements in the ui"""
        self.ie_display.setStyleSheet("QTextBrowser {{\n"
                                      "    border: 2px solid rgb(216, 215, 215);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 8px;\n"
                                      "    padding-left: 10px;\n"
                                      "    padding-right: 10px;\n"
                                      "    padding: 1px;\n"
                                      "    qproperty-alignment: AlignLeft;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.ram_display.setStyleSheet("QTextBrowser {{\n"
                                       "    border: 2px solid rgb(216, 215, 215);\n"
                                       "    border-width: 1px;\n"
                                       "    border-radius: 8px;\n"
                                       "    padding-left: 10px;\n"
                                       "    padding-right: 10px;\n"
                                       "    padding: 1px;\n"
                                       "    qproperty-alignment: AlignLeft;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    color: rgb{tc};\n"
                                       "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.cpt_status.setStyleSheet("QTextBrowser {{\n"
                                      "    border: 2px solid rgb(216, 215, 215);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 8px;\n"
                                      "    padding-left: 10px;\n"
                                      "    padding-right: 10px;\n"
                                      "    padding: 1px;\n"
                                      "    qproperty-alignment: AlignLeft;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.cpt_status.setStyleSheet("QTextBrowser {{\n"
                                      "    border: 2px solid rgb(216, 215, 215);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 8px;\n"
                                      "    padding-left: 10px;\n"
                                      "    padding-right: 10px;\n"
                                      "    padding: 1px;\n"
                                      "    qproperty-alignment: AlignLeft;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.d_space_display.setStyleSheet("QTextBrowser {{\n"
                                           "    border: 2px solid rgb(216, 215, 215);\n"
                                           "    border-width: 1px;\n"
                                           "    border-radius: 8px;\n"
                                           "    padding-left: 10px;\n"
                                           "    padding-right: 10px;\n"
                                           "    padding: 1px;\n"
                                           "    qproperty-alignment: AlignLeft;\n"
                                           "    background-color: rgb{bc};\n"
                                           "    color: rgb{tc};\n"
                                           "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.user_status.setStyleSheet("QTextBrowser {{\n"
                                       "    border: 2px solid rgb(216, 215, 215);\n"
                                       "    border-width: 1px;\n"
                                       "    border-radius: 8px;\n"
                                       "    padding-left: 10px;\n"
                                       "    padding-right: 10px;\n"
                                       "    padding: 1px;\n"
                                       "    qproperty-alignment: AlignLeft;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    color: rgb{tc};\n"
                                       "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.copy_btn.setStyleSheet("QPushButton {{\n"
                                    "    border: 2px solid rgb(138, 138, 138);\n"
                                    "    border-width: 1px;\n"
                                    "    border-radius: 15px;\n"
                                    "    padding-left: 2px;\n"
                                    "    padding-right: 2px;\n"
                                    "    background-color: rgb{bc};\n"
                                    "    color: rgb{tc};\n"
                                    "}}\n"
                                    "QPushButton:pressed {{\n"
                                    "    border: 2px solid rgb(138, 138, 138);\n"
                                    "    border-width: 2px;\n"
                                    "    border-radius: 15px;\n"
                                    "    padding-left: 2px;\n"
                                    "    padding-right: 2px;\n"
                                    "    background-color: rgb{bc};\n"
                                    "    border-style: inset;\n"
                                    "    color: rgb{tc};\n"
                                    "}}\n"
                                    "QPushButton:disabled {{\n"
                                    "    color:grey;\n"
                                    "}}\n"
                                    "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.submit_btn.setStyleSheet("QPushButton {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 15px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "QPushButton:pressed {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 2px;\n"
                                      "    border-radius: 15px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    border-style: inset;\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "QPushButton:disabled {{\n"
                                      "    color:grey;\n"
                                      "}}\n"
                                      "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.label_for_pg.setStyleSheet("QLabel {{\n"
                                        "    background-color: rgb{bc};\n"
                                        "    border: 0px solid rgb(238, 238, 245);\n"
                                        "    border-width: 0px;\n"
                                        "    color: rgb{tc};\n"
                                        "}}".format(bc=config.colors["background"], tc=config.colors["text_color"]))
        self.reset_spool_btn.setStyleSheet("QPushButton {{\n"
                                           "    border: 2px solid rgb(138, 138, 138);\n"
                                           "    border-width: 1px;\n"
                                           "    border-radius: 20px;\n"
                                           "    padding-left: 12px;\n"
                                           "    padding-right: 12px;\n"
                                           "    background-color: rgb{bc};\n"
                                           "    color: rgb{tc};\n"
                                           "}}\n"
                                           "QPushButton:pressed {{\n"
                                           "    border: 2px solid rgb(138, 138, 138);\n"
                                           "    border-width: 3px;\n"
                                           "    border-radius: 20px;\n"
                                           "    padding-left: 12px;\n"
                                           "    padding-right: 12px;\n"
                                           "    background-color: rgb{bc};\n"
                                           "    border-style: inset;\n"
                                           "    color: rgb{tc};\n"
                                           "}}\n"
                                           "QPushButton:disabled {{\n"
                                           "    color:grey;\n"
                                           "}}\n"
                                           "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.close_outlook_btn.setStyleSheet("QPushButton {{\n"
                                             "    border: 2px solid rgb(138, 138, 138);\n"
                                             "    border-width: 1px;\n"
                                             "    border-radius: 20px;\n"
                                             "    padding-left: 12px;\n"
                                             "    padding-right: 12px;\n"
                                             "    background-color: rgb{bc};\n"
                                             "    color: rgb{tc};\n"
                                             "}}\n"
                                             "QPushButton:pressed {{\n"
                                             "    border: 2px solid rgb(138, 138, 138);\n"
                                             "    border-width: 3px;\n"
                                             "    border-radius: 20px;\n"
                                             "    padding-left: 12px;\n"
                                             "    padding-right: 12px;\n"
                                             "    background-color: rgb{bc};\n"
                                             "    border-style: inset;\n"
                                             "    color: rgb{tc};\n"
                                             "}}\n"
                                             "QPushButton:disabled {{\n"
                                             "    color:grey;\n"
                                             "}}\n"
                                             "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.del_ost_btn.setStyleSheet("QPushButton {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 1px;\n"
                                       "    border-radius: 20px;\n"
                                       "    padding-left: 20px;\n"
                                       "    padding-right: 20px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:pressed {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 3px;\n"
                                       "    border-radius: 20px;\n"
                                       "    padding-left: 20px;\n"
                                       "    padding-right: 20px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    border-style: inset;\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:disabled {{\n"
                                       "    color:grey;\n"
                                       "}}\n"
                                       "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.del_teams_btn.setStyleSheet("QPushButton {{\n"
                                         "    border: 2px solid rgb(138, 138, 138);\n"
                                         "    border-width: 1px;\n"
                                         "    border-radius: 20px;\n"
                                         "    padding-left: 17px;\n"
                                         "    padding-right: 17px;\n"
                                         "    background-color: rgb{bc};\n"
                                         "    color: rgb{tc};\n"
                                         "}}\n"
                                         "QPushButton:pressed {{\n"
                                         "   border: 2px solid rgb(138, 138, 138);\n"
                                         "    border-width: 3px;\n"
                                         "    border-radius: 20px;\n"
                                         "    padding-left: 17px;\n"
                                         "    padding-right: 17px;\n"
                                         "    background-color: rgb{bc};\n"
                                         "    border-style: inset;\n"
                                         "    color: rgb{tc};\n"
                                         "}}\n"
                                         "QPushButton:disabled {{\n"
                                         "    color:grey;\n"
                                         "}}\n"
                                         "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.sample_btn.setStyleSheet("\n"
                                       "QPushButton {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 1px;\n"
                                       "    border-radius: 20px;\n"
                                       "    padding-left: 11px;\n"
                                       "    padding-right: 11px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:pressed {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 3px;\n"
                                       "    border-radius: 20px;\n"
                                       "    padding-left: 11px;\n"
                                       "    padding-right: 11px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    border-style: inset;\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:disabled {{\n"
                                       "    color:grey;\n"
                                       "}}\n"
                                       "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.clear_space_btn.setStyleSheet("QPushButton {{\n"
                                           "    border: 2px solid rgb(138, 138, 138);\n"
                                           "    border-width: 1px;\n"
                                           "    border-radius: 20px;\n"
                                           "    padding-left: 20px;\n"
                                           "    padding-right: 20px;\n"
                                           "    background-color: rgb{bc};\n"
                                           "    color: rgb{tc};\n"
                                           "}}\n"
                                           "QPushButton:pressed {{\n"
                                           "    border: 2px solid rgb(138, 138, 138);\n"
                                           "    border-width: 3px;\n"
                                           "    border-radius: 20px;\n"
                                           "    padding-left: 20px;\n"
                                           "    padding-right: 20px;\n"
                                           "    background-color: rgb{bc};\n"
                                           "    border-style: inset;\n"
                                           "    color: rgb{tc};\n"
                                           "}}\n"
                                           "QPushButton:disabled {{\n"
                                           "    color:grey;\n"
                                           "}}\n"
                                           "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.del_users_btn.setStyleSheet("QPushButton {{\n"
                                         "    border: 2px solid rgb(138, 138, 138);\n"
                                         "    border-width: 1px;\n"
                                         "    border-radius: 20px;\n"
                                         "    padding-left: 20px;\n"
                                         "    padding-right: 20px;\n"
                                         "    background-color: rgb{bc};\n"
                                         "    color: rgb{tc};\n"
                                         "}}\n"
                                         "QPushButton:pressed {{\n"
                                         "   border: 2px solid rgb(138, 138, 138);\n"
                                         "    border-width: 3px;\n"
                                         "    border-radius: 20px;\n"
                                         "    padding-left: 20px;\n"
                                         "    padding-right: 20px;\n"
                                         "    background-color: rgb{bc};\n"
                                         "    border-style: inset;\n"
                                         "    color: rgb{tc};\n"
                                         "}}\n"
                                         "QPushButton:disabled {{\n"
                                         "    color:grey;\n"
                                         "}}\n"
                                         "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.printers_btn.setStyleSheet("QPushButton {{\n"
                                        "    border: 2px solid rgb(138, 138, 138);\n"
                                        "    border-width: 1px;\n"
                                        "    border-radius: 20px;\n"
                                        "    padding-left: 20px;\n"
                                        "    padding-right: 20px;\n"
                                        "    background-color: rgb{bc};\n"
                                        "    color: rgb{tc};\n"
                                        "}}\n"
                                        "QPushButton:pressed {{\n"
                                        "    border: 2px solid rgb(138, 138, 138);\n"
                                        "    border-width: 3px;\n"
                                        "    border-radius: 20px;\n"
                                        "    padding-left: 20px;\n"
                                        "    padding-right: 20px;\n"
                                        "    background-color: rgb{bc};\n"
                                        "    border-style: inset;\n"
                                        "    color: rgb{tc};\n"
                                        "}}\n"
                                        "QPushButton:disabled {{\n"
                                        "    color:grey;\n"
                                        "}}\n"
                                        "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.del_zoom_btn.setStyleSheet("QPushButton {{\n"
                                        "    border: 2px solid rgb(138, 138, 138);\n"
                                        "    border-width: 1px;\n"
                                        "    border-radius: 20px;\n"
                                        "    padding-left: 20px;\n"
                                        "    padding-right: 20px;\n"
                                        "    background-color: rgb{bc};\n"
                                        "    color: rgb{tc};\n"
                                        "}}\n"
                                        "QPushButton:pressed {{\n"
                                        "    border: 2px solid rgb(138, 138, 138);\n"
                                        "    border-width: 3px;\n"
                                        "    border-radius: 20px;\n"
                                        "    padding-left: 20px;\n"
                                        "    padding-right: 20px;\n"
                                        "    background-color: rgb{bc};\n"
                                        "    border-style: inset;\n"
                                        "    color: rgb{tc};\n"
                                        "}}\n"
                                        "QPushButton:disabled {{\n"
                                        "    color:grey;\n"
                                        "}}\n"
                                        "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.export_btn.setStyleSheet("QPushButton {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 20px;\n"
                                      "    padding-left: 20px;\n"
                                      "    padding-right: 20px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "QPushButton:pressed {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 3px;\n"
                                      "    border-radius: 20px;\n"
                                      "    padding-left: 20px;\n"
                                      "    padding-right: 20px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    border-style: inset;\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "QPushButton:disabled {{\n"
                                      "    color:grey;\n"
                                      "}}\n"
                                      "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.fix_cpt_btn.setStyleSheet("QPushButton {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 1px;\n"
                                       "    border-radius: 20px;\n"
                                       "    padding-left: 20px;\n"
                                       "    padding-right: 20px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:pressed {{\n"
                                       "    border: 2px solid rgb(138, 138, 138);\n"
                                       "    border-width: 3px;\n"
                                       "    border-radius: 20px;\n"
                                       "    padding-left: 20px;\n"
                                       "    padding-right: 20px;\n"
                                       "    background-color: rgb{bc};\n"
                                       "    border-style: inset;\n"
                                       "    color: rgb{tc};\n"
                                       "}}\n"
                                       "QPushButton:disabled {{\n"
                                       "    color:grey;\n"
                                       "}}\n"
                                       "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.fix_ie_btn.setStyleSheet("QPushButton {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 20px;\n"
                                      "    padding-left: 20px;\n"
                                      "    padding-right: 20px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "QPushButton:pressed {{\n"
                                      "    border: 2px solid rgb(138, 138, 138);\n"
                                      "    border-width: 3px;\n"
                                      "    border-radius: 20px;\n"
                                      "    padding-left: 20px;\n"
                                      "    padding-right: 20px;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    border-style: inset;\n"
                                      "    color: rgb{tc};\n"
                                      "}}\n"
                                      "QPushButton:disabled {{\n"
                                      "    color:grey;\n"
                                      "}}\n"
                                      "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.fix_3_lang_btn.setStyleSheet("QPushButton {{\n"
                                          "    border: 2px solid rgb(138, 138, 138);\n"
                                          "    border-width: 1px;\n"
                                          "    border-radius: 20px;\n"
                                          "    padding-left: 10px;\n"
                                          "    padding-right: 10px;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    color: rgb{tc};\n"
                                          "}}\n"
                                          "QPushButton:pressed {{\n"
                                          "    border: 2px solid rgb(138, 138, 138);\n"
                                          "    border-width: 3px;\n"
                                          "    border-radius: 20px;\n"
                                          "    padding-left: 11px;\n"
                                          "    padding-right: 11px;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    border-style: inset;\n"
                                          "    color: rgb{tc};\n"
                                          "}}\n"
                                          "QPushButton:disabled {{\n"
                                          "    color:grey;\n"
                                          "}}\n"
                                          "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.restart_pc_btn.setStyleSheet("QPushButton {{\n"
                                          "    border: 2px solid rgb(138, 138, 138);\n"
                                          "    border-width: 1px;\n"
                                          "    border-radius: 20px;\n"
                                          "    padding-left: 17px;\n"
                                          "    padding-right: 17px;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    color: rgb{tc};\n"
                                          "}}\n"
                                          "QPushButton:pressed {{\n"
                                          "   border: 2px solid rgb(138, 138, 138);\n"
                                          "    border-width: 3px;\n"
                                          "    border-radius: 20px;\n"
                                          "    padding-left: 17px;\n"
                                          "    padding-right: 17px;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    border-style: inset;\n"
                                          "    color: rgb{tc};\n"
                                          "}}\n"
                                          "QPushButton:disabled {{\n"
                                          "    color:grey;\n"
                                          "}}\n"
                                          "".format(bc=config.colors["buttons"], tc=config.colors["text_color"]))
        self.setStyleSheet("QWidget {{\n"
                           "    background-color: rgb{bc};\n"
                           "    border-color: rgb(0, 0, 0);\n"
                           "    border-width: 1;\n"
                           "}}\n"
                           "".format(bc=config.colors["background"]))
        self.computer_entry.setStyleSheet("QLineEdit {{\n"
                                          "    border: 2px solid rgb(138, 138, 138);\n"
                                          "    border-width: 1px;\n"
                                          "    border-radius: 25px;\n"
                                          "    padding-left: 10px;\n"
                                          "    padding-right: 10px;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    qproperty-alignment: AlignCenter;\n"
                                          "    color: rgb{tc};\n"
                                          "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.console.setStyleSheet("QTextBrowser {{\n"
                                   "    border: 2px solid rgb(216, 215, 215);\n"
                                   "    border-width: 1px;\n"
                                   "    border-radius: 29px;\n"
                                   "    padding-left: 10px;\n"
                                   "    padding-right: 10px;\n"
                                   "    padding: 10px;\n"
                                   "    qproperty-alignment: AlignLeft;\n"
                                   "    background-color: rgb{bc};\n"
                                   "    color: rgb{tc};\n"
                                   "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.pc_display.setStyleSheet("QTextBrowser {{\n"
                                      "    border: 2px solid rgb(216, 215, 215);\n"
                                      "    border-width: 1px;\n"
                                      "    border-radius: 8px;\n"
                                      "    padding-left: 10px;\n"
                                      "    padding-right: 10px;\n"
                                      "    padding: 1px;\n"
                                      "    qproperty-alignment: AlignLeft;\n"
                                      "    background-color: rgb{bc};\n"
                                      "    color: rgb{tc};\n"
                                      "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.pc_status.setStyleSheet("QTextBrowser {{\n"
                                     "    border: 2px solid rgb(216, 215, 215);\n"
                                     "    border-width: 1px;;\n"
                                     "    border-radius: 8px;\n"
                                     "    padding-left: 10px;\n"
                                     "    padding-right: 10px;\n"
                                     "    padding: 1px;\n"
                                     "    qproperty-alignment: AlignLeft;\n"
                                     "    background-color: rgb{bc};\n"
                                     "    color: rgb{tc};\n"
                                     "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.user_display.setStyleSheet("QTextBrowser {{\n"
                                        "    border: 2px solid rgb(216, 215, 215);\n"
                                        "    border-width: 1px;\n"
                                        "    border-radius: 8px;\n"
                                        "    padding-left: 10px;\n"
                                        "    padding-right: 10px;\n"
                                        "    padding: 1px;\n"
                                        "    qproperty-alignment: AlignLeft;\n"
                                        "    background-color: rgb{bc};\n"
                                        "    color: rgb{tc};\n"
                                        "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.c_space_display.setStyleSheet("QTextBrowser {{\n"
                                           "    border: 2px solid rgb(216, 215, 215);\n"
                                           "    border-width: 1px;\n"
                                           "    border-radius: 8px;\n"
                                           "    padding-left: 10px;\n"
                                           "    padding-right: 10px;\n"
                                           "    padding: 1px;\n"
                                           "    qproperty-alignment: AlignLeft;\n"
                                           "    background-color: rgb{bc};\n"
                                           "    color: rgb{tc};\n"
                                           "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.uptime_display.setStyleSheet("QTextBrowser {{\n"
                                          "    border: 2px solid rgb(216, 215, 215);\n"
                                          "    border-width: 1px;\n"
                                          "    border-radius: 8px;\n"
                                          "    padding-left: 10px;\n"
                                          "    padding-right: 10px;\n"
                                          "    padding: 1px;\n"
                                          "    qproperty-alignment: AlignLeft;\n"
                                          "    background-color: rgb{bc};\n"
                                          "    color: rgb{tc};\n"
                                          "}}".format(bc=config.colors["text_area"], tc=config.colors["text_color"]))
        self.settings_btn.setStyleSheet("QPushButton {{\n"
                                        "    background-color: rgb{bc};\n"
                                        "}}\n"
                                        "QPushButton:pressed {{\n"
                                        "    background-color: rgb{bc};\n"
                                        "    border-style: inset;\n"
                                        "}}".format(bc=config.colors["background"]))

    def translate_ui(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("self", config.title))
        self.computer_entry.setPlaceholderText(_translate("self", "Computer or User"))
        self.reset_spool_btn.setText(_translate("self", "Reset Spooler"))
        self.close_outlook_btn.setText(_translate("self", "Close Outlook"))
        self.del_ost_btn.setText(_translate("self", "Delete OST"))
        self.del_teams_btn.setText(_translate("self", "Delete Teams"))
        self.sample_btn.setText(_translate("self", "sample"))
        self.clear_space_btn.setText(_translate("self", "Clear space"))
        self.del_users_btn.setText(_translate("self", "Delete Users"))
        self.printers_btn.setText(_translate("self", "Get Printers"))
        self.del_zoom_btn.setText(_translate("self", "Delete Zoom"))
        self.export_btn.setText(_translate("self", "Export"))
        self.fix_cpt_btn.setText(_translate("self", "Fix CPT"))
        self.fix_ie_btn.setText(_translate("self", "Fix IE"))
        self.copy_btn.setText(_translate("self", "Copy"))
        self.submit_btn.setText(_translate("self", "Submit"))
        self.restart_pc_btn.setText(_translate("self", "Restart PC"))
        self.fix_3_lang_btn.setText(_translate("self", "3 languages"))

    def call_pg(self, title: str = "") -> None:
        """Calls the progressbar"""
        self.progress_bar.setProperty("value", 0)
        self.progress_bar.setHidden(False)
        self.label_for_pg.setText(title)
        refresh()

    def clean_pg(self, end: str) -> None:
        """Cleans after the progressbar"""
        self.progress_bar.setHidden(True)
        self.progress_bar.setProperty("value", 0)
        self.label_for_pg.clear()
        if end:
            print(end)
        refresh()

    def update_pg(self, val) -> None:
        """Updates the progressbar"""
        self.progress_bar.setProperty("value", val)

    def runit(self, func) -> None:
        """Runs a function via passing it to a worker"""
        self.computer_entry.returnPressed.disconnect()
        disable_1(ui.settings_btn)
        disable(disable_submit=True)
        if func != on_submit:
            if not wmi_connectable() or not reg_connect():
                enable_1(ui.submit_btn)
                enable_1(ui.settings_btn)
                if config.current_computer:
                    enable_1(ui.copy_btn)
                self.computer_entry.returnPressed.connect(lambda: self.runit(on_submit))
                print_error(ui.console, "Couldn't connect to wmi or registry")
                return
        worker = Worker(func)
        worker.signals.call_pb.connect(self.call_pg)
        worker.signals.progress.connect(self.update_pg)
        worker.signals.clean_pb.connect(self.clean_pg)
        worker.signals.finished.connect(self.on_done)
        worker.signals.progress.connect(self.update_pg)
        worker.signals.print_.connect(print)
        worker.signals.print_error_.connect(print_error)
        worker.signals.print_success_.connect(print_success)
        worker.signals.ask_yes_no.connect(YesNo)
        worker.signals.show_info_.connect(show_info)
        worker.signals.copy.connect(copy_clip)
        worker.signals.update.connect(update)
        worker.signals.update_error.connect(update_error)
        worker.signals.update_success.connect(update_success)
        worker.signals.enable_1.connect(enable_1)
        worker.signals.disable_1.connect(disable_1)
        worker.signals.clear_all.connect(clear_all)
        worker.signals.del_users.connect(del_users_selection_actual)
        worker.signals.run_without_waiting.connect(self.run_without_waiting)
        worker.signals.zoom.connect(zoom_dialog)
        self.threadpool.start(worker)

    def pass_(self, *args, **kwargs):
        pass

    def run_without_waiting(self, func: Callable, args: list):
        """Runs a function without waiting for it to finish"""
        worker = Worker(func, *args)
        if func not in (on_submit, sample_function):
            if not wmi_connectable():
                print_error(ui.console, "Couldn't connect to remote WMI")
                config.disable = True
                self.on_done()
                return
            if not reg_connect():
                print_error(ui.console, "Couldn't connect to remote registry")
                config.disable = True
                self.on_done()
                return
        worker.signals.call_pb.connect(self.call_pg)
        worker.signals.progress.connect(self.update_pg)
        worker.signals.clean_pb.connect(self.clean_pg)
        worker.signals.finished.connect(self.pass_)
        worker.signals.progress.connect(self.update_pg)
        worker.signals.print_.connect(print)
        worker.signals.print_error_.connect(print_error)
        worker.signals.print_success_.connect(print_success)
        worker.signals.ask_yes_no.connect(YesNo)
        worker.signals.show_info_.connect(show_info)
        worker.signals.copy.connect(copy_clip)
        worker.signals.update.connect(update)
        worker.signals.update_error.connect(update_error)
        worker.signals.update_success.connect(update_success)
        worker.signals.enable_1.connect(enable_1)
        worker.signals.disable_1.connect(disable_1)
        worker.signals.clear_all.connect(clear_all)
        worker.signals.del_users.connect(del_users_selection_actual)
        worker.signals.run_without_waiting.connect(self.run_without_waiting)
        worker.signals.zoom.connect(zoom_dialog)
        self.threadpool.start(worker)

    def on_done(self) -> None:
        """Handles the UI after finishing with a worker call"""
        enable_1(ui.submit_btn)
        enable_1(ui.settings_btn)
        disable() if config.disable else enable()
        self.computer_entry.returnPressed.connect(lambda: self.runit(on_submit))

    def closeEvent(self, event):
        """Closes the UI without waiting for the qt threads, not graceful, remove it if you wish"""
        os._exit(0)


if __name__ == "__main__":
    sys._excepthook = sys.excepthook
    refresh = QCoreApplication.processEvents
    sys.excepthook = my_exception_hook
    app = QtWidgets.QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    ui = GUI()
    sys.stdout.write = redirect
    clear_all(first=True)
    disable()
    Formats.normal = '<span style="color:rgb%%;">{}</span>'.replace("%%", config.colors["text_color"])
    for obj_, name_ in (
            (ui.pc_display, "ui.pc_display"), (ui.pc_status, "ui.pc_status"), (ui.user_display, "ui.user_display"),
            (ui.uptime_display, "ui.uptime_display"), (ui.c_space_display, "ui.c_space_display"),
            (ui.d_space_display, "ui.d_space_display"), (ui.ram_display, "ui.ram_display"),
            (ui.ie_display, "ui.ie_display"), (ui.cpt_status, "ui.cpt_status"), (ui.user_status, "ui.user_status"),
            (ui.console, "ui.console")
    ):
        Objects.objects[name_] = obj_
    for obj_, name_ in (
            (ui.submit_btn, "ui.submit_btn"), (ui.copy_btn, "ui.copy_btn"), (ui.restart_pc_btn, "ui.restart_pc_btn"),
            (ui.settings_btn, "ui.settings_btn"), (ui.export_btn, "ui.export_btn"), (ui.fix_ie_btn, "ui.fix_ie_btn"),
            (ui.del_teams_btn, "ui.del_teams_btn"), (ui.del_zoom_btn, "ui.del_zoom_btn"),
            (ui.fix_cpt_btn, "ui.fix_cpt_btn"), (ui.clear_space_btn, "ui.clear_space_btn"),
            (ui.reset_spool_btn, "ui.reset_spool_btn"), (ui.del_users_btn, "ui.del_users_btn"),
            (ui.close_outlook_btn, "ui.close_outlook_btn"), (ui.printers_btn, "ui.printers_btn"),
            (ui.sample_btn, "ui.sample_btn"), (ui.del_ost_btn, "ui.del_ost_btn"),
            (ui.fix_3_lang_btn, "ui.fix_3_lang_btn")
    ):
        Objects.buttons[name_] = obj_
    Objects.pc_display = "ui.pc_display"
    Objects.pc_status = "ui.pc_status"
    Objects.user_display = "ui.user_display"
    Objects.uptime_display = "ui.uptime_display"
    Objects.c_space_display = "ui.c_space_display"
    Objects.d_space_display = "ui.d_space_display"
    Objects.ram_display = "ui.ram_display"
    Objects.ie_display = "ui.ie_display"
    Objects.cpt_status = "ui.cpt_status"
    Objects.user_status = "ui.user_status"
    Objects.console = "ui.console"
    Objects.submit_btn = "ui.submit_btn"
    Objects.copy_btn = "ui.copy_btn"
    Objects.restart_pc_btn = "ui.restart_pc_btn"
    Objects.settings_btn = "ui.settings_btn"
    Objects.export_btn = "ui.export_btn"
    Objects.fix_ie_btn = "ui.fix_ie_btn"
    Objects.del_teams_btn = "ui.del_teams_btn"
    Objects.del_zoom_btn = "ui.del_zoom_btn"
    Objects.fix_cpt_btn = "ui.fix_cpt_btn"
    Objects.clear_space_btn = "ui.clear_space_btn"
    Objects.reset_spool_btn = "ui.reset_spool_btn"
    Objects.del_users_btn = "ui.del_users_btn"
    Objects.close_outlook_btn = "ui.close_outlook_btn"
    Objects.printers_btn = "ui.printers_btn"
    Objects.sample_btn = "ui.sample_btn"
    Objects.del_ost_btn = "ui.del_ost_btn"
    Objects.fix_3_lang_btn = "ui.fix_3_lang_btn"
    try:
        sys.exit(app.exec())
    except:
        log()
