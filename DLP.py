import ctypes
import getpass
import os
import sys
import time
import signal
import traceback
from ctypes import wintypes
from datetime import datetime

import win32api
import win32con
import win32gui
import win32process

from win32gui import GetWindowText, GetForegroundWindow

import keyboard  # using module keyboard

programs = []
programs_list = []
programs_new_list = []


def name_user():
    return getpass.getuser()


def time_now():
    return datetime.now()


def enum_windows_proc(hwnd, l_param):
    if (l_param is None) or ((l_param is not None) and (win32process.GetWindowThreadProcessId(hwnd)[1] == l_param)):
        program_name = win32gui.GetWindowText(hwnd)
        if program_name:
            w_style = win32api.GetWindowLong(hwnd, win32con.GWL_STYLE)
            if w_style & win32con.WS_VISIBLE:
                programs_new_list.append(program_name)
                try:
                    if programs_list.index(program_name):
                        pass
                except:
                    programs_list.append(program_name)
                    time_star = time_now()
                    program = {'program_name': program_name, 'time_star': time_star}
                    programs.append(program)


def enum_proc_wnds(pid=None):
    win32gui.EnumWindows(enum_windows_proc, pid)


def enum_procs(proc_name=None):
    pids = win32process.EnumProcesses()
    if proc_name is not None:
        buf_len = 0x100

        _open_process = ctypes.cdll.kernel32.OpenProcess
        _open_process.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        _open_process.restype = wintypes.HANDLE

        _GetProcessImageFileName = ctypes.cdll.psapi.GetProcessImageFileNameA
        _GetProcessImageFileName.argtypes = [wintypes.HANDLE, wintypes.LPSTR, wintypes.DWORD]
        _GetProcessImageFileName.restype = wintypes.DWORD

        _close_handle = ctypes.cdll.kernel32.CloseHandle
        _close_handle.argtypes = [wintypes.HANDLE]
        _close_handle.restype = wintypes.BOOL

        filtered_pids = ()
        for pid in pids:
            try:
                h_proc = _open_process(win32con.PROCESS_ALL_ACCESS, 0, pid)
            except:
                print("Process [%d] couldn't be opened: %s" % (pid, traceback.format_exc()))
                continue
            try:
                buf = ctypes.create_string_buffer(buf_len)
                _GetProcessImageFileName(h_proc, buf, buf_len)
                if buf.value:
                    name = buf.value.decode().split(os.path.sep)[-1]
                    # print name
                else:
                    _close_handle(h_proc)
                    continue
            except:
                print("Error getting process name: %s" % traceback.format_exc())
                _close_handle(h_proc)
                continue
            if name.lower() == proc_name.lower():
                filtered_pids += (pid,)
        return filtered_pids
    else:
        return pids


# # determines if the loop is running
# running = True


# def cleanup():
#   print("Cleaning up ...")

active_windows_list = []
active_window = []
active_window_new = []


def active_windows(active_window):
    active_window_new.append(GetWindowText(GetForegroundWindow()))
    # print('active_window ken- ', active_window[len(active_window) - 2])
    # print('active_window_new', active_window_new[0])

    if not active_windows_list:
        active_program = {'active_program_name': active_window[0], 'time_star': time_now()}
        active_windows_list.append(active_program)

    if active_window_new[0] != active_window[len(active_window) - 2]:
        # print('работает')
        active_program = {'active_program_name': active_window[len(active_window) - 2], 'time_finish': time_now()}
        active_windows_list.append(active_program)
        active_program = {'active_program_name': active_window_new[0], 'time_star': time_now()}
        active_windows_list.append(active_program)

    # programs = []
    # programs_list = []
    # programs_new_list = []
    #
    # program_name = win32gui.GetWindowText(hwnd)
    # if program_name:
    #     w_style = win32api.GetWindowLong(hwnd, win32con.GWL_STYLE)
    #     if w_style & win32con.WS_VISIBLE:
    #         programs_new_list.append(program_name)
    #         try:
    #             if programs_list.index(program_name):
    #                 pass
    #         except:
    #             programs_list.append(program_name)
    #             time_star = time_now()
    #             program = {'program_name': program_name, 'time_star': time_star}
    #             programs.append(program)


        # active_program = {'active_program_name': active_window, 'time_star': time_now()}
        # active_windows_list.append(active_program)


def main(args):
    # global running
    if args:
        proc_name = args[0]
    else:
        proc_name = None
    pids = enum_procs(proc_name)
    # print(pids)
    for pid in pids:
        enum_proc_wnds(pid)

    # # add a hook for TERM (15) and INT (2)
    # signal.signal(signal.SIGTERM, _handle_signal)
    # signal.signal(signal.SIGINT, _handle_signal)

    # # is True by default - will be set False on signal
    # while running:
    #     time.sleep(5)


# # when receiving a signal ...
# def _handle_signal(signal, frame):
#     global running
#
#     # mark the loop stopped
#     running = False
#     # cleanup
#     cleanup()


# def on_exit(sig, func=None):
#     print("exit handler")
#     time.sleep(10)

if __name__ == '__main__':
    print(name_user())

    print(time_now())

    # # print("Python {:s} on {:s}\n".format(sys.version, sys.platform))
    # while True:  # making a loop
    #     try:  # used try so that if user pressed other than the given key error will not be shown
    #         if keyboard.is_pressed('q'):  # if key 'q' is pressed
    #             print('You Pressed A Key!')
    #             break  # finishing the loop
    #         else:
    #             time.sleep(10)
    #             main(sys.argv[1:])
    #     except:
    #         break  # if user pressed a key other than the given key the loop will break

    while True:
        active_window.append(GetWindowText(GetForegroundWindow()))
        active_windows(active_window)
        main(sys.argv[1:])

        print('\n\n-----------Статистика по запущенным программам (название программы; время запуска программы / время '
              'завершения работы программы)------------')
        for program in programs:
            print(program)

        print('--------------------------------------------------------------------------------------------------------'
              '------------------------------------------')

        print('\n\n-----------Статистика по активным программам (название программы; время начала активности программы '
              '/ время завершения активности программы)------------')
        for program in active_windows_list:
            print(program)

        print('--------------------------------------------------------------------------------------------------------'
              '------------------------------------------')

        time.sleep(5)

        for program in programs_list:
            try:
                if programs_new_list.index(program):
                    pass
            except:
                # print('!!!! Прорамма закрылась !!!!')
                program = {'program_name': program, 'time_finish': time_now()}
                programs.append(program)

        programs_new_list = []
        active_window_new = []
        # print('чистка - ', programs_new_list)

        # signal.signal(signal.SIGTERM, on_exit)
        # win32api.SetConsoleCtrlHandler(on_exit, True)
