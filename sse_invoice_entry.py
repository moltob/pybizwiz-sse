"""
Utility to enter invoice data into SSE.

After importing JSON from bizwiz, the file can be uploaded to SSE (Steuersparerklärung) by
entering the corresponding data into the currently active table.
"""
import datetime
import json
import locale
import win32gui

import time
import win32con


def find_tax_application_window():
    hwnd_desktop = win32gui.GetDesktopWindow()
    hwnd_child = win32gui.GetWindow(hwnd_desktop, win32con.GW_CHILD)

    while hwnd_child:
        if win32gui.IsWindowVisible(hwnd_child):
            title = win32gui.GetWindowText(hwnd_child)
            if "Gewinn-Erfassung" in title:
                return hwnd_child

        hwnd_child = win32gui.GetWindow(hwnd_child, win32con.GW_HWNDNEXT)

    # not found:
    return None


def export_to_sse(hwnd, invoice):
    number = 'R{}'.format(invoice['number'])
    total = locale.currency(float(invoice['total']), symbol=False)
    name = '{} {}'.format(invoice['first_name'], invoice['last_name']).strip()
    date = datetime.datetime.strptime(invoice['date_paid'], '%Y-%m-%d').strftime('%d.%m.%y')
    print('Entering data for {} {}, {}, {} €.'.format(number, date, name, total))

    send(hwnd, number)
    send(hwnd, date)
    send(hwnd, name)
    send(hwnd, total)


def send(hwnd, string):
    sleep = 0.04
    for char in string:
        win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
        time.sleep(sleep)

    # enter next field:
    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
    time.sleep(sleep)


def main(filepath):
    locale.setlocale(locale.LC_ALL, '')

    with open(filepath) as file:
        export = json.load(file)
    assert export['version'] == 1, 'Unexpected export version.'

    hwnd = find_tax_application_window()
    if not hwnd:
        print('Tax application window not found, aborting.')
        return

    win32gui.SetForegroundWindow(hwnd)
    for invoice in export['invoices']:
        export_to_sse(hwnd, invoice)


if __name__ == '__main__':
    main(r'D:\Home\Downloads\invoices (2).json')
