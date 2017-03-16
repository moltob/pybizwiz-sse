"""
Utility to enter invoice data into SSE.

After importing JSON from bizwiz, the file can be uploaded to SSE (Steuersparerklärung) by
entering the corresponding data into the currently active table.
"""
import argparse
import datetime
import json
import locale
import time
import win32gui

import win32con

SSE_WINDOW_TITLE = "Gewinn-Erfassung"
SLEEP = 0.01


def send_json_export(filepath):
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
        send_invoice(hwnd, invoice)


def find_tax_application_window():
    hwnd_desktop = win32gui.GetDesktopWindow()
    hwnd_child = win32gui.GetWindow(hwnd_desktop, win32con.GW_CHILD)

    while hwnd_child:
        if win32gui.IsWindowVisible(hwnd_child):
            title = win32gui.GetWindowText(hwnd_child)
            if SSE_WINDOW_TITLE in title:
                return hwnd_child

        hwnd_child = win32gui.GetWindow(hwnd_child, win32con.GW_HWNDNEXT)

    # not found:
    return None


def send_invoice(hwnd, invoice):
    number = 'R{}'.format(invoice['number'])
    total = locale.currency(float(invoice['total']), symbol=False)
    name = '{} {}'.format(invoice['first_name'], invoice['last_name']).strip()
    date = datetime.datetime.strptime(invoice['date_paid'], '%Y-%m-%d').strftime('%d.%m.%y')
    print('Entering data for {} {}, {}, {}.'.format(number, date, name, total))

    send_field(hwnd, number)
    send_field(hwnd, date)
    send_field(hwnd, name)
    send_field(hwnd, total)


def send_field(hwnd, string):
    for char in string:
        send_char(hwnd, char)

    # enter next field:
    send_tab(hwnd)


def send_char(hwnd, char):
    win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
    time.sleep(SLEEP)


def send_tab(hwnd):
    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
    time.sleep(SLEEP)


def main():
    parser = argparse.ArgumentParser(description='Send JSON export of invoices to SSE.')
    parser.add_argument('file', help="Path to JSON export of invoices to send to SSE.")
    args = parser.parse_args()
    send_json_export(args.file)


if __name__ == '__main__':
    main()
