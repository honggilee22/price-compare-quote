from pathlib import Path
import os
import sys
import threading
import traceback

import pythoncom
import win32api
import win32com.client
import win32con
import win32process


def exit_worker(code):
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(code)


def terminate_excel_process(process_id):
    try:
        process = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, process_id)
    except Exception:
        return
    try:
        win32process.TerminateProcess(process, 0)
    finally:
        win32api.CloseHandle(process)


def timeout_worker(process_id):
    terminate_excel_process(process_id)
    print("Excel PDF conversion timed out.", file=sys.stderr)
    exit_worker(1)


def main():
    if len(sys.argv) != 3:
        print("usage: excel_pdf_worker.py INPUT_XLSX OUTPUT_PDF", file=sys.stderr)
        exit_worker(2)

    xlsx_path = Path(sys.argv[1]).resolve()
    pdf_path = Path(sys.argv[2]).resolve()
    excel = None
    workbook = None
    excel_process_id = None
    watchdog = None

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        _, excel_process_id = win32process.GetWindowThreadProcessId(excel.Hwnd)
        watchdog = threading.Timer(25, timeout_worker, args=(excel_process_id,))
        watchdog.daemon = True
        watchdog.start()
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(xlsx_path))
        workbook.ExportAsFixedFormat(0, str(pdf_path))
        workbook.Close(False)
        excel.Quit()
        terminate_excel_process(excel_process_id)
        watchdog.cancel()
        if not pdf_path.exists():
            raise RuntimeError("Excel 변환 후 PDF 파일이 생성되지 않았습니다.")
        # Releasing pywin32 COM proxies can block for about 60 seconds even
        # after Excel has quit. This worker owns the proxies, so exit directly.
        exit_worker(0)
    except Exception:
        traceback.print_exc()
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if excel_process_id is not None:
            terminate_excel_process(excel_process_id)
        if watchdog is not None:
            watchdog.cancel()
        exit_worker(1)


if __name__ == "__main__":
    main()
