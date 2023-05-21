import subprocess, sys
import time


def run(cmd):
    completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    return completed


if __name__ == '__main__':
    cmd_open_excel = "$Excel = New-Object -ComObject Excel.Application"
    cmd_open_excel_workbook = "$wb = $Excel.Workbooks.Open(r\"C:\Users\sanji\Downloads\sampledocs-50mb-xls-file.xls\")"
    cmd_save_csv = "foreach ($ws in $wb.Worksheets) {$ws.SaveAs(r'C:\Users\sanji\Downloads\' + $ws.name + '.csv', 6)}"
    cmd_close_excel = "$Excel.Quit()"
    
    tstart = time.perf_counter()
    hello_info = run(cmd_open_excel)
    tend = time.perf_counter()
    print(f"Excel object creation time is {tend - tstart:0.4f} seconds")

    tstart = time.perf_counter()
    hello_info = run(cmd_open_excel_workbook)
    tend = time.perf_counter()
    print(f"Excel file opening time is {tend - tstart:0.4f} seconds")

    tstart = time.perf_counter()
    hello_info = run(cmd_save_csv)
    tend = time.perf_counter()
    print(f"Conversion time is {tend - tstart:0.4f} seconds")

    tstart = time.perf_counter()
    hello_info = run(cmd_close_excel)
    tend = time.perf_counter()
    print(f"Excel object closing time is {tend - tstart:0.4f} seconds")

    if hello_info.returncode != 0:
        print("An error occured: %s", hello_info.stderr)
    else:
        print("Hello command executed successfully!")
