import subprocess, sys, time

def execute_ps1(scriptpath):
    tstart = time.perf_counter()
    p = subprocess.Popen(["powershell.exe", 
              scriptpath], 
              stdout=sys.stdout)
    p.communicate()
    tend = time.perf_counter()
    print(f"Conversion time is {tend - tstart:0.4f} seconds")


if __name__ == '__main__':
    process_response = execute_ps1("D:\\work\\py-pshell\\xls-csv-converter.ps1")
    
    print(process_response)

    # if process_response.returncode != 0:
    #     print("An error occured: %s", process_response.stderr)
    # else:
    #     print("Hello command executed successfully!")