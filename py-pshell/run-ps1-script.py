import subprocess, sys, time

def execute_ps1(scriptpath, args):

    tstart = time.perf_counter()
    pshell_process = subprocess.Popen(["powershell.exe", 
              scriptpath, args], 
              stdout=sys.stdout)
    output, errors = pshell_process.communicate()
    tend = time.perf_counter()

    # Need to play around with the returncode value to handle errors. Also, wrap this in a try...except 
    print("Process Response : " + str(pshell_process.returncode))
    
    print(f"Conversion time is {tend - tstart:0.4f} seconds")


if __name__ == '__main__':
    process_response = execute_ps1("D:\\work\\arbitscripts\\py-pshell\\xls-csv-converter.ps1", " -source C:\\Users\\sanji\\Downloads\\small-test-file.xlsx -dest C:\\Users\\sanji\\Downloads\\converted-files")
    
    print(process_response)