#!/usr/bin/env python3

import subprocess
import sys
import PySimpleGUI as sg


def main():

    working_directory = '/mnt/c'

    sg.theme('HotDogStand')

    layout = [
                [sg.Text('EB Linelist CSV File Path:'), sg.In(size=(63,1), enable_events=True ,key='-file_path-'), sg.FileBrowse(initial_folder=working_directory)], 
                [sg.Text('Output Directory:              '), sg.In(size=(63,1), enable_events=True ,key='-output_dir-'), sg.FolderBrowse(initial_folder=working_directory)],
                [sg.Output(size=(90,30))],
                [sg.Button('Generate'), sg.Button('Exit')] ]

    window = sg.Window('Linelist Report Generator', layout)

    while True:             # Event Loop
        event, values = window.Read()
        if event == sg.WIN_CLOSED or event == 'Exit':  
            break

        elif event == 'Generate':
            
            csv_file = values['-file_path-']
            output_dir = values["-output_dir-"]
            
            command = f"python df_to_linelist.py {csv_file} {output_dir}"       #python command
            runCommand(cmd=command, window=window)                  
            
            command = f'explorer.exe $(wslpath -w {output_dir})'         #open folder of the output directory
            runCommand(cmd=command, window=window)                 
            

    window.Close()


def runCommand(cmd, timeout=None, window=None):
    """ run shell command
    @param cmd: command to execute
    @param timeout: timeout for command execution
    @param window: the PySimpleGUI window that the output is going to (needed to do refresh on)
    @return: (return code from command, command output)
    """
    p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    output = ''
    for line in p.stdout:
        line = line.decode(errors='replace' if (sys.version_info) < (3, 5) else 'backslashreplace').rstrip()
        output += line
        print(line)
        window.Refresh() if window else None        

    retval = p.wait(timeout)
    return (retval, output)


if __name__ == '__main__':
    main()
