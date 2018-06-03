#pylint: skip-file

from cx_Freeze import setup, Executable

base = None

executables = [Executable("work_timesheet_calculator.py", base = base)]

packages = ["idna","win32com","signal","os"]
options = {
    'build_exe': {
        'packages': packages,
    },
}

setup(
    name = "Work Timesheet Calculator - Stefanini",
    options = options,
    version = "1.0.0",
    description = 'This program calculates the workhour balance based on the downloaded timesheets.',
    executables = executables
)