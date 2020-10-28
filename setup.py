from cx_Freeze import setup, Executable

setup(name='Microsoft Teams Attendance Parser',
	  version='1.0',
	  executables = [Executable('MSTeamsAttendance.py')])
