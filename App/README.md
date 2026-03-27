Classflow

This project can run manually or through Windows Task Scheduler.

Daily schedule support (12:00 PM and 6:00 PM) is built into classflow.py.

Install schedule (Python script mode)
1. Open PowerShell in this folder.
2. Run:
	 py classflow.py --install-schedule

Install schedule (compiled EXE mode)
1. Build your EXE as usual.
2. Run the EXE with:
	 Classflow.exe --install-schedule

The command automatically detects whether it is running as a .py script or as an .exe and creates two tasks:
- Classflow Daily 12PM
- Classflow Daily 6PM

Remove schedule
- Python script mode:
	py classflow.py --remove-schedule
- EXE mode:
	Classflow.exe --remove-schedule

Notes
- Tasks are created as daily triggers at 12:00 and 18:00.
- Existing tasks with the same names are overwritten when installing.
