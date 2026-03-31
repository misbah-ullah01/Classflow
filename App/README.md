Classflow Guide 

This project can run manually or through a custom scheduled event made in Windows Task Scheduler.

First-time setup (interactive dialog with checkboxes)
On first launch, Classflow opens an interactive setup window where user can choose:
- Assignment download folder (browse feature) Enable Sticky Notes sync checkbox
- Enable timetable sync checkbox
- Calendar id text field (default: primary), for other calendars head over to Google Calendar -> Settings (Gear icon) -> Choose desired calendar -> Navigate to the integrate calendar part and copy paste the Calendar ID this field. 
- Enable Task Scheduler checkbox
- Task Scheduler time (HH:MM)
- Optional task name (default: Classflow Daily)
- To check the task, press Windows + R on your keyboard, type taskschd.msc in the Run window, find your created task in the Active Tasks tab, you may delete the task from there. 
- Command to manually remove leftover files incase automatic removal fails, this command clears the assignment history and removes all files from Local Appdata 
- Remove-Item -Path "$env:LOCALAPPDATA\Classflow" -Recurse -Force -ErrorAction SilentlyContinue; Remove-Item -Path ".\assignment_history.json",".\deadlines.txt",".\assignment_deadlines.ics",".\assignment_deadlines_delta.ics" -Force -ErrorAction SilentlyContinue


Notes:
- No default assignment folder is used.
- If calendar sync is enabled and no calendar id is provided, primary is used.
- If scheduler is enabled, one daily task is created at the provided time.
- Setup values are saved under LOCALAPPDATA\\Classflow\\classflow_settings.json.
