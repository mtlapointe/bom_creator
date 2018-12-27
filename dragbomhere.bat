REM This batch script will allow CSV BOM to be run through the bom_creator.py script

REM Change to current directory (not necessary if specifying full path below):
cd /D "%~dp0"

REM Run BOM script using specified Python instance
REM Replace python.exe path and bom_creator.py path accordingly:
"C:\Users\mattlapointe\Dropbox\Code\.virtualenv\bom_creator\Scripts\python.exe" C:\Users\mattlapointe\PycharmProjects\bom_creator\bom_creator.py --file %1

PAUSE