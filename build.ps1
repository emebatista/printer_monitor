pyinstaller --onefile --hidden-import=eventlet --hidden-import=gevent print_service.py
copy dist\print_service.exe .