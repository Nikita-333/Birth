import threading

from UI.UI import BirthdayApp
from check import run_scheduler

if __name__ == '__main__':
    app = BirthdayApp().run()

scheduler_thread = threading.Thread(target=run_scheduler)
scheduler_thread.start()
