def main():
    cache_file_path = os.environ['USERPROFILE']+r"\Documents\log2eventv\.cache"
    if os.path.exists(cache_file_path):
        os.remove(cache_file_path)

    # Load settings from config.json
    log_file_path = os.environ['USERPROFILE']+r"\Documents\log2eventv\config.json"
    with open(config_file_path, 'r') as config_file:
        config = json.load(config_file)

    source_name = config['event_source_name']
    log_type = config['event_log_type']
    event_id = config['event_id']
    event_category = config['event_category']
    event_type = config['event_type']
    polling_interval_seconds = config['polling_interval_seconds']

    # Map event type from string to win32 constant
    event_type_mapping = {
        "Information": win32evtlog.EVENTLOG_INFORMATION_TYPE,
        "Warning": win32evtlog.EVENTLOG_WARNING_TYPE,
        "Error": win32evtlog.EVENTLOG_ERROR_TYPE
    }
    event_type = event_type_mapping.get(event_type, win32evtlog.EVENTLOG_INFORMATION_TYPE)

    # Create a class to handle file changes
    class LogFileHandler(FileSystemEventHandler):
        def __init__(self, file_path):
            self.file_path = file_path
            self.last_position = 0  # Track the last read position

        def on_modified(self, event):
            # Triggered when the file is modified
            if event.src_path == self.file_path:
                with open(self.file_path, 'r') as file:
                    file.seek(self.last_position)  # Go to the last read position
                    new_logs = file.read()  # Read any new content

                    if new_logs:  # If new log content is available
                        for line in new_logs.splitlines():
                            self.write_to_event_log(line)
                    self.last_position = file.tell()  # Update the last read position

        def write_to_event_log(self, message):
            try:
                # Report the event to Windows Event Log
                win32evtlogutil.ReportEvent(
                    source_name,  # Source
                    event_id,  # Event ID
                    eventCategory=event_category,  # Event category
                    eventType=event_type,  # Event type
                    strings=[message],  # Log message
                    data=b'Python log data'  # Optional data
                )
                print(f"Forwarded log to Windows Event Log: {message}")
            except Exception as e:
                print(f"Failed to write to Event Log: {e}")

    # Create the event source (if it doesn't exist already)
    try:
        win32evtlogutil.AddSourceToRegistry(
            source_name,
            log_type,
            eventID=event_id,
            eventCategory=event_category
        )
    except Exception as e:
        print(f"Source may already exist or another error: {e}")

    # Set up the observer and file handler for file changes
    event_handler = LogFileHandler(log_file_path)
    observer = Observer()
    observer.schedule(event_handler, path=log_file_path, recursive=False)

    # Start the observer to listen for file changes
    observer.start()

    try:
        # Run indefinitely, listening for file changes
        observer.join()
    except KeyboardInterrupt:
        # Stop the observer when interrupted (Ctrl+C)
        observer.stop()

    observer.join()

    # Remove event source if needed later (optional)
    # win32evtlogutil.RemoveSourceFromRegistry(source_name)

if __name__ == "__main__":
    import_error_count = 0
    e = ""
    try:
        import win32evtlogutil
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "win32evtlogutil"
    try:
        import win32evtlog
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "win32evtlog"
    try:
        import win32con
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "win32con"
    try:
        from watchdog.observers import Observer
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "watchdog"
    try:
        from watchdog.events import FileSystemEventHandler
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "watchdog"
    try:
        import time
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "time"
    try:
        import json
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "json"
    try:
        import sys
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "sys"
    try:
        from sys import platform
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "sys"
    try:
        import os
    except ImportError as e:
        import_error_count += 1
        import_error = e
        packaged_which_caused_an_import_error = "os"

    copy_hint = "To copy in the windows console it's enter not ctrl+c"
    error_count = 0
    config_file_path = os.environ['USERPROFILE']+r"\Documents\log2eventv\config.json"
    if platform != "win32":
        print("Platform is Windows only.")
        error_count += 1
    elif sys.version_info[0] < 3:
        print("Must be using Python 3, download is here: https://www.python.org/downloads/")
        print(copy_hint)
        error_count += 1
    elif import_error_count > 0:
        print(import_error)
        if packaged_which_caused_an_import_error == "win32con":
            print("Try run 'pip install pywin32'?")
            print("If that dosen't work then try:")
            print("Try run 'pip install pypiwin32'?")
        print("Try run 'pip install " + packaged_which_caused_an_import_error + "'?")
        print(copy_hint)
        error_count += 1
    elif not os.path.exists(config_file_path):
        directory = os.path.dirname(config_file_path)

        # Create the directory if it doesn't exist
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        f = open(config_file_path, "w")
        f.write("""{
  "event_source_name": "PythonLogSource",
  "event_log_type": "Application",
  "event_id": 1000,
  "event_category": 1,
  "event_type": "Information",
  "polling_interval_seconds": 5
}""")
        f.close()
        print("go to", config_file_path, "to configure settings")
        print(copy_hint)
        error_count += 1

    if error_count > 0:
        cache_file_path = os.environ['USERPROFILE']+r"\Documents\log2eventv\.cache"
        print("THEN RUN THIS FILE AGAIN")
        if not os.path.exists(cache_file_path):
            f = open(cache_file_path, "w")
            f.write(str(error_count))
        if os.path.exists(cache_file_path):
            with open(cache_file_path) as f:
                first_line = f.readline().strip('\n')
            print("After this step you are", str(int((error_count / int(first_line)) * 100)) + f"% done with the setup.")
    else:
        main()