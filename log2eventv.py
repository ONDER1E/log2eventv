def main():
    cache_file_path = os.environ['USERPROFILE']+r"\Documents\log2eventv\.cache"
    if os.path.exists(cache_file_path):
        os.remove(cache_file_path)

    with open(config_file_path, 'r', encoding="utf8") as config_file:
        config = json.load(config_file)

    source_name = config['event_source_name']
    event_id = config['event_id']
    event_category = config['event_category']
    event_type = config['event_type']

    # Map event type from string to win32 constant
    event_type_mapping = {
        "Information": win32evtlog.EVENTLOG_INFORMATION_TYPE,
        "Warning": win32evtlog.EVENTLOG_WARNING_TYPE,
        "Error": win32evtlog.EVENTLOG_ERROR_TYPE
    }
    event_type = event_type_mapping.get(event_type, win32evtlog.EVENTLOG_INFORMATION_TYPE)

    class FileChangeHandler(FileSystemEventHandler):
        def __init__(self, file_path):
            self.file_path = file_path
            self.last_known_lines = self.get_current_lines()

        def on_modified(self, event):
            
            # Check if the modified file is the one we are watching
            if os.path.abspath(event.src_path) == os.path.abspath(self.file_path):
                self.print_new_content()

        def get_current_lines(self):
            """Reads and returns the current lines from the file."""
            try:
                with open(self.file_path, 'r', encoding="utf8") as file:
                    return file.readlines()
            except FileNotFoundError:
                print(f"File '{self.file_path}' not found.")
                return []

        def print_new_content(self):
            """Compares the new file content with the previous one and prints only the new lines."""
            current_lines = self.get_current_lines()
            
            # Find the new lines by comparing with the last known lines
            new_lines = current_lines[len(self.last_known_lines):]
            
            if new_lines:
                for line in new_lines:
                    self.write_to_event_log(line)
            else:
                print("No new content added.")
            
            # Update the last known lines
            self.last_known_lines = current_lines

        

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

    # Path to the directory containing the file
    source_name.split("\\")
    directory_to_watch = source_name.split("\\")
    file_to_watch = directory_to_watch.pop()
    directory_to_watch = "\\".join(directory_to_watch)
    print(directory_to_watch)
    print(file_to_watch)
    
    
    # Use os.path.join to ensure proper file path construction for Windows
    full_file_path = os.path.join(directory_to_watch, file_to_watch)
    
    event_handler = FileChangeHandler(full_file_path)
    observer = Observer()
    observer.schedule(event_handler, path=directory_to_watch, recursive=False)

    # Start observing the directory
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


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
    clear = lambda: os.system('cls' if os.name == 'nt' else 'clear')
    if platform != "win32":
        clear()
        print("Platform is Windows only.")
        error_count += 1
    if sys.version_info[0] < 3:
        clear()
        print("Must be using Python 3, download is here: https://www.python.org/downloads/")
        print(copy_hint)
        error_count += 1
    if import_error_count > 0:
        clear()
        print(import_error)
        if packaged_which_caused_an_import_error == "win32con":
            print("Try run 'pip install pywin32'?")
            print("If that dosen't work then try:")
            print("Try run 'pip install pypiwin32'?")
        else:
            print("Try run 'pip install " + packaged_which_caused_an_import_error + "'?")
        print(copy_hint)
        error_count += 1
    elif not os.path.exists(config_file_path):
        clear()
        directory = os.path.dirname(config_file_path)

        # Create the directory if it doesn't exist
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        f = open(config_file_path, "w", encoding="utf8")
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
        cache_file_path = os.environ['USERPROFILE'] + r"\Documents\log2eventv\.cache"
        print("THEN RUN THIS FILE AGAIN")

        # Calculate magic_number based on import_error_count
        if import_error_count > 0:
            magic_number = import_error_count - 1
        else:
            magic_number = 0

        # If the cache file does not exist, create it and store initial error count
        if not os.path.exists(cache_file_path):
            directory = os.path.dirname(cache_file_path)
            if not os.path.exists(directory):
                os.makedirs(directory)
            with open(cache_file_path, "w", encoding="utf8") as f:
                f.write(str(error_count + magic_number))

        # If the cache file exists, read the initial error count (denominator)
        if os.path.exists(cache_file_path):
            with open(cache_file_path, encoding="utf8") as f:
                first_line = f.readline().strip('\n')
                initial_error_count = int(first_line)

            # Calculate the current numerator based on import_error_count
            if import_error_count > 0:
                numerator = (error_count + (import_error_count - 1))-1
            else:
                numerator = error_count

            # Calculate progress percentage
            progress_percentage = ((numerator ) / initial_error_count) * 100
            if numerator == initial_error_count:
                progress_percentage = 100  # Ensure 100% when numerator equals denominator
            if numerator <= 1:
                progress_percentage = 100

            # Print progress and completion status
            print(f"After this step, you are {int(progress_percentage)}% done with the setup.")
            #print(f"{numerator}/{initial_error_count}")

        input("Press enter to close")

    else:
        main()