import os
import signal
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class RobocopyGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Robocopy GUI")
        self.root.geometry("900x650")

        self.process = None
        self.running = False

        self.src_var = tk.StringVar()
        self.dst_var = tk.StringVar()

        self.mirror_var = tk.BooleanVar(value=True)
        self.unbuffered_var = tk.BooleanVar(value=True)
        self.no_progress_var = tk.BooleanVar(value=True)
        self.no_file_list_var = tk.BooleanVar(value=True)
        self.no_dir_list_var = tk.BooleanVar(value=True)
        self.no_job_header_var = tk.BooleanVar(value=True)
        self.no_job_summary_var = tk.BooleanVar(value=True)

        self.retry_var = tk.StringVar(value="1")
        self.wait_var = tk.StringVar(value="1")
        self.mt_var = tk.StringVar(value="32")
        self.extra_var = tk.StringVar(value="")

        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 6}

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True)

        # Paths
        paths_frame = ttk.LabelFrame(main, text="Folders")
        paths_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(paths_frame, text="Source").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(paths_frame, textvariable=self.src_var, width=80).grid(
            row=0, column=1, sticky="ew", **pad
        )
        ttk.Button(paths_frame, text="Browse...", command=self.browse_source).grid(
            row=0, column=2, **pad
        )

        ttk.Label(paths_frame, text="Destination").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(paths_frame, textvariable=self.dst_var, width=80).grid(
            row=1, column=1, sticky="ew", **pad
        )
        ttk.Button(paths_frame, text="Browse...", command=self.browse_dest).grid(
            row=1, column=2, **pad
        )

        paths_frame.columnconfigure(1, weight=1)

        # Options
        options_frame = ttk.LabelFrame(main, text="Options")
        options_frame.pack(fill="x", padx=10, pady=5)

        ttk.Checkbutton(options_frame, text="/MIR (mirror)", variable=self.mirror_var).grid(
            row=0, column=0, sticky="w", **pad
        )
        ttk.Checkbutton(options_frame, text="/J (unbuffered I/O)", variable=self.unbuffered_var).grid(
            row=0, column=1, sticky="w", **pad
        )
        ttk.Checkbutton(options_frame, text="/NP", variable=self.no_progress_var).grid(
            row=0, column=2, sticky="w", **pad
        )

        ttk.Checkbutton(options_frame, text="/NFL", variable=self.no_file_list_var).grid(
            row=1, column=0, sticky="w", **pad
        )
        ttk.Checkbutton(options_frame, text="/NDL", variable=self.no_dir_list_var).grid(
            row=1, column=1, sticky="w", **pad
        )
        ttk.Checkbutton(options_frame, text="/NJH", variable=self.no_job_header_var).grid(
            row=1, column=2, sticky="w", **pad
        )
        ttk.Checkbutton(options_frame, text="/NJS", variable=self.no_job_summary_var).grid(
            row=1, column=3, sticky="w", **pad
        )

        ttk.Label(options_frame, text="Threads (/MT:)").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(options_frame, textvariable=self.mt_var, width=8).grid(
            row=2, column=1, sticky="w", **pad
        )

        ttk.Label(options_frame, text="Retries (/R:)").grid(row=2, column=2, sticky="w", **pad)
        ttk.Entry(options_frame, textvariable=self.retry_var, width=8).grid(
            row=2, column=3, sticky="w", **pad
        )

        ttk.Label(options_frame, text="Wait (/W:)").grid(row=2, column=4, sticky="w", **pad)
        ttk.Entry(options_frame, textvariable=self.wait_var, width=8).grid(
            row=2, column=5, sticky="w", **pad
        )

        ttk.Label(options_frame, text="Extra args").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(options_frame, textvariable=self.extra_var, width=60).grid(
            row=3, column=1, columnspan=5, sticky="ew", **pad
        )

        for c in range(6):
            options_frame.columnconfigure(c, weight=1)

        # Buttons
        button_frame = ttk.Frame(main)
        button_frame.pack(fill="x", padx=10, pady=5)

        self.run_btn = ttk.Button(button_frame, text="Run", command=self.start_copy)
        self.run_btn.pack(side="left", padx=5)

        self.stop_btn = ttk.Button(button_frame, text="Stop", command=self.stop_copy, state="disabled")
        self.stop_btn.pack(side="left", padx=5)

        ttk.Button(button_frame, text="Show command", command=self.show_command).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Clear log", command=self.clear_log).pack(side="left", padx=5)

        self.status_var = tk.StringVar(value="Idle")
        ttk.Label(button_frame, textvariable=self.status_var).pack(side="right", padx=5)

        # Log
        log_frame = ttk.LabelFrame(main, text="Output")
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.log = tk.Text(log_frame, wrap="word", height=20)
        self.log.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scrollbar.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=scrollbar.set)

    def browse_source(self) -> None:
        path = filedialog.askdirectory(title="Select source folder")
        if path:
            self.src_var.set(path)

    def browse_dest(self) -> None:
        path = filedialog.askdirectory(title="Select destination folder")
        if path:
            self.dst_var.set(path)

    def append_log(self, text: str) -> None:
        self.log.insert("end", text)
        self.log.see("end")

    def clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def build_command(self) -> list[str]:
        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()

        if not src or not dst:
            raise ValueError("Please select both source and destination folders.")

        if not os.path.exists(src):
            raise ValueError("Source folder does not exist.")

        cmd = ["robocopy", src, dst]

        if self.mirror_var.get():
            cmd.append("/MIR")

        mt = self.mt_var.get().strip()
        if mt:
            cmd.append(f"/MT:{mt}")

        retries = self.retry_var.get().strip()
        if retries:
            cmd.append(f"/R:{retries}")

        wait = self.wait_var.get().strip()
        if wait:
            cmd.append(f"/W:{wait}")

        if self.unbuffered_var.get():
            cmd.append("/J")
        if self.no_progress_var.get():
            cmd.append("/NP")
        if self.no_file_list_var.get():
            cmd.append("/NFL")
        if self.no_dir_list_var.get():
            cmd.append("/NDL")
        if self.no_job_header_var.get():
            cmd.append("/NJH")
        if self.no_job_summary_var.get():
            cmd.append("/NJS")

        extra = self.extra_var.get().strip()
        if extra:
            cmd.extend(extra.split())

        return cmd

    def show_command(self) -> None:
        try:
            cmd = self.build_command()
            self.append_log("Command:\n" + " ".join(cmd) + "\n\n")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def start_copy(self) -> None:
        if self.running:
            return

        try:
            cmd = self.build_command()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.running = True
        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.status_var.set("Running")
        self.append_log("Starting:\n" + " ".join(cmd) + "\n\n")

        thread = threading.Thread(target=self._run_process, args=(cmd,), daemon=True)
        thread.start()

    def _run_process(self, cmd: list[str]) -> None:
        try:
            creationflags = 0
            if os.name == "nt":
                creationflags = subprocess.CREATE_NEW_PROCESS_GROUP

            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=creationflags,
            )

            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.root.after(0, self.append_log, line)

            rc = self.process.wait()
            self.root.after(0, self._finish_run, rc)

        except FileNotFoundError:
            self.root.after(0, self.append_log, "Error: robocopy not found. This tool only works on Windows.\n")
            self.root.after(0, self._finish_run, -1)
        except Exception as e:
            self.root.after(0, self.append_log, f"Error: {e}\n")
            self.root.after(0, self._finish_run, -1)

    def _finish_run(self, return_code: int) -> None:
        self.running = False
        self.process = None
        self.run_btn.config(state="normal")
        self.stop_btn.config(state="disabled")

        # Robocopy return codes: 0-7 are usually success-ish, 8+ are failures.
        if 0 <= return_code <= 7:
            self.status_var.set(f"Finished (code {return_code})")
            self.append_log(f"\nFinished with robocopy exit code {return_code}.\n")
        else:
            self.status_var.set(f"Error (code {return_code})")
            self.append_log(f"\nRobocopy ended with error code {return_code}.\n")

    def stop_copy(self) -> None:
        if not self.process:
            return

        try:
            if os.name == "nt":
                self.process.send_signal(signal.CTRL_BREAK_EVENT)
            else:
                self.process.terminate()
            self.append_log("\nStop requested.\n")
            self.status_var.set("Stopping...")
        except Exception as e:
            self.append_log(f"\nFailed to stop process: {e}\n")


if __name__ == "__main__":
    root = tk.Tk()
    app = RobocopyGUI(root)
    root.mainloop()