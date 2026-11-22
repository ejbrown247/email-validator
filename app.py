
"""Tkinter front-end for email syntax, DNS, and optional SMTP validation."""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from validator import load_data, email_columns, validate_file


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('Email Validator')
        self.geometry('700x440')
        self.minsize(600, 410)

        self.source = None
        self.headers = []
        self.column = tk.StringVar()
        self.use_smtp = tk.BooleanVar(value=True)
        self.message = tk.StringVar(value='Select an Excel or CSV file to begin.')
        self.progress = tk.DoubleVar(value=0)

        frame = ttk.Frame(self, padding=22)
        frame.pack(fill='both', expand=True)

        ttk.Label(
            frame, text='Email Validator',
            font=('Arial', 20, 'bold')
        ).pack(anchor='w')

        ttk.Label(
            frame,
            text='Checks email formatting, mail routing, and optionally '
                 'attempts to verify individual mailboxes.',
            wraplength=640
        ).pack(anchor='w', pady=(6, 18))

        self.pick = ttk.Button(
            frame, text='Choose Excel / CSV File', command=self.choose
        )
        self.pick.pack(anchor='w')

        self.filename = ttk.Label(
            frame, text='No file selected', wraplength=640
        )
        self.filename.pack(anchor='w', pady=(8, 16))

        ttk.Label(
            frame, text='Column containing email addresses:'
        ).pack(anchor='w')

        self.columns = ttk.Combobox(
            frame, textvariable=self.column,
            state='readonly', width=48
        )
        self.columns.pack(anchor='w', pady=(5, 14))

        self.smtp_checkbox = ttk.Checkbutton(
            frame,
            text='Attempt SMTP mailbox verification',
            variable=self.use_smtp
        )
        self.smtp_checkbox.pack(anchor='w')

        ttk.Label(
            frame,
            text='SMTP checks may be blocked by networks or mail servers. '
                 'An accepted address is not proof that the mailbox exists.',
            wraplength=640
        ).pack(anchor='w', pady=(4, 14))

        self.run = ttk.Button(
            frame,
            text='Validate and Save Results…',
            command=self.start,
            state='disabled'
        )
        self.run.pack(anchor='w')

        ttk.Progressbar(
            frame, variable=self.progress, maximum=100
        ).pack(fill='x', pady=(16, 6))

        ttk.Label(
            frame, textvariable=self.message, wraplength=640
        ).pack(anchor='w')

    def choose(self):
        path = filedialog.askopenfilename(
            title='Select spreadsheet',
            filetypes=[('Excel and CSV', '*.xlsx *.csv')]
        )
        if not path:
            return

        try:
            headers, _ = load_data(path)
            self.headers = headers
            self.source = path
            self.filename.configure(text=path)

            self.columns['values'] = [
                f'{i+1}: {h or "(unnamed)"}'
                for i, h in enumerate(headers)
            ]

            matches = email_columns(headers)
            self.columns.current(matches[0] if matches else 0)

            self.run.configure(state='normal')
            self.message.set(
                'Ready. Select the email column, then validate.'
            )
        except Exception as exc:
            messagebox.showerror('Cannot read file', str(exc))

    def start(self):
        if not self.source or self.columns.current() < 0:
            return

        output = filedialog.asksaveasfilename(
            title='Save validation results',
            defaultextension='.xlsx',
            initialfile=Path(self.source).stem + '_validated.xlsx',
            filetypes=[('Excel workbook', '*.xlsx')]
        )
        if not output:
            return

        if Path(output).resolve() == Path(self.source).resolve():
            messagebox.showerror(
                'Choose another name',
                'The output must not overwrite the input file.'
            )
            return

        self.run.configure(state='disabled')
        self.pick.configure(state='disabled')
        self.smtp_checkbox.configure(state='disabled')

        self.progress.set(0)
        self.message.set('Checking addresses…')

        index = self.columns.current()
        use_smtp = self.use_smtp.get()

        threading.Thread(
            target=self.worker,
            args=(self.source, output, index, use_smtp),
            daemon=True
        ).start()

    def worker(self, source, output, index, use_smtp):
        try:
            def progress(done, total):
                self.after(
                    0,
                    lambda d=done, t=total: (
                        self.progress.set(100 * d / max(1, t)),
                        self.message.set(
                            f'Checked {d} of {t} rows…'
                        )
                    )
                )

            totals = validate_file(
                source,
                output,
                index,
                progress,
                use_smtp=use_smtp
            )

            summary = ', '.join(
                f'{status}: {count}'
                for status, count in sorted(totals.items())
            ) or 'No data rows'

            self.after(
                0,
                lambda: self.finish(
                    f'Saved: {output}\n{summary}', None
                )
            )

        except Exception as exc:
            self.after(
                0,
                lambda error=str(exc): self.finish('', error)
            )

    def finish(self, message, error):
        self.run.configure(state='normal')
        self.pick.configure(state='normal')
        self.smtp_checkbox.configure(state='normal')

        if error:
            self.message.set('Validation failed.')
            messagebox.showerror('Error', error)
        else:
            self.progress.set(100)
            self.message.set(message)
            messagebox.showinfo(
                'Finished',
                message +
                '\n\nSMTP acceptance does not prove mailbox existence.'
            )


if __name__ == '__main__':
    App().mainloop()
