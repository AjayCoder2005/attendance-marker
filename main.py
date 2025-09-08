import gspread
import datetime
import webbrowser
import customtkinter
import threading
import os
import dotenv 
dotenv.load_dotenv()
# Set the appearance mode and default color theme
customtkinter.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue", "dark-blue", "green"


# Authenticate with Google Sheets using service account credentials
try:
    gc = gspread.service_account(filename=os.getenv('data'))
    gsheet = gc.open_by_url(os.getenv('gsheet'))
    wsheet = gsheet.worksheet("Sheet1")
except Exception as e:
    print(f"Error authenticating with Google Sheets: {e}")
    wsheet = None


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Interactive Attendance Marker")
        self.geometry("500x550")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=1)

        self.label = customtkinter.CTkLabel(self, text="Attendance Marker", font=customtkinter.CTkFont(size=24, weight="bold"))
        self.label.grid(row=0, column=0, padx=20, pady=20, sticky="n")

        self.input_frame = customtkinter.CTkFrame(self)
        self.input_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.input_frame.grid_columnconfigure(0, weight=1)

        self.label_input = customtkinter.CTkLabel(self.input_frame, text="Enter Absent Roll Numbers (e.g., 1, 5, 12)")
        self.label_input.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")

        self.input_entry = customtkinter.CTkEntry(self.input_frame, placeholder_text="Enter roll numbers here...")
        self.input_entry.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        self.submit_button = customtkinter.CTkButton(self.input_frame, text="Submit Attendance", command=self.submit_attendance_thread)
        self.submit_button.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")

        self.action_frame = customtkinter.CTkFrame(self)
        self.action_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.action_frame.grid_columnconfigure((0, 1), weight=1)

        self.mark_all_present_button = customtkinter.CTkButton(self.action_frame, text="Mark All Present", command=lambda: self.mark_all_thread('P'))
        self.mark_all_present_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.mark_all_absent_button = customtkinter.CTkButton(self.action_frame, text="Mark All Absent", command=lambda: self.mark_all_thread('A'))
        self.mark_all_absent_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.view_button = customtkinter.CTkButton(self, text="View Google Sheet", command=self.open_google_sheets)
        self.view_button.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        self.status_label = customtkinter.CTkLabel(self, text="", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.status_label.grid(row=4, column=0, padx=20, pady=10)

        if wsheet is None:
            self.status_label.configure(text="Error: Could not connect to Google Sheets.\nPlease check your credentials and network connection.", text_color="red")
            self.submit_button.configure(state="disabled")
            self.mark_all_present_button.configure(state="disabled")
            self.mark_all_absent_button.configure(state="disabled")

    def update_status(self, message, color="white"):
        self.status_label.configure(text=message, text_color=color)
        self.update()

    def set_buttons_state(self, state):
        self.submit_button.configure(state=state)
        self.mark_all_present_button.configure(state=state)
        self.mark_all_absent_button.configure(state=state)
        self.view_button.configure(state=state)

    def submit_attendance_thread(self):
        self.update_status("Processing...", "yellow")
        self.set_buttons_state("disabled")
        threading.Thread(target=self._submit_attendance).start()

    def _submit_attendance(self):
        absent_rollnos_input = self.input_entry.get()
        if not absent_rollnos_input:
            self.update_status("No roll numbers entered.", "red")
            self.set_buttons_state("normal")
            return
        try:
            absent_rollnos = [int(rollno.strip()) for rollno in absent_rollnos_input.split(',') if rollno.strip()]
            self._update_attendance(absent_rollnos)
            self.update_status("Attendance updated successfully!", "green")
            self.input_entry.delete(0, 'end')
        except ValueError:
            self.update_status("Invalid input. Please enter numbers separated by commas.", "red")
        except Exception as e:
            self.update_status(f"An error occurred: {e}", "red")
        finally:
            self.set_buttons_state("normal")

    def mark_all_thread(self, status):
        self.update_status("Processing...", "yellow")
        self.set_buttons_state("disabled")
        threading.Thread(target=self._mark_all, args=(status,)).start()

    def _get_new_date_column_position(self, header_row):
        if "..." in header_row:
            return header_row.index("...") + 1
        elif "Attendance %" in header_row:
            return header_row.index("Attendance %") + 1
        else:
            return len(header_row) + 1

    def _mark_all(self, status):
        try:
            header_row = wsheet.row_values(1)
            today = datetime.date.today().strftime('%d/%m/%Y')
            insert_col = self._get_new_date_column_position(header_row)
            if today not in header_row:
                wsheet.insert_cols([[]], insert_col)
                wsheet.update_cell(1, insert_col, today)
                header_row.insert(insert_col - 1, today)
            date_column = header_row.index(today) + 1
            total_rollnos = len(wsheet.col_values(1)) - 1
            updates = []
            for row_num in range(2, total_rollnos + 2):
                updates.append({'range': f'{gspread.utils.rowcol_to_a1(row_num, date_column)}', 'values': [[status]]})
            wsheet.batch_update(updates)
            self.update_status(f"All students marked as '{status}' successfully!", "green")
        except Exception as e:
            self.update_status(f"An error occurred: {e}", "red")
        finally:
            self.set_buttons_state("normal")

    def _update_attendance(self, absent_rollnos):
        header_row = wsheet.row_values(1)
        today = datetime.date.today().strftime('%d/%m/%Y')
        insert_col = self._get_new_date_column_position(header_row)
        if today not in header_row:
            wsheet.insert_cols([[]], insert_col)
            wsheet.update_cell(1, insert_col, today)
            header_row.insert(insert_col - 1, today)
        date_column = header_row.index(today) + 1
        total_rollnos = len(wsheet.col_values(1)) - 1
        updates = []
        for rollno in range(1, total_rollnos + 1):
            status = 'P'
            if rollno in absent_rollnos:
                status = 'A'
            cell = wsheet.find(str(rollno))
            if cell:
                updates.append({'range': f'{gspread.utils.rowcol_to_a1(cell.row, date_column)}', 'values': [[status]]})
        if updates:
            wsheet.batch_update(updates)

    def open_google_sheets(self):
        url = "https://docs.google.com/spreadsheets/d/1OuXbHes2XjxjxzNU26mq69yKj3012R1zbiuA1vFD35g/edit?gid=0"
        webbrowser.open_new(url)

if __name__ == "__main__":
    app = App()
    app.mainloop()
