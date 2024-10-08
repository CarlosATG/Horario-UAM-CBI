import pandas as pd
from openpyxl import Workbook
import random
import tkinter as tk
from tkinter import simpledialog, messagebox


class ScheduleAppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("University Schedule Generator")

        # Initialize necessary variables
        self.professors = []
        self.groups = []
        self.classrooms = []
        self.start_time = 8
        self.end_time = 17
        self.trimester_type = "odd"
        self.schedule = []
        self.days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        self.class_duration = {}
        self.classes = {}
        self.selected_professor = None
        self.selected_frequency = None

        # Predefined class structure
        self.predefined_classes = {
            1: ["Cálculo Diferencial", "Mecánica Clásica", "Ecología", "Química Universitaria"],
            2: ["Cálculo Integral", "Lab de Mediciones y Mecánica", "Ondas Calor Fluidos", "Probabilidad Estadística",
                "Álgebra Lineal"],
            3: ["Cálculo de Varias Variables", "Electricidad Magnetismo", "Laboratorio Física",
                "Circuitos Eléctricos 1", "Fundamentos de Programación", "Fundamentos Diseño Lógico"],
            4: ["Ecuaciones Diferenciales", "Campos Electromagnéticos", "Dispositivos Electrónicos",
                "Circuitos Eléctricos 2", "Métodos Numéricos"],
            5: ["Matemáticas para ICT", "Acondicionamiento de Señales Eléctricas", "Programación Orientada a Objetos",
                "Diseño Lógico Avanzado"],
            6: ["Señales Sistemas", "Administración de Organizaciones", "Comunicaciones Analógicas",
                "Algoritmos Estructuras de Datos", "Sistemas Basados en Microcontroladores"],
            7: ["Control Analógico", "Bases de Datos", "Sistemas Operativos"],
            8: ["Comunicaciones Digitales", "Óptica Física Moderna", "Fundamentos de Admin de Proyectos de SW",
                "Redes de Comunicación"],
            9: ["Procesamiento Digital de Señales", "Teoría de Información Codificación", "Física Electrónica",
                "Formulación de proyecto fundamento económico"],
            10: ["Control Digital", "Laboratorio de Control", "Factibilidad tec económica financiera"],
            11: ["Emprendimiento social"],
            12: []  # Group 12 has no classes (thesis work)
        }
        self.current_group_index = 0  # To track which group/class we're on
        self.current_class_index = 0

        # Create the input form
        self.setup_gui()

    def setup_gui(self):
        # Step 1: Professor input
        tk.Label(self.root, text="Enter Professors (comma-separated):").pack()
        self.professor_entry = tk.Entry(self.root, width=50)
        self.professor_entry.pack()

        # Step 2: Classroom and time setup
        tk.Label(self.root, text="Number of classrooms:").pack()
        self.classroom_entry = tk.Entry(self.root)
        self.classroom_entry.pack()

        tk.Label(self.root, text="Starting hour (e.g., 8 for 8 AM):").pack()
        self.start_time_entry = tk.Entry(self.root)
        self.start_time_entry.pack()

        tk.Label(self.root, text="Latest hour classes can start (e.g., 17 for 5 PM):").pack()
        self.end_time_entry = tk.Entry(self.root)
        self.end_time_entry.pack()

        # Step 3: Trimester type
        tk.Label(self.root, text="Select trimester type:").pack()
        self.trimester_var = tk.StringVar(value="odd")
        tk.Radiobutton(self.root, text="Odd", variable=self.trimester_var, value="odd").pack()
        tk.Radiobutton(self.root, text="Even", variable=self.trimester_var, value="even").pack()

        # Step 4: Button to proceed to the next step
        tk.Button(self.root, text="Next: Assign Professors and Class Durations", command=self.collect_classes).pack()

    def collect_classes(self):
        # Get the entered professors
        self.professors = [p.strip() for p in self.professor_entry.get().split(',')]
        self.classrooms = [f"Room {i + 1}" for i in range(int(self.classroom_entry.get()))]
        self.start_time = int(self.start_time_entry.get())
        self.end_time = int(self.end_time_entry.get())
        self.trimester_type = self.trimester_var.get()

        self.groups = [2, 4, 6, 8, 10, 12] if self.trimester_type == "even" else [1, 3, 5, 7, 9, 11]

        # Open a window to select professor and class frequency
        self.professor_selection_window()

    def professor_selection_window(self):
        # Window for professor and class frequency selection
        self.professor_window = tk.Toplevel(self.root)
        self.professor_window.title("Assign Professors and Class Frequencies")

        # Create containers for dynamic content
        self.class_label = tk.Label(self.professor_window)
        self.class_label.pack()

        self.professor_buttons = tk.Frame(self.professor_window)
        self.professor_buttons.pack()

        self.frequency_buttons = tk.Frame(self.professor_window)
        self.frequency_buttons.pack()

        self.submit_button = tk.Button(self.professor_window, text="Submit", command=self.submit_professor_selection)
        self.submit_button.pack()

        self.display_next_class()

    def display_next_class(self):
        # Clear previous content
        for widget in self.professor_buttons.winfo_children():
            widget.destroy()
        for widget in self.frequency_buttons.winfo_children():
            widget.destroy()

        # Get the current group and class
        if self.current_group_index < len(self.groups):
            group = self.groups[self.current_group_index]
            group_classes = self.predefined_classes.get(group, [])
            if self.current_class_index < len(group_classes):
                class_name = group_classes[self.current_class_index]

                # Update class label
                self.class_label.config(text=f"Assign professor for {class_name} (Group {group})")

                # Create buttons for professors
                for professor in self.professors:
                    tk.Button(self.professor_buttons, text=professor,
                              command=lambda p=professor, g=group, c=class_name: self.select_professor(p, g, c)).pack(
                        side="left")

                # Create buttons for class frequency
                tk.Label(self.frequency_buttons, text=f"Select frequency for {class_name} (Group {group}):").pack()
                tk.Button(self.frequency_buttons, text="90 mins, 3x/week",
                          command=lambda g=group, c=class_name: self.select_frequency(g, c, 90, 3)).pack(side="left")
                tk.Button(self.frequency_buttons, text="120 mins, 1x/week",
                          command=lambda g=group, c=class_name: self.select_frequency(g, c, 120, 1)).pack(side="left")
                tk.Button(self.frequency_buttons, text="180 mins, 1x/week",
                          command=lambda g=group, c=class_name: self.select_frequency(g, c, 180, 1)).pack(side="left")

    def select_professor(self, professor, group, class_name):
        self.selected_professor = professor
        print(f"Selected {professor} for {class_name} (Group {group})")

        # Add professor to the class details
        if f"{class_name} (Group {group})" in self.classes:
            self.classes[f"{class_name} (Group {group})"]['professor'] = professor
        else:
            self.classes[f"{class_name} (Group {group})"] = {
                "group": group,
                "professor": professor,
                "classroom": random.choice(self.classrooms),
                "scheduled_times": []
            }

    def select_frequency(self, group, class_name, duration, times_per_week):
        self.selected_frequency = (duration, times_per_week)
        print(f"Selected {duration} mins, {times_per_week}x/week for {class_name} (Group {group})")

        # Add frequency details to the class
        if f"{class_name} (Group {group})" in self.classes:
            self.classes[f"{class_name} (Group {group})"]['duration'] = duration
            self.classes[f"{class_name} (Group {group})"]['times_per_week'] = times_per_week
        else:
            self.classes[f"{class_name} (Group {group})"] = {
                "group": group,
                "professor": None,
                "duration": duration,
                "times_per_week": times_per_week,
                "classroom": random.choice(self.classrooms),
                "scheduled_times": []
            }

    def submit_professor_selection(self):
        # Move to the next class or group
        self.current_class_index += 1
        group = self.groups[self.current_group_index]
        group_classes = self.predefined_classes.get(group, [])
        if self.current_class_index >= len(group_classes):
            self.current_class_index = 0
            self.current_group_index += 1

        # If we're out of groups, close the window and proceed
        if self.current_group_index >= len(self.groups):
            self.professor_window.destroy()
            self.generate_schedule()
            self.export_schedule_to_excel()
        else:
            # Display the next class
            self.display_next_class()

    def generate_time_slots(self):
        # Generate time slots for the day in 30-minute intervals
        available_times = []
        current_time = self.start_time
        while current_time + 0.5 <= self.end_time:  # 30-minute intervals
            available_times.append((current_time, current_time + 0.5))
            current_time += 0.5  # Move to the next 30-minute slot
        return available_times

    def days_are_consecutive(self, last_two_days, new_day):
        # Convert days to indices (Monday = 0, Tuesday = 1, etc.)
        day_indices = {day: index for index, day in enumerate(self.days_of_week)}

        # Get the indices of the last two scheduled days and the new day
        day1_index = day_indices[last_two_days[0]]
        day2_index = day_indices[last_two_days[1]]
        new_day_index = day_indices[new_day]

        # Check if the last two days are consecutive and if the new day is consecutive as well
        return (day2_index == day1_index + 1) and (new_day_index == day2_index + 1)

    def is_time_slot_available(self, classroom, professor, group, timeslot, day, class_name, duration):
        start_time, _ = timeslot
        class_duration_hours = duration / 60  # Convert duration from minutes to hours
        end_time = start_time + class_duration_hours  # Class end time

        # Check for availability considering consecutive classes and required break
        for scheduled_class in self.schedule:
            scheduled_start, scheduled_end = scheduled_class['timeslot']

            # Group scheduling check
            if scheduled_class['group'] == group and scheduled_class['day'] == day:
                # Ensure no more than 2 consecutive classes, and add a 1-hour break after them
                if start_time < scheduled_end and end_time > scheduled_start:
                    return False  # Time conflict for the group

                # Check if a 1-hour break is needed after 2 consecutive classes
                if scheduled_end == start_time and len(
                        [sc for sc in self.schedule if sc['group'] == group and sc['day'] == day]) % 2 == 0:
                    return False  # A break is required after 2 consecutive classes

            # Professor scheduling check
            if scheduled_class['professor'] == professor and scheduled_class['day'] == day:
                if start_time < scheduled_end and end_time > scheduled_start:
                    return False  # Time conflict for the professor

            # Classroom scheduling check
            if scheduled_class['classroom'] == classroom and scheduled_class['day'] == day:
                if start_time < scheduled_end and end_time > scheduled_start:
                    return False  # Time conflict for the classroom

        return True  # Time slot is available

    def generate_schedule(self):
        print("Generating schedule...")

        for class_name, details in self.classes.items():
            group = details["group"]
            professor = details["professor"]
            classroom = details["classroom"]
            duration = details["duration"]  # Class duration in minutes
            times_per_week = details["times_per_week"]

            print(f"Processing class: {class_name} for Group {group}, Duration: {duration} minutes")

            available_days = random.sample(self.days_of_week, len(self.days_of_week))  # Shuffle the days of the week
            scheduled_days = []  # Track the days this class has been scheduled on
            scheduled_times = 0  # Track how many times the class has been scheduled

            for day in available_days:
                if scheduled_times >= times_per_week:
                    break

                # Skip if this day would make it 3 days in a row
                if len(scheduled_days) >= 2:
                    # Check if the last two scheduled days are consecutive
                    last_two_days = [scheduled_days[-2], scheduled_days[-1]]
                    if self.days_are_consecutive(last_two_days, day):
                        print(f"Skipping {day} for {class_name} (Group {group}) to avoid 3 consecutive days")
                        continue  # Skip this day to avoid scheduling 3 days in a row

                available_times = self.generate_time_slots()

                for timeslot in available_times:
                    if self.is_time_slot_available(classroom, professor, group, timeslot, day, class_name, duration):
                        class_duration_slots = duration / 30  # Class duration in multiples of 30 minutes
                        end_time_slot = timeslot[0] + class_duration_slots * 0.5  # Total class duration in hours

                        # Add the class to the schedule
                        self.schedule.append({
                            "group": group,
                            "class_name": class_name,
                            "professor": professor,
                            "classroom": classroom,
                            "timeslot": (timeslot[0], end_time_slot),
                            "duration": duration,
                            "day": day
                        })
                        scheduled_times += 1
                        scheduled_days.append(day)  # Track the day this class was scheduled on
                        print(
                            f"Scheduled {class_name} (Group {group}) on {day} at {timeslot[0]}:00 for {duration} minutes")
                        break

            if scheduled_times < times_per_week:
                print(f"Warning: Could not schedule {class_name} (Group {group}) enough times.")

    def export_schedule_to_excel(self):
        # Export the original group-based schedule
        with pd.ExcelWriter("university_schedule_fixed.xlsx", engine="openpyxl") as writer:
            for group in self.groups:
                schedule_grid = pd.DataFrame(
                    index=[f"{int(start)}:00 - {int(start + 0.5)}:30" for start, end in self.generate_time_slots()],
                    columns=self.days_of_week)
                group_schedule = [s for s in self.schedule if s['group'] == group]

                if not group_schedule:
                    print(f"Warning: No schedule data for Group {group}. Writing empty sheet.")
                else:
                    for s_class in group_schedule:
                        start_time, end_time = s_class['timeslot']
                        class_duration_slots = int((end_time - start_time) / 0.5)  # Number of slots to fill

                        # Spread the class over the appropriate time blocks in the schedule
                        for i in range(class_duration_slots):
                            block_start = start_time + (i * 0.5)
                            row_time = f"{int(block_start)}:00 - {int(block_start + 0.5)}:30"
                            col_day = s_class['day']
                            schedule_grid.at[
                                row_time, col_day] = f"{s_class['class_name']} ({s_class['professor']}) - {s_class['classroom']}"

                # Write the schedule for this group to a separate sheet
                schedule_grid.to_excel(writer, sheet_name=f"Group {group}")

        print("\nSchedule exported to 'group_schedule.xlsx'.\n")

        # Export the schedule by professor
        self.export_schedule_by_professor()

        # Export the schedule by classroom
        self.export_schedule_by_classroom()

    def export_schedule_by_professor(self):
        # Export schedule grouped by professor
        with pd.ExcelWriter("professor_schedule.xlsx", engine="openpyxl") as writer:
            professors = set([s['professor'] for s in self.schedule])

            for professor in professors:
                schedule_grid = pd.DataFrame(
                    index=[f"{int(start)}:00 - {int(start + 0.5)}:30" for start, end in self.generate_time_slots()],
                    columns=self.days_of_week)
                professor_schedule = [s for s in self.schedule if s['professor'] == professor]

                if not professor_schedule:
                    print(f"Warning: No schedule data for Professor {professor}. Writing empty sheet.")
                else:
                    for s_class in professor_schedule:
                        start_time, end_time = s_class['timeslot']
                        class_duration_slots = int((end_time - start_time) / 0.5)

                        for i in range(class_duration_slots):
                            block_start = start_time + (i * 0.5)
                            row_time = f"{int(block_start)}:00 - {int(block_start + 0.5)}:30"
                            col_day = s_class['day']
                            schedule_grid.at[
                                row_time, col_day] = f"{s_class['class_name']} (Group {s_class['group']}) - {s_class['classroom']}"

                # Write the schedule for this professor to a separate sheet
                schedule_grid.to_excel(writer, sheet_name=f"Professor {professor}")

        print("\nProfessor schedule exported to 'professor_schedule.xlsx'.\n")

    def export_schedule_by_classroom(self):
        # Export schedule grouped by classroom
        with pd.ExcelWriter("classroom_schedule.xlsx", engine="openpyxl") as writer:
            classrooms = set([s['classroom'] for s in self.schedule])

            for classroom in classrooms:
                schedule_grid = pd.DataFrame(
                    index=[f"{int(start)}:00 - {int(start + 0.5)}:30" for start, end in self.generate_time_slots()],
                    columns=self.days_of_week)
                classroom_schedule = [s for s in self.schedule if s['classroom'] == classroom]

                if not classroom_schedule:
                    print(f"Warning: No schedule data for Classroom {classroom}. Writing empty sheet.")
                else:
                    for s_class in classroom_schedule:
                        start_time, end_time = s_class['timeslot']
                        class_duration_slots = int((end_time - start_time) / 0.5)

                        for i in range(class_duration_slots):
                            block_start = start_time + (i * 0.5)
                            row_time = f"{int(block_start)}:00 - {int(block_start + 0.5)}:30"
                            col_day = s_class['day']
                            schedule_grid.at[
                                row_time, col_day] = f"{s_class['class_name']} (Group {s_class['group']}) - {s_class['professor']}"

                # Write the schedule for this classroom to a separate sheet
                schedule_grid.to_excel(writer, sheet_name=f"Classroom {classroom}")

        print("\nClassroom schedule exported to 'classroom_schedule.xlsx'.\n")


# Initialize the main window
root = tk.Tk()
app = ScheduleAppGUI(root)
root.mainloop()
