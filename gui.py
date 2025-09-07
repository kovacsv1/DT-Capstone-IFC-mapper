import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ifcopenshell
import addproperty
import mapper
import os
import threading
import queue
import subprocess
import datetime

def start_gui():
    root = tk.Tk()
    root.title("IFC-Excel Mapper")
    root.geometry("1000x600")  # Increased window size for better visibility
    root.resizable(True, True)
    root.configure(bg="#f0f0f0")

    # Center the window on the screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    # Apply a modern ttk theme
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TButton", padding=6, font=("Helvetica", 10), background="#4CAF50", foreground="white")
    style.map("TButton", background=[("active", "#45a049")])
    style.configure("TLabel", font=("Helvetica", 10), background="#f0f0f0")
    style.configure("TProgressbar", thickness=20, troughcolor="#f0f0f0", background="#4CAF50")
    style.map("TProgressbar", background=[("active", "#45a049")])

    # Left frame for mapping
    left_frame = ttk.Frame(root, padding="20 20 20 20")
    left_frame.grid(row=0, column=0, sticky="nsew")

    # Right frame for add property module
    right_frame = ttk.Frame(root, padding="20 20 20 20")
    right_frame.grid(row=0, column=1, sticky="nsew")

    root.columnconfigure(0, weight=1)
    root.columnconfigure(1, weight=1)
    root.rowconfigure(0, weight=1)

    # Title label in left frame
    title_label = ttk.Label(left_frame, text="IFC-Excel Data Mapping Tool", font=("Helvetica", 14, "bold"))
    title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky="ew")

    # IFC file selection in left frame
    ifc_label = ttk.Label(left_frame, text="IFC File:")
    ifc_label.grid(row=1, column=0, sticky="w", pady=5)
    ifc_path_var = tk.StringVar()
    ifc_entry = ttk.Entry(left_frame, textvariable=ifc_path_var, width=50)  # Increased width for long paths
    ifc_entry.grid(row=1, column=1, pady=5, padx=5, sticky="ew")
    def browse_ifc():
        path = filedialog.askopenfilename(filetypes=[("IFC files", "*.ifc")])
        if path:
            ifc_path_var.set(path)
            base, ext = os.path.splitext(path)
            suffix = suffix_var.get() or "mapped"
            output_path_var.set(f"{base}_{suffix}{ext}")
    ifc_button = ttk.Button(left_frame, text="Browse", command=browse_ifc)
    ifc_button.grid(row=1, column=2, pady=5, padx=5)

    # Excel file selection in left frame
    excel_label = ttk.Label(left_frame, text="Excel File:")
    excel_label.grid(row=2, column=0, sticky="w", pady=5)
    excel_path_var = tk.StringVar()
    excel_entry = ttk.Entry(left_frame, textvariable=excel_path_var, width=50)  # Increased width for long paths
    excel_entry.grid(row=2, column=1, pady=5, padx=5, sticky="ew")
    def browse_excel():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            excel_path_var.set(path)
    excel_button = ttk.Button(left_frame, text="Browse", command=browse_excel)
    excel_button.grid(row=2, column=2, pady=5, padx=5)

    # Suffix field
    suffix_label = ttk.Label(left_frame, text="Output Suffix:")
    suffix_label.grid(row=3, column=0, sticky="w", pady=5)
    suffix_var = tk.StringVar(value="mapped")
    suffix_entry = ttk.Entry(left_frame, textvariable=suffix_var, width=20)
    suffix_entry.grid(row=3, column=1, pady=5, padx=5, sticky="w")
    def update_output_path(*args):
        ifc_path = ifc_path_var.get()
        if ifc_path:
            base, ext = os.path.splitext(ifc_path)
            suffix = suffix_var.get() or "mapped"
            output_path_var.set(f"{base}_{suffix}{ext}")
    suffix_var.trace("w", update_output_path)

    # Output path selection
    output_label = ttk.Label(left_frame, text="Output IFC:")
    output_label.grid(row=4, column=0, sticky="w", pady=5)
    output_path_var = tk.StringVar()
    output_entry = ttk.Entry(left_frame, textvariable=output_path_var, width=50)  # Increased width for long paths
    output_entry.grid(row=4, column=1, pady=5, padx=5, sticky="ew")
    def browse_output():
        path = filedialog.asksaveasfilename(
            defaultextension=".ifc",
            filetypes=[("IFC files", "*.ifc")],
            initialfile=os.path.basename(output_path_var.get())
        )
        if path:
            output_path_var.set(path)
    output_button = ttk.Button(left_frame, text="Browse", command=browse_output)
    output_button.grid(row=4, column=2, pady=5, padx=5)

    # Progress bar and percentage label in left frame
    progress_frame = ttk.Frame(left_frame)
    progress_frame.grid(row=5, column=0, columnspan=3, pady=(20, 10), sticky="ew")
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100, length=400)
    progress_bar.grid(row=0, column=0, sticky="ew")
    percentage_label = ttk.Label(progress_frame, text="0.0%", font=("Helvetica", 10))
    percentage_label.grid(row=0, column=1, padx=10)

    # Status text box with scrollbar in left frame (Report box)
    status_frame = ttk.Frame(left_frame)
    status_frame.grid(row=6, column=0, columnspan=3, pady=(10, 10), sticky="nsew")
    status_text = tk.Text(status_frame, height=10, width=70, font=("Helvetica", 10), wrap="word", bg="#ffffff", relief="flat", borderwidth=1)  # Increased size
    status_text.grid(row=0, column=0, sticky="nsew")
    status_scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=status_text.yview)
    status_scrollbar.grid(row=0, column=1, sticky="ns")
    status_text.configure(yscrollcommand=status_scrollbar.set)
    status_text.configure(state="normal")
    status_text.insert("end", "Ready\n")
    status_text.configure(state="disabled")

    # Buttons frame for Run, Abort, and About in left frame
    button_frame = ttk.Frame(left_frame)
    button_frame.grid(row=7, column=0, columnspan=3, pady=10, sticky="ew")
    run_button = ttk.Button(button_frame, text="Run Mapping")
    run_button.grid(row=0, column=0, padx=(0, 5))
    abort_button = ttk.Button(button_frame, text="Abort", state="disabled")
    abort_button.grid(row=0, column=1, padx=(5, 5))
    about_button = ttk.Button(button_frame, text="?", command=lambda: show_about())
    about_button.grid(row=0, column=2, padx=(5, 0))
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=1)
    button_frame.columnconfigure(2, weight=1)

    # Configure grid weights for left frame
    left_frame.columnconfigure(1, weight=1)
    left_frame.rowconfigure(6, weight=1)  # Ensure status text box expands
    progress_frame.columnconfigure(0, weight=1)
    status_frame.columnconfigure(0, weight=1)
    status_frame.rowconfigure(0, weight=1)

    # About dialog
    def show_about():
        about_text = (
            "IFC-Excel Mapper\n"
            "Created by Valentin Kovács, 2025\n\n"
            "This software maps data from Excel to IFC files.\n\n"
            "License Information:\n"
            "- ifcopenshell: LGPL-3.0 (Source: https://github.com/IfcOpenShell/IfcOpenShell)\n"
            "- pandas: BSD-3-Clause\n"
            "- openpyxl: MIT\n"
            "- Python (tkinter): PSF License"
        )
        messagebox.showinfo("About IFC-Excel Mapper", about_text, parent=root)

    # Queue for thread-safe updates
    update_queue = queue.Queue()
    cancel_event = threading.Event()

    # Right frame for add property module
    title_label_right = ttk.Label(right_frame, text="Add Property to Element", font=("Helvetica", 14, "bold"))
    title_label_right.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky="ew")

    # Load Zones button
    ifc_file = None
    zone_dropdown_var = tk.StringVar()
    course_dropdown_var = tk.StringVar()
    pset_dropdown_var = tk.StringVar()
    prop_name_var = tk.StringVar()
    prop_value_var = tk.StringVar()

    def load_zones():
        nonlocal ifc_file
        ifc_path = ifc_path_var.get()
        if not ifc_path:
            messagebox.showerror("Error", "Please select IFC file first.", parent=root)
            return
        try:
            ifc_file = ifcopenshell.open(ifc_path)
            zones = []
            corridors = ifc_file.by_type("IfcRoad") or ifc_file.by_type("IfcFacility")
            for corridor in corridors:
                if hasattr(corridor, "IsDecomposedBy"):
                    for rel in corridor.IsDecomposedBy:
                        if hasattr(rel, "RelatedObjects"):
                            for baseline in rel.RelatedObjects:
                                if hasattr(baseline, "IsDecomposedBy"):
                                    for sub_rel in baseline.IsDecomposedBy:
                                        if hasattr(sub_rel, "RelatedObjects"):
                                            for region in sub_rel.RelatedObjects:
                                                if (
                                                    region.is_a("IfcRoadPart") and
                                                    getattr(region, "PredefinedType", None) == "ROADSEGMENT" and
                                                    getattr(region, "ObjectType", None) == "BaselineRegion"
                                                ):
                                                    zones.append(str(getattr(region, "Name", "N/A")).strip())
            zone_dropdown_var.set("Select Zone")
            zone_dropdown['values'] = sorted(zones)
            messagebox.showinfo("Success", f"Loaded {len(zones)} zones.", parent=root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load zones: {str(e)}", parent=root)

    load_zones_button = ttk.Button(right_frame, text="Load Zones", command=load_zones)
    load_zones_button.grid(row=1, column=0, columnspan=2, pady=5)

    # Zone dropdown
    zone_label = ttk.Label(right_frame, text="Zone:")
    zone_label.grid(row=2, column=0, sticky="w", pady=5)
    zone_dropdown = ttk.Combobox(right_frame, textvariable=zone_dropdown_var, state="readonly", width=40)  # Increased width
    zone_dropdown.grid(row=2, column=1, pady=5, padx=5)

    # Technique dropdown
    technique_label = ttk.Label(right_frame, text="Technique:")
    technique_label.grid(row=3, column=0, sticky="w", pady=5)
    technique_dropdown = ttk.Combobox(right_frame, textvariable=course_dropdown_var, state="readonly", width=40)  # Increased width
    technique_dropdown.grid(row=3, column=1, pady=5, padx=5)

    # Property set dropdown
    pset_label = ttk.Label(right_frame, text="Property Set:")
    pset_label.grid(row=4, column=0, sticky="w", pady=5)
    pset_dropdown = ttk.Combobox(right_frame, textvariable=pset_dropdown_var, state="readonly", width=40)  # Increased width
    pset_dropdown.grid(row=4, column=1, pady=5, padx=5)

    # Property name entry
    prop_name_label = ttk.Label(right_frame, text="Property Name:")
    prop_name_label.grid(row=5, column=0, sticky="w", pady=5)
    prop_name_entry = ttk.Entry(right_frame, textvariable=prop_name_var, width=40)  # Increased width
    prop_name_entry.grid(row=5, column=1, pady=5, padx=5)

    # Property value entry
    prop_value_label = ttk.Label(right_frame, text="Value:")
    prop_value_label.grid(row=6, column=0, sticky="w", pady=5)
    prop_value_entry = ttk.Entry(right_frame, textvariable=prop_value_var, width=40)  # Increased width
    prop_value_entry.grid(row=6, column=1, pady=5, padx=5)

    # Add Property button
    def add_property():
        nonlocal ifc_file
        if not ifc_file:
            messagebox.showerror("Error", "Load IFC file first.", parent=root)
            return
        zone = zone_dropdown_var.get()
        course_name = course_dropdown_var.get()
        pset_name = pset_dropdown_var.get()
        prop_name = prop_name_var.get()
        value = prop_value_var.get()
        if not zone or zone == "Select Zone" or not course_name or course_name == "Select Technique" or not pset_name or pset_name == "Select Property Set" or not prop_name or not value:
            messagebox.showerror("Error", "Fill all fields.", parent=root)
            return
        region = None
        corridors = ifc_file.by_type("IfcRoad") or ifc_file.by_type("IfcFacility")
        for corridor in corridors:
            if hasattr(corridor, "IsDecomposedBy"):
                for rel in corridor.IsDecomposedBy:
                    if hasattr(rel, "RelatedObjects"):
                        for baseline in rel.RelatedObjects:
                            if hasattr(baseline, "IsDecomposedBy"):
                                for sub_rel in baseline.IsDecomposedBy:
                                    if hasattr(sub_rel, "RelatedObjects"):
                                        for r in sub_rel.RelatedObjects:
                                            if (
                                                r.is_a("IfcRoadPart") and
                                                getattr(r, "PredefinedType", None) == "ROADSEGMENT" and
                                                getattr(r, "ObjectType", None) == "BaselineRegion" and
                                                str(getattr(r, "Name", "N/A")).strip() == zone
                                            ):
                                                region = r
                                                break
                            if hasattr(baseline, "ContainsElements"):
                                for sub_rel in baseline.ContainsElements:
                                    if sub_rel.is_a("IfcRelContainedInSpatialStructure") and hasattr(sub_rel, "RelatedElements"):
                                        for r in sub_rel.RelatedElements:
                                            if (
                                                r.is_a("IfcRoadPart") and
                                                getattr(r, "PredefinedType", None) == "ROADSEGMENT" and
                                                getattr(r, "ObjectType", None) == "BaselineRegion" and
                                                str(getattr(r, "Name", "N/A")).strip() == zone
                                            ):
                                                region = r
                                                break
                            if region:
                                break
                        if region:
                            break
                    if region:
                        break
        if not region:
            messagebox.showerror("Error", f"Zone '{zone}' not found.", parent=root)
            return
        courses = find_course_elements_recursively(region)
        course = next((c for c in courses if c.Name == course_name), None)
        if not course:
            messagebox.showerror("Error", f"Technique '{course_name}' not found.", parent=root)
            return
        try:
            addproperty.add_property(ifc_file, course, pset_name, prop_name, value)
            messagebox.showinfo("Success", f"Property '{prop_name}' added/overwritten in '{pset_name}' for Technique '{course_name}'.", parent=root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add property: {str(e)}", parent=root)

    add_property_button = ttk.Button(right_frame, text="Add Property", command=add_property)
    add_property_button.grid(row=7, column=0, columnspan=2, pady=10)

    # Save IFC button
    def save_ifc():
        nonlocal ifc_file
        if not ifc_file:
            messagebox.showerror("Error", "Load IFC file first.", parent=root)
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".ifc", filetypes=[("IFC files", "*.ifc")])
        if save_path:
            try:
                ifc_file.write(save_path)
                messagebox.showinfo("Success", f"IFC file saved to {save_path}.", parent=root)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save IFC file: {str(e)}", parent=root)

    save_ifc_button = ttk.Button(right_frame, text="Save Updated IFC", command=save_ifc)
    save_ifc_button.grid(row=8, column=0, columnspan=2, pady=5)

    # Zone dropdown selection handler
    def find_course_elements_recursively(current_element, depth=0):
        found_courses = []
        if hasattr(current_element, "IsDecomposedBy"):
            for rel in current_element.IsDecomposedBy:
                if hasattr(rel, "RelatedObjects"):
                    for obj in rel.RelatedObjects:
                        if obj.is_a("IfcCourse"):
                            found_courses.append(obj)
                        elif obj.is_a("IfcPavement") or obj.is_a("IfcRoadPart") or obj.is_a("IfcElement"):
                            found_courses.extend(find_course_elements_recursively(obj, depth + 1))
        if hasattr(current_element, "ContainsElements"):
            for rel in current_element.ContainsElements:
                if rel.is_a("IfcRelContainedInSpatialStructure") and hasattr(rel, "RelatedElements"):
                    for obj in rel.RelatedElements:  # Fixed: RelatedObjects -> RelatedElements
                        if obj.is_a("IfcCourse"):
                            found_courses.append(obj)
                        elif obj.is_a("IfcPavement") or obj.is_a("IfcRoadPart") or obj.is_a("IfcElement"):
                            found_courses.extend(find_course_elements_recursively(obj, depth + 1))
        return found_courses

    def on_zone_select(event):
        nonlocal ifc_file
        zone = zone_dropdown_var.get()
        if zone == "Select Zone":
            technique_dropdown['values'] = []
            course_dropdown_var.set("Select Technique")
            return
        region = None
        corridors = ifc_file.by_type("IfcRoad") or ifc_file.by_type("IfcFacility")
        for corridor in corridors:
            if hasattr(corridor, "IsDecomposedBy"):
                for rel in corridor.IsDecomposedBy:
                    if hasattr(rel, "RelatedObjects"):
                        for baseline in rel.RelatedObjects:
                            if hasattr(baseline, "IsDecomposedBy"):
                                for sub_rel in baseline.IsDecomposedBy:
                                    if hasattr(sub_rel, "RelatedObjects"):
                                        for r in sub_rel.RelatedObjects:
                                            if (
                                                r.is_a("IfcRoadPart") and
                                                getattr(r, "PredefinedType", None) == "ROADSEGMENT" and
                                                getattr(r, "ObjectType", None) == "BaselineRegion" and
                                                str(getattr(r, "Name", "N/A")).strip() == zone
                                            ):
                                                region = r
                                                break
                            if hasattr(baseline, "ContainsElements"):
                                for sub_rel in baseline.ContainsElements:
                                    if sub_rel.is_a("IfcRelContainedInSpatialStructure") and hasattr(sub_rel, "RelatedElements"):
                                        for r in sub_rel.RelatedElements:
                                            if (
                                                r.is_a("IfcRoadPart") and
                                                getattr(r, "PredefinedType", None) == "ROADSEGMENT" and
                                                getattr(r, "ObjectType", None) == "BaselineRegion" and
                                                str(getattr(r, "Name", "N/A")).strip() == zone
                                            ):
                                                region = r
                                                break
                            if region:
                                break
                        if region:
                            break
                    if region:
                        break
        if region:
            courses = find_course_elements_recursively(region)
            course_names = [c.Name for c in courses if c.Name]
            technique_dropdown['values'] = sorted(course_names)
            course_dropdown_var.set("Select Technique")
        else:
            technique_dropdown['values'] = []
            course_dropdown_var.set("Select Technique")

    zone_dropdown.bind("<<ComboboxSelected>>", on_zone_select)

    # Technique dropdown selection handler
    def on_technique_select(event):
        nonlocal ifc_file
        zone = zone_dropdown_var.get()
        course_name = course_dropdown_var.get()
        if course_name == "Select Technique":
            pset_dropdown['values'] = []
            pset_dropdown_var.set("Select Property Set")
            return
        region = None
        corridors = ifc_file.by_type("IfcRoad") or ifc_file.by_type("IfcFacility")
        for corridor in corridors:
            if hasattr(corridor, "IsDecomposedBy"):
                for rel in corridor.IsDecomposedBy:
                    if hasattr(rel, "RelatedObjects"):
                        for baseline in rel.RelatedObjects:
                            if hasattr(baseline, "IsDecomposedBy"):
                                for sub_rel in baseline.IsDecomposedBy:
                                    if hasattr(sub_rel, "RelatedObjects"):
                                        for r in sub_rel.RelatedObjects:
                                            if (
                                                r.is_a("IfcRoadPart") and
                                                getattr(r, "PredefinedType", None) == "ROADSEGMENT" and
                                                getattr(r, "ObjectType", None) == "BaselineRegion" and
                                                str(getattr(r, "Name", "N/A")).strip() == zone
                                            ):
                                                region = r
                                                break
                            if hasattr(baseline, "ContainsElements"):
                                for sub_rel in baseline.ContainsElements:
                                    if sub_rel.is_a("IfcRelContainedInSpatialStructure") and hasattr(sub_rel, "RelatedElements"):
                                        for r in sub_rel.RelatedElements:
                                            if (
                                                r.is_a("IfcRoadPart") and
                                                getattr(r, "PredefinedType", None) == "ROADSEGMENT" and
                                                getattr(r, "ObjectType", None) == "BaselineRegion" and
                                                str(getattr(r, "Name", "N/A")).strip() == zone
                                            ):
                                                region = r
                                                break
                            if region:
                                break
                        if region:
                            break
                    if region:
                        break
        if region:
            courses = find_course_elements_recursively(region)
            course = next((c for c in courses if c.Name == course_name), None)
            if course:
                psets = []
                for definition in course.IsDefinedBy:
                    if definition.RelatingPropertyDefinition.is_a("IfcPropertySet"):
                        psets.append(definition.RelatingPropertyDefinition.Name)
                pset_dropdown['values'] = sorted(psets)
                pset_dropdown_var.set("Select Property Set")
            else:
                pset_dropdown['values'] = []
                pset_dropdown_var.set("Select Property Set")
        else:
            pset_dropdown['values'] = []
            pset_dropdown_var.set("Select Property Set")

    technique_dropdown.bind("<<ComboboxSelected>>", on_technique_select)

    def update_gui():
        try:
            for _ in range(10):  # Limit to 10 items per cycle
                type_, data = update_queue.get_nowait()
                if type_ == "status":
                    status_text.configure(state="normal")
                    status_text.insert("end", data + "\n")
                    status_text.see("end")
                    status_text.configure(state="disabled")
                elif type_ == "progress":
                    current, total = data
                    percentage = (current / total) * 100
                    progress_var.set(percentage)
                    percentage_label.configure(text=f"{percentage:.1f}%")
                elif type_ == "complete":
                    log_file = data
                    root.after(0, lambda: prompt_open_log(log_file))
        except queue.Empty:
            pass
        if update_queue.qsize() > 100:
            while not update_queue.empty():
                try:
                    update_queue.get_nowait()
                except queue.Empty:
                    break
        root.after(50, update_gui)

    def run_mapping():
        ifc_path = ifc_path_var.get()
        excel_path = excel_path_var.get()
        output_path = output_path_var.get()
        if not ifc_path or not excel_path or not output_path:
            messagebox.showerror("Error", "Please select IFC, Excel, and output files.", parent=root)
            return

        run_button.configure(state="disabled")
        abort_button.configure(state="normal")
        progress_var.set(0)
        percentage_label.configure(text="0.0%")
        status_text.configure(state="normal")
        status_text.delete("1.0", "end")
        status_text.insert("end", "Starting mapping...\n")
        status_text.configure(state="disabled")
        cancel_event.clear()

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(os.path.dirname(ifc_path), f"mapping_log_{timestamp}.txt")

        def mapping_thread():
            try:
                def update_progress(current, total):
                    if not cancel_event.is_set():
                        update_queue.put(("progress", (current, total)))

                def update_status(message):
                    if not cancel_event.is_set():
                        update_queue.put(("status", message))

                def complete_callback():
                    update_queue.put(("complete", log_file))

                mapper.run_mapping(ifc_path, excel_path, output_path, update_progress, update_status, cancel_event, log_file, complete_callback)
            except Exception as e:
                if not cancel_event.is_set():
                    update_queue.put(("status", f"Error: {str(e)}"))
                    root.after(0, lambda: messagebox.showerror("Error", str(e), parent=root))
            finally:
                root.after(0, lambda: run_button.configure(state="normal"))
                root.after(0, lambda: abort_button.configure(state="disabled"))

        threading.Thread(target=mapping_thread, daemon=True).start()

    def prompt_open_log(log_file):
        if messagebox.askyesno("Open Log", "Mapping completed. Would you like to view the log file?", parent=root):
            try:
                subprocess.run(["notepad.exe", log_file], check=True)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open log file: {str(e)}", parent=root)
        messagebox.showinfo("Success", "Mapping completed successfully!", parent=root)

    def cancel_mapping():
        cancel_event.set()
        update_queue.put(("status", "Aborting mapping process..."))
        run_button.configure(state="normal")
        abort_button.configure(state="disabled")
        root.after(0, lambda: prompt_open_log(log_file))

    run_button.configure(command=run_mapping)
    abort_button.configure(command=cancel_mapping)
    load_zones_button.configure(command=load_zones)

    root.after(50, update_gui)
    root.mainloop()

if __name__ == "__main__":
    start_gui()