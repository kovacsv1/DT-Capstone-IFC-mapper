import ifcopenshell
import pandas as pd
import os
import time
import sys
import threading
from io import StringIO

class LogRedirector:
    def __init__(self, log_file):
        self.log_file = log_file
        self.buffer = StringIO()

    def write(self, message):
        self.buffer.write(message)
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(message)

    def flush(self):
        self.buffer.flush()

def normalize_string(s, is_numeric=False):
    """Normalize strings for comparison: lowercase, remove spaces, normalize slashes, and handle numeric formats."""
    if s is None or pd.isna(s):
        return ""
    s = str(s).lower().strip()
    s = s.replace(" ", "").replace("/", "").replace("\\", "")
    if is_numeric:
        try:
            num = float(s)
            return str(int(num)) if num.is_integer() else str(num)
        except ValueError:
            pass
    return s

def run_mapping(ifc_path, excel_path, output_path, progress_callback, status_callback, cancel_event, log_file, complete_callback):
    sys.stdout = LogRedirector(log_file)
    print(f"Mapping started at {time.strftime('%Y-%m-%d %H:%M:%S')}")

    start_time = time.time()
    if not os.path.exists(ifc_path):
        raise FileNotFoundError(f"IFC file not found at {ifc_path}")
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found at {excel_path}")

    status_callback("Loading IFC and Excel files...")
    try:
        ifc_file = ifcopenshell.open(ifc_path)
        status_callback(f"Successfully loaded IFC file: {os.path.basename(ifc_path)}")
        print(f"Successfully loaded IFC file: {os.path.basename(ifc_path)}")
    except Exception as e:
        raise Exception(f"Error loading IFC file: {e}")

    try:
        df = pd.read_excel(excel_path)
        status_callback(f"Successfully loaded Excel file: {os.path.basename(excel_path)}")
        print(f"Successfully loaded Excel file: {os.path.basename(excel_path)}")
    except Exception as e:
        raise Exception(f"Error loading Excel file: {e}")

    def get_all_roadparts():
        roadparts = []
        corridors = ifc_file.by_type("IfcRoad") or ifc_file.by_type("IfcFacility")
        if not corridors:
            status_callback("Warning: No IfcRoad or IfcFacility found in IFC file.")
            print("Warning: No IfcRoad or IfcFacility found in IFC file.")
            return roadparts

        for corridor in corridors:
            print(f"Found corridor: '{getattr(corridor, 'Name', 'N/A')}' (GlobalId: {corridor.GlobalId})")
            if hasattr(corridor, "IsDecomposedBy"):
                for rel in corridor.IsDecomposedBy:
                    if hasattr(rel, "RelatedObjects"):
                        for baseline in rel.RelatedObjects:
                            if baseline.is_a("IfcRoadPart") or baseline.is_a("IfcElementAssembly"):
                                print(f"Found baseline: '{getattr(baseline, 'Name', 'N/A')}' (GlobalId: {baseline.GlobalId})")
                                if hasattr(baseline, "IsDecomposedBy"):
                                    for sub_rel in baseline.IsDecomposedBy:
                                        if hasattr(sub_rel, "RelatedObjects"):
                                            for region in sub_rel.RelatedObjects:
                                                if (
                                                    region.is_a("IfcRoadPart") and
                                                    getattr(region, "PredefinedType", None) == "ROADSEGMENT" and
                                                    getattr(region, "ObjectType", None) == "BaselineRegion"
                                                ):
                                                    roadparts.append(region)
                                if hasattr(baseline, "ContainsElements"):
                                    for sub_rel in baseline.ContainsElements:
                                        if sub_rel.is_a("IfcRelContainedInSpatialStructure") and hasattr(sub_rel, "RelatedElements"):
                                            for region in sub_rel.RelatedElements:
                                                if (
                                                    region.is_a("IfcRoadPart") and
                                                    getattr(region, "PredefinedType", None) == "ROADSEGMENT" and
                                                    getattr(region, "ObjectType", None) == "BaselineRegion"
                                                ):
                                                    roadparts.append(region)
        return roadparts

    roadparts = get_all_roadparts()
    total_zones = len(roadparts)
    status_callback(f"Found {total_zones} zones (IfcRoadPart with ROADSEGMENT and BaselineRegion)")
    print(f"Found {total_zones} zones (IfcRoadPart with ROADSEGMENT and BaselineRegion)")
    if total_zones == 0:
        status_callback("Warning: No matching IfcRoadPart elements found.")
        print("Warning: No matching IfcRoadPart elements found.")

    zone_to_region = {str(rp.Name).strip(): rp for rp in roadparts if rp.Name}
    status_callback(f"Found {len(zone_to_region)} unique zone names")
    print(f"Found {len(zone_to_region)} unique zone names: {', '.join(sorted(zone_to_region.keys()))}")

    def get_property(element, pset_name, prop_name):
        if not hasattr(element, "IsDefinedBy"):
            return None
        for definition in element.IsDefinedBy:
            if hasattr(definition, "RelatingPropertyDefinition"):
                if (
                    definition.RelatingPropertyDefinition.is_a("IfcPropertySet")
                    and definition.RelatingPropertyDefinition.Name == pset_name
                ):
                    for prop in definition.RelatingPropertyDefinition.HasProperties:
                        if prop.Name == prop_name:
                            if hasattr(prop, "NominalValue") and prop.NominalValue:
                                return prop.NominalValue.wrappedValue if hasattr(prop.NominalValue, "wrappedValue") else prop.NominalValue
                            return prop
        return None

    def ensure_property_set(element, pset_name):
        for definition in element.IsDefinedBy:
            if (
                definition.RelatingPropertyDefinition.is_a("IfcPropertySet")
                and definition.RelatingPropertyDefinition.Name == pset_name
            ):
                return definition.RelatingPropertyDefinition
        pset = ifc_file.createIfcPropertySet(
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=element.OwnerHistory,
            Name=pset_name,
            HasProperties=[],
        )
        ifc_file.createIfcRelDefinesByProperties(
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=element.OwnerHistory,
            RelatedObjects=[element],
            RelatingPropertyDefinition=pset,
        )
        return pset

    def find_course_elements_recursively(current_element, depth=0):
        found_courses = []
        if cancel_event.is_set():
            return found_courses
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
                    for obj in rel.RelatedElements:
                        if obj.is_a("IfcCourse"):
                            found_courses.append(obj)
                        elif obj.is_a("IfcPavement") or obj.is_a("IfcRoadPart") or obj.is_a("IfcElement"):
                            found_courses.extend(find_course_elements_recursively(obj, depth + 1))
        return found_courses

    def process_zone(zone, region, zone_rows, excel_columns_to_add, updated_count, zone_start_time, total_zones_processed, total_zones):
        matches = 0
        courses_to_process = find_course_elements_recursively(region)
        status_callback(f"Found {len(courses_to_process)} IfcCourse elements under zone '{zone}'")
        print(f"Found {len(courses_to_process)} IfcCourse elements under zone '{zone}'")
        if not courses_to_process:
            status_callback(f"Completed zone {total_zones_processed}/{total_zones}: '{zone}' with 0 matches in {time.time() - zone_start_time:.2f} seconds")
            return 0

        course_info = []
        for course in courses_to_process:
            code_name = get_property(course, "Corridor Shape Information", "CodeName")
            if code_name and " - " in str(code_name):
                technique, surface = str(code_name).split(" - ", 1)
                technique_norm = normalize_string(technique)
                surface_norm = normalize_string(surface, is_numeric=True)
                course_info.append((course, technique_norm, surface_norm))
                print(f"IfcCourse '{course.Name}' (GlobalId: {course.GlobalId}): CodeName='{code_name}' (Technique='{technique_norm}', Surface='{surface_norm}')")
            else:
                print(f"IfcCourse '{course.Name}' (GlobalId: {course.GlobalId}): Invalid or missing CodeName='{code_name}'")
                continue

        for i, row in zone_rows.iterrows():
            if cancel_event.is_set():
                print(f"Mapping cancelled during processing for zone '{zone}'")
                status_callback(f"Mapping cancelled during zone {total_zones_processed}/{total_zones}: '{zone}'")
                return 0
            technique = str(row.get("TECHNIQUE_", ""))
            surface = str(row.get("SURFACE", ""))
            if not technique or not surface:
                print(f"Skipping row for zone '{zone}': Invalid TECHNIQUE_='{technique}' or SURFACE='{surface}'")
                continue
            technique_norm = normalize_string(technique)
            surface_norm = normalize_string(surface, is_numeric=True)
            print(f"Excel row for zone '{zone}': TECHNIQUE_='{technique}', SURFACE='{surface}' (Normalized: Technique='{technique_norm}', Surface='{surface_norm}')")

            for course, course_technique_norm, course_surface_norm in course_info:
                if technique_norm == course_technique_norm and surface_norm == course_surface_norm:
                    print(f"MATCH! IfcCourse '{course.Name}' (GlobalId: {course.GlobalId}) matched with Excel row: TECHNIQUE_='{technique}', SURFACE='{surface}'")
                    pset = ensure_property_set(course, "Excel Layer Info")
                    pset.HasProperties = []

                    def add_excel_property(prop_name_in_excel):
                        if cancel_event.is_set():
                            return
                        val = row.get(prop_name_in_excel)
                        if pd.notna(val) and str(val).strip() != "":
                            prop = ifc_file.createIfcPropertySingleValue(
                                prop_name_in_excel,
                                None,
                                ifc_file.create_entity("IfcText", str(val)),
                                None,
                            )
                            pset.HasProperties = list(pset.HasProperties) + [prop]

                    for col in excel_columns_to_add:
                        add_excel_property(col)
                    updated_count[0] += 1
                    matches += 1
                    break
                else:
                    print(f"No match for row in zone '{zone}': TECHNIQUE_='{technique_norm}', SURFACE='{surface_norm}' != CodeName='{course_technique_norm} - {course_surface_norm}'")

        zone_time = time.time() - zone_start_time
        status_callback(f"Completed zone {total_zones_processed}/{total_zones}: '{zone}' with {matches} matches in {zone_time:.2f} seconds")
        print(f"Completed zone {total_zones_processed}/{total_zones}: '{zone}' with {matches} matches in {zone_time:.2f} seconds")
        # Check if all IfcCourse elements were matched
        valid_courses = len(course_info)  # Only count courses with valid CodeName
        if matches == valid_courses:
            print(f"Verification: All {matches} valid IfcCourse elements in zone '{zone}' were successfully matched.")
            status_callback(f"Verification: All {matches} valid IfcCourse elements in zone '{zone}' matched.")
        else:
            print(f"Warning: Only {matches} of {valid_courses} valid IfcCourse elements in zone '{zone}' were matched. {valid_courses - matches} courses not updated.")
            status_callback(f"Warning: Only {matches} of {valid_courses} valid IfcCourse elements in zone '{zone}' matched.")

        return matches

    updated = [0]
    excel_columns_to_add = [
        "PR_1", "PR_2", "FOND", "SURFACE", "TYPE_COUCH",
        "CHANTIER", "ENTREPRISE", "DATE_MS", "N°_ORDRE",
    ]

    status_callback("Starting data mapping process...")
    print("Starting data mapping process...")
    processed_zone_names = set()
    total_rows = len(df)

    if cancel_event.is_set():
        status_callback("Mapping cancelled before processing started.")
        print("Mapping cancelled before processing started.")
        return

    df["ZONE"] = df["ZONE"].apply(lambda x: str(x).strip() if pd.notna(x) else "")
    zone_groups = df.groupby("ZONE")
    status_callback(f"Found {len(zone_groups)} zones in Excel with {total_rows} total rows")
    print(f"Found {len(zone_groups)} zones in Excel with {total_rows} total rows")
    for zone, count in zone_groups.size().items():
        if zone in zone_to_region:
            print(f"Zone '{zone}' has {count} Excel row(s)")

    total_zones_processed = 0
    for zone, zone_rows in zone_groups:
        if cancel_event.is_set():
            status_callback("Mapping cancelled during processing.")
            print("Mapping cancelled during processing.")
            return
        if not zone:
            print("Skipping empty ZONE value")
            continue

        region = zone_to_region.get(zone)
        if not region:
            print(f"No IfcRoadPart found matching ZONE='{zone}'")
            continue

        if zone in processed_zone_names:
            continue

        total_zones_processed += 1
        zone_start_time = time.time()
        status_callback(f"Processing zone {total_zones_processed}/{total_zones}: '{getattr(region, 'Name', 'N/A')}'")
        print(f"Processing zone {total_zones_processed}/{total_zones}: '{getattr(region, 'Name', 'N/A')}' (GlobalId: {region.GlobalId})")

        zone_thread = threading.Thread(
            target=process_zone,
            args=(zone, region, zone_rows, excel_columns_to_add, updated, zone_start_time, total_zones_processed, total_zones),
            daemon=True
        )
        zone_thread.start()
        zone_thread.join()

        processed_zone_names.add(zone)
        progress_callback(total_zones_processed, total_zones)
        print(f"Updated progress: {total_zones_processed}/{total_zones} zones completed")

    if cancel_event.is_set():
        status_callback("Mapping cancelled before saving.")
        print("Mapping cancelled before saving.")
        return

    def save_ifc_file():
        try:
            status_callback("Saving file...")
            print("Saving file...")
            time.sleep(2)  # Delay to ensure GUI updates
            ifc_file.write(output_path)
            status_callback(f"Successfully updated {updated[0]} IfcCourse elements. Saved as {output_path}")
            print(f"Successfully updated {updated[0]} IfcCourse elements. Saved as {output_path}")
            complete_callback()
        except Exception as e:
            status_callback(f"Error saving updated IFC file: {str(e)}")
            print(f"Error saving updated IFC file: {str(e)}")
            raise

    save_thread = threading.Thread(target=save_ifc_file, daemon=True)
    save_thread.start()
    status_callback("IFC file save started in background...")

    end_time = time.time()
    runtime = end_time - start_time
    status_callback(f"Total runtime: {runtime:.2f} seconds")
    print(f"Total runtime: {runtime:.2f} seconds")
    print(f"Mapping finished at {time.strftime('%Y-%m-%d %H:%M:%S')}")