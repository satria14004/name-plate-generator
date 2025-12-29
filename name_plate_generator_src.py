import csv
import copy
import os
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_ANCHOR

def duplicate_slide(pres, index):
    """
    Duplicate the slide at the given index in the presentation.
    """
    source = pres.slides[index]
    # Use blank slide layout to avoid copying content placeholders
    blank_slide_layout = pres.slide_layouts[6] # Often the blank layout
    dest = pres.slides.add_slide(blank_slide_layout)

    # Copy shapes from source to destination
    for shape in source.shapes:
        newel = copy.deepcopy(shape.element)
        dest.shapes._spTree.insert_element_before(newel, 'p:extLst')

    return pres.slides[-1]

# Removed: calculate_optimal_font_size function

def replace_text_in_shape(shape, name, org_country_text, initial_name_size=60, initial_org_size=30):
    """
    Replace text in shape with proper formatting:
    - Name at initial large size and bold, then auto-fits
    - Organization - Country at initial smaller size and normal font, then auto-fits
    - Auto-adjusts to fit within boundaries (word-wrap then shrink font)
    - All text centered
    """
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    text_frame.clear()

    # Add name paragraph
    p_name = text_frame.paragraphs[0]
    p_name.text = name
    p_name.alignment = PP_ALIGN.CENTER

    # Set name font to Cambria with initial size and bold
    if p_name.runs:
        for run in p_name.runs:
            run.font.name = 'Cambria'
            run.font.size = Pt(initial_name_size)
            run.font.bold = True

    # Add organization-country paragraph if provided
    if org_country_text:
        p_org = text_frame.add_paragraph()
        p_org.text = org_country_text
        p_org.alignment = PP_ALIGN.CENTER

        # Set org-country font to Cambria with initial size and normal font
        if p_org.runs:
            for run in p_org.runs:
                run.font.name = 'Cambria'
                run.font.size = Pt(initial_org_size)
                run.font.bold = False

    # Enable auto-fit to word-wrap and then shrink text if needed
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    # Center vertically
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

def generate_nameplates(csv_path, pptx_path, output_path):
    print("Loading data...")

    # Parse CSV Data
    participants = []
    try:
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            # Read first line to detect delimiter
            first_line = f.readline()
            f.seek(0)

            # Count delimiters in first line
            comma_count = first_line.count(',')
            semicolon_count = first_line.count(';')
            tab_count = first_line.count('\t')

            # Choose delimiter with highest count
            if semicolon_count > comma_count and semicolon_count > tab_count:
                delimiter = ';'
            elif tab_count > comma_count:
                delimiter = '\t'
            else:
                delimiter = ','

            print(f"Using delimiter: '{delimiter}'")

            reader = csv.DictReader(f, delimiter=delimiter)

            # Normalize column names (strip whitespace and handle case)
            fieldnames = [name.strip() for name in reader.fieldnames if name]
            print(f"CSV Columns found: {fieldnames}")

            for row in reader:
                # Create normalized dictionary with stripped keys and values
                normalized_row = {}
                for k, v in row.items():
                    if k:  # Only process non-None keys
                        clean_key = k.strip()
                        clean_value = v.strip() if v else ''
                        normalized_row[clean_key] = clean_value

                # Only add rows that have at least a name
                if any(normalized_row.values()):
                    participants.append(normalized_row)

    except Exception as e:
        print(f"Error reading CSV: {e}")
        import traceback
        traceback.print_exc()
        return

    print(f"Found {len(participants)} participants.")

    # Debug: print first participant
    if participants:
        print(f"First participant: {participants[0]}")
    print("Loading PowerPoint template...")

    # Load Presentation
    try:
        prs = Presentation(pptx_path)
    except Exception as e:
        print(f"Error loading PPTX: {e}")
        return

    template_slide_index = 0

    # Generate slides
    for index, person in enumerate(participants):
        # Try different possible column name variations
        name = (person.get('Name') or person.get('name') or
                person.get('NAME') or person.get('Nama') or '').strip()
        org = (person.get('Organization') or person.get('organisation') or
               person.get('ORGANIZATION') or person.get('Org') or '').strip()
        country = (person.get('Country') or person.get('country') or
                   person.get('COUNTRY') or '').strip()

        # Skip empty rows
        if not name:
            print(f"Skipping row {index + 1}: No name found")
            continue

        # Combine Org and Country with dash separator
        if org and country:
            org_country_text = f"{org} - {country}"
        elif org:
            org_country_text = org
        elif country:
            org_country_text = country
        else:
            org_country_text = ""

        print(f"Processing {index + 1}/{len(participants)}: {name} | {org_country_text}")

        # Duplicate the template slide
        new_slide = duplicate_slide(prs, template_slide_index)

        # Collect all text shapes and clear their content
        all_text_shapes = []
        for shape in new_slide.shapes:
            if shape.has_text_frame:
                shape.text_frame.clear() # Clear all existing content
                all_text_shapes.append(shape)

        main_text_shape = None

        # Heuristic: Main nameplate area is the largest text box
        max_area = 0
        for shape in all_text_shapes:
            current_area = shape.width * shape.height
            if current_area > max_area:
                max_area = current_area
                main_text_shape = shape

        if main_text_shape:
            # Apply content to the main shape (bottom part)
            replace_text_in_shape(main_text_shape, name, org_country_text)

            # Now, create the 'top reversed part' by duplicating the main_text_shape's boundary and content
            # Get properties of the main text shape
            left = main_text_shape.left
            width = main_text_shape.width
            height = main_text_shape.height

            # Calculate new 'top' for the reversed shape for symmetrical placement
            # The distance from the bottom of the main_text_shape to the bottom of the slide
            # will be used as the distance from the top of the reversed shape to the top of the slide.
            slide_height = prs.slide_height # Get the total slide height in EMUs
            bottom_shape_bottom_edge = main_text_shape.top + main_text_shape.height
            margin_from_bottom_edge = slide_height - bottom_shape_bottom_edge

            new_top = margin_from_bottom_edge # This will be in EMUs, which `add_textbox` expects

            # Add a new text box with the same dimensions as the main one
            top_reversed_shape = new_slide.shapes.add_textbox(left, new_top, width, height)

            # Fill this new text box with the same content and formatting
            replace_text_in_shape(top_reversed_shape, name, org_country_text)

            # Rotate the new shape 180 degrees
            top_reversed_shape.rotation = 180

        else:
            print(f"Warning: No suitable main text shape found on slide for {name}. No text placed.")

    # Remove the original template slide
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[0])

    # Save Output
    try:
        prs.save(output_path)
        print(f"\nSuccess! File saved to: {output_path}")
    except Exception as e:
        print(f"Error saving file: {e}")

if __name__ == "__main__":
    print("--- PowerPoint Nameplate Generator ---")

    # Get file paths from user
    csv_path = input("Enter CSV file path: ").strip()
    pptx_path = input("Enter PPTX template path: ").strip()

    if not os.path.exists(csv_path):
        print(f"Error: CSV file not found at {csv_path}")
    elif not os.path.exists(pptx_path):
        print(f"Error: PPTX file not found at {pptx_path}")
    else:
        output_path = "Generated_Nameplates.pptx"
        generate_nameplates(csv_path, pptx_path, output_path)