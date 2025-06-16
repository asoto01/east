import os
import re
from pypdf import PdfReader, PdfWriter

def extract_info_from_page(page_text, pdf_source_type):
    """
    Extracts student ID, advisor, and student name from the given page text
    based on the PDF source type.

    Args:
        page_text (str): The text content of a single PDF page.
        pdf_source_type (str): 'ednovate' or 'grade_document' to determine parsing rules.

    Returns:
        dict: A dictionary containing extracted information, or None if ID is not found.
    """
    extracted_data = {}
    student_id = None

    if pdf_source_type == 'ednovate':
        # Regex for Student ID: "Student ID / ID de Estudiante: 16703" (5 digits)
        id_match = re.search(r"Student ID / ID de Estudiante:\s*(\d{5})", page_text)
        if id_match:
            student_id = id_match.group(1)
            extracted_data['student_id'] = student_id
        else:
            return None # ID is crucial for Ednovate PDF, return None if not found

        # Regex for Advisor: "Advisor / Asesor: Duggan, Evan" (capture "Duggan")
        advisor_match = re.search(r"Advisor / Asesor:\s*([^,]+),", page_text)
        if advisor_match:
            extracted_data['advisor'] = advisor_match.group(1).strip()
        else:
            extracted_data['advisor'] = "UnknownAdvisor" # Default if not found
            print(f"Warning: Advisor not found for potential student ID {student_id} in Ednovate PDF.")

        # Regex for Student Name: "Student / Estudiante: Aranda, Allyson"
        student_name_match = re.search(r"Student / Estudiante:\s*([^\n]+)", page_text)
        if student_name_match:
            # Clean up the student name (remove potential trailing spaces/newlines)
            extracted_data['student_name'] = student_name_match.group(1).strip()
        else:
            extracted_data['student_name'] = f"Student_{student_id}" # Default if not found
            print(f"Warning: Student name not found for student ID {student_id} in Ednovate PDF.")

    # CHANGE HERE: Use 'grade_document' for both grade_10 and grade_11 PDFs
    # as long as their student ID parsing is the same.
    elif pdf_source_type == 'grade_document':
        # Regex for Student ID: "Student ID: 11839" (5 digits)
        id_match = re.search(r"Student ID:\s*(\d{5})", page_text)
        if id_match:
            student_id = id_match.group(1)
            extracted_data['student_id'] = student_id
        else:
            return None # ID is crucial for grade document, return None if not found

    return extracted_data

def process_pdf(pdf_path, pdf_source_type):
    """
    Reads a PDF, extracts text from each page, and maps student IDs to page objects
    and their extracted information.

    Args:
        pdf_path (str): The path to the PDF file.
        pdf_source_type (str): 'ednovate' or 'grade_document'.

    Returns:
        dict: A dictionary where keys are student IDs and values are
              dictionaries containing the page object and extracted info.
    """
    print(f"\n--- Processing '{pdf_path}' ({pdf_source_type} format) ---")
    try:
        reader = PdfReader(pdf_path)
        id_to_page_data = {}
        for i, page in enumerate(reader.pages):
            try:
                text = page.extract_text()
                if not text:
                    print(f"  Warning: Page {i+1} of '{pdf_path}' extracted no text. Skipping.")
                    continue

                extracted_info = extract_info_from_page(text, pdf_source_type)

                if extracted_info and 'student_id' in extracted_info:
                    student_id = extracted_info['student_id']
                    # Store the page object and all extracted info
                    id_to_page_data[student_id] = {
                        'page': page,
                        'info': extracted_info
                    }
                else:
                    print(f"  Could not extract student ID from page {i+1} of '{pdf_path}'. Please check regex patterns if this is unexpected.")
            except Exception as e:
                print(f"  Error processing page {i+1} of '{pdf_path}': {e}")
        print(f"--- Finished processing '{pdf_path}'. Found {len(id_to_page_data)} unique student IDs. ---")
        return id_to_page_data
    except Exception as e:
        print(f"ERROR: Could not read PDF file '{pdf_path}': {e}")
        return {}

def merge_student_pdfs(ednovate_pdf_path, grade_document_pdf_path, output_base_dir):
    """
    Merges student pages from two PDFs based on matching student IDs,
    organizing output by advisor.

    Args:
        ednovate_pdf_path (str): Path to the Ednovate PDF.
        grade_document_pdf_path (str): Path to the Grade document PDF (e.g., grade_10.pdf).
        output_base_dir (str): The base directory to save merged PDFs (e.g., 'grade_10').
    """
    # Process both PDFs to get their respective student data
    ednovate_data = process_pdf(ednovate_pdf_path, 'ednovate')
    # CHANGE HERE: Pass 'grade_document' as the type for the grade-level PDF
    grade_document_data = process_pdf(grade_document_pdf_path, 'grade_document')

    if not ednovate_data and not grade_document_data:
        print("No student data found in either PDF. Exiting.")
        return

    # Ensure the base output directory exists
    os.makedirs(output_base_dir, exist_ok=True)
    print(f"\nMerged PDFs will be saved in '{os.path.abspath(output_base_dir)}'.")

    # Get all unique student IDs from both sets for comprehensive checking
    all_student_ids = set(ednovate_data.keys()).union(set(grade_document_data.keys()))

    merged_count = 0
    skipped_count = 0
    total_ids_in_ednovate = len(ednovate_data)
    total_ids_in_grade_doc = len(grade_document_data)

    print(f"\n--- Initiating Merge Process ---")
    print(f"Students found in '{ednovate_pdf_path}': {total_ids_in_ednovate}")
    print(f"Students found in '{grade_document_pdf_path}': {total_ids_in_grade_doc}")
    print(f"Unique IDs to check for merge: {len(all_student_ids)}")

    for student_id in sorted(list(all_student_ids)): # Sort for consistent output order
        ednovate_entry = ednovate_data.get(student_id)
        grade_document_entry = grade_document_data.get(student_id)

        if ednovate_entry and grade_document_entry:
            # Both PDFs have data for this student ID, proceed with merging
            ednovate_page = ednovate_entry['page']
            grade_document_page = grade_document_entry['page']
            advisor_name = ednovate_entry['info'].get('advisor', 'UnknownAdvisor').replace(" ", "_").replace(",", "") # Sanitize for filename, remove commas
            student_name = ednovate_entry['info'].get('student_name', f"Student_{student_id}")

            # Create advisor-specific directory
            advisor_output_dir = os.path.join(output_base_dir, advisor_name)
            os.makedirs(advisor_output_dir, exist_ok=True)

            # Sanitize student name for filename (remove invalid characters and leading/trailing spaces)
            sanitized_student_name = re.sub(r'[\\/:*?"<>|]', '', student_name).strip()
            if not sanitized_student_name: # Fallback if name becomes empty after sanitization
                sanitized_student_name = f"Student_{student_id}"

            output_filename = os.path.join(advisor_output_dir, f"{sanitized_student_name}.pdf")

            writer = PdfWriter()
            writer.add_page(ednovate_page)
            writer.add_page(grade_document_page)

            try:
                with open(output_filename, "wb") as output_pdf_file:
                    writer.write(output_pdf_file)
                print(f"SUCCESS: Merged '{student_id}' ({student_name}) into '{output_filename}'")
                merged_count += 1
            except Exception as e:
                print(f"ERROR: Failed to save merged PDF for student ID {student_id} ('{student_name}') to '{output_filename}': {e}")
                skipped_count += 1
        else:
            skipped_count += 1
            if student_id not in ednovate_data:
                print(f"INFO: Student ID '{student_id}' (from grade document) was NOT found in '{ednovate_pdf_path}'. Skipping merge.")
            elif student_id not in grade_document_data:
                print(f"INFO: Student ID '{student_id}' (from Ednovate) was NOT found in '{grade_document_pdf_path}'. Skipping merge.")
            else:
                print(f"WARNING: Student ID '{student_id}' had an unexpected issue during lookup. Skipping merge.")


    print(f"\n--- Merging Summary ---")
    print(f"Total unique student IDs considered: {len(all_student_ids)}")
    print(f"Successfully merged PDFs: {merged_count}")
    print(f"Skipped students (missing from one PDF or error during save): {skipped_count}")
    print(f"Merged PDFs are located in the '{output_base_dir}' directory.")


# --- Main execution block ---
if __name__ == "__main__":
    # --- IMPORTANT: Configure these paths for grade_10 ---
    EDNOVATE_PDF = "Ednovate_East_College_Prep.pdf"
    GRADE_DOCUMENT_PDF = "grade_10.pdf" # Make sure this points to grade_10.pdf
    OUTPUT_BASE_DIR = "grade_10"        # Make sure this is 'grade_10'

    # Call the main merging function
    merge_student_pdfs(EDNOVATE_PDF, GRADE_DOCUMENT_PDF, OUTPUT_BASE_DIR)


