import os
import re
from pypdf import PdfReader, PdfWriter

def extract_info_from_page(page_text, pdf_source_type):
    """
    Extracts student ID, advisor, and student name from the given page text
    based on the PDF source type.

    Args:
        page_text (str): The text content of a single PDF page.
        pdf_source_type (str): 'ednovate' or 'grade11' to determine parsing rules.

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
            return None # ID is crucial, return None if not found

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

    elif pdf_source_type == 'grade11':
        # Regex for Student ID: "Student ID: 11839" (5 digits)
        id_match = re.search(r"Student ID:\s*(\d{5})", page_text)
        if id_match:
            student_id = id_match.group(1)
            extracted_data['student_id'] = student_id
        else:
            return None # ID is crucial, return None if not found

    return extracted_data

def process_pdf(pdf_path, pdf_source_type):
    """
    Reads a PDF, extracts text from each page, and maps student IDs to page objects
    and their extracted information.

    Args:
        pdf_path (str): The path to the PDF file.
        pdf_source_type (str): 'ednovate' or 'grade11'.

    Returns:
        dict: A dictionary where keys are student IDs and values are
              dictionaries containing the page object and extracted info.
    """
    print(f"Processing '{pdf_path}' ({pdf_source_type} format)...")
    try:
        reader = PdfReader(pdf_path)
        id_to_page_data = {}
        for i, page in enumerate(reader.pages):
            try:
                text = page.extract_text()
                if not text:
                    print(f"Warning: Page {i+1} of '{pdf_path}' extracted no text. Skipping.")
                    continue

                extracted_info = extract_info_from_page(text, pdf_source_type)

                if extracted_info and 'student_id' in extracted_info:
                    student_id = extracted_info['student_id']
                    # Store the page object and all extracted info
                    id_to_page_data[student_id] = {
                        'page': page,
                        'info': extracted_info
                    }
                    # print(f"  Found ID: {student_id} on page {i+1}") # For debugging
                else:
                    print(f"  Could not extract student ID from page {i+1} of '{pdf_path}'.")
            except Exception as e:
                print(f"Error processing page {i+1} of '{pdf_path}': {e}")
        print(f"Finished processing '{pdf_path}'. Found {len(id_to_page_data)} unique student IDs.")
        return id_to_page_data
    except Exception as e:
        print(f"Error reading PDF file '{pdf_path}': {e}")
        return {}

def merge_student_pdfs(ednovate_pdf_path, grade11_pdf_path, output_base_dir="grade_11"):
    """
    Merges student pages from two PDFs based on matching student IDs,
    organizing output by advisor.

    Args:
        ednovate_pdf_path (str): Path to the Ednovate PDF.
        grade11_pdf_path (str): Path to the Grade 11 PDF.
        output_base_dir (str): The base directory to save merged PDFs (e.g., 'grade_11').
    """
    # Process both PDFs to get their respective student data
    ednovate_data = process_pdf(ednovate_pdf_path, 'ednovate')
    grade11_data = process_pdf(grade11_pdf_path, 'grade11')

    if not ednovate_data and not grade11_data:
        print("No student data found in either PDF. Exiting.")
        return

    # Ensure the base output directory exists
    os.makedirs(output_base_dir, exist_ok=True)
    print(f"\nMerged PDFs will be saved in '{os.path.abspath(output_base_dir)}'.")

    # Get all unique student IDs from both sets for comprehensive checking
    all_student_ids = set(ednovate_data.keys()).union(set(grade11_data.keys()))

    merged_count = 0
    skipped_count = 0

    for student_id in sorted(list(all_student_ids)): # Sort for consistent output order
        ednovate_entry = ednovate_data.get(student_id)
        grade11_entry = grade11_data.get(student_id)

        if ednovate_entry and grade11_entry:
            # Both PDFs have data for this student ID, proceed with merging
            ednovate_page = ednovate_entry['page']
            grade11_page = grade11_entry['page']
            advisor_name = ednovate_entry['info'].get('advisor', 'UnknownAdvisor').replace(" ", "_") # Sanitize for filename
            student_name = ednovate_entry['info'].get('student_name', f"Student_{student_id}")

            # Create advisor-specific directory
            advisor_output_dir = os.path.join(output_base_dir, advisor_name)
            os.makedirs(advisor_output_dir, exist_ok=True)

            # Sanitize student name for filename (remove invalid characters)
            sanitized_student_name = re.sub(r'[\\/:*?"<>|]', '', student_name)
            output_filename = os.path.join(advisor_output_dir, f"{sanitized_student_name}.pdf")

            writer = PdfWriter()
            writer.add_page(ednovate_page)
            writer.add_page(grade11_page)

            try:
                with open(output_filename, "wb") as output_pdf_file:
                    writer.write(output_pdf_file)
                print(f"Successfully merged '{student_id}' into '{output_filename}'")
                merged_count += 1
            except Exception as e:
                print(f"Error saving merged PDF for student ID {student_id} to '{output_filename}': {e}")
                skipped_count += 1
        else:
            # Student ID found in one PDF but not the other
            skipped_count += 1
            if ednovate_entry:
                print(f"Student ID '{student_id}' found in '{ednovate_pdf_path}' but not in '{grade11_pdf_path}'. Skipping merge.")
            elif grade11_entry:
                print(f"Student ID '{student_id}' found in '{grade11_pdf_path}' but not in '{ednovate_pdf_path}'. Skipping merge.")

    print(f"\n--- Merging Summary ---")
    print(f"Total student IDs processed: {len(all_student_ids)}")
    print(f"Successfully merged PDFs: {merged_count}")
    print(f"Skipped students (missing from one PDF or error): {skipped_count}")
    print(f"Merged PDFs are located in the '{output_base_dir}' directory.")


# --- Main execution block ---
if __name__ == "__main__":
    # Define your PDF file paths based on the provided tree structure
    EDNOVATE_PDF = "Ednovate_East_College_Prep.pdf"
    GRADE11_PDF = "grade_11.pdf"
    OUTPUT_BASE_DIR = "grade_11" # This will be the parent for advisor directories

    # Call the main merging function
    merge_student_pdfs(EDNOVATE_PDF, GRADE11_PDF, OUTPUT_BASE_DIR)


