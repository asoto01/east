import email
from email.policy import default
import os
from bs4 import BeautifulSoup
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter
import tkinter as tk
from tkinter import filedialog, messagebox # Import messagebox for user feedback

def eml_to_pdf_batch_converter(output_pdf_name="merged_emails.pdf"):
    """
    Converts all .eml files in a user-selected folder into a single PDF document.

    Each email's subject, sender, and date will be displayed, followed by its
    body content (HTML content is converted to plain text for simplicity).
    Attachments are not included in the PDF but their presence can be noted
    if desired (currently not implemented for brevity).

    The output PDF will be generated in the current working directory
    where the script is executed.

    Args:
        output_pdf_name (str): The name of the output PDF file.
                               Defaults to "merged_emails.pdf".
    """
    # Create a Tkinter root window, but keep it hidden
    root = tk.Tk()
    root.withdraw() # Hide the main window

    # Prompt user to select the folder containing EML files
    eml_folder_path = filedialog.askdirectory(
        title="Select Folder Containing .eml Files"
    )

    if not eml_folder_path:
        messagebox.showinfo("Operation Cancelled", "No folder selected. Aborting conversion.")
        print("No folder selected. Aborting conversion.")
        root.destroy() # Destroy the hidden root window
        return

    if not os.path.isdir(eml_folder_path):
        messagebox.showerror("Error", f"Folder '{eml_folder_path}' not found.")
        print(f"Error: Folder '{eml_folder_path}' not found.")
        root.destroy()
        return

    # Story is a list of flowables (ReportLab's content elements)
    story = []
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    heading_style = styles['h2']
    subheading_style = styles['h3']
    body_style = styles['BodyText']

    # Define the full path for the output PDF file in the current working directory
    current_working_directory = os.getcwd()
    pdf_file_path = os.path.join(current_working_directory, output_pdf_name)
    
    # Create the PDF document object
    doc = SimpleDocTemplate(pdf_file_path, pagesize=letter)

    print(f"Starting conversion of .eml files from '{eml_folder_path}'...")
    messagebox.showinfo("Conversion Started", f"Processing emails from:\n{eml_folder_path}\n\nOutput PDF will be saved as:\n{pdf_file_path}")

    # Get all .eml files in the specified folder
    eml_files = [f for f in os.listdir(eml_folder_path) if f.lower().endswith('.eml')]
    
    if not eml_files:
        messagebox.showinfo("No Files Found", f"No .eml files found in '{eml_folder_path}'.")
        print(f"No .eml files found in '{eml_folder_path}'.")
        root.destroy()
        return

    # Sort files by name to ensure a consistent order in the PDF
    eml_files.sort()

    for i, filename in enumerate(eml_files):
        file_path = os.path.join(eml_folder_path, filename)
        
        try:
            # Open the .eml file in binary read mode
            with open(file_path, 'rb') as f:
                # Parse the email message using the default policy for robustness
                msg = email.message_from_binary_file(f, policy=default)

            # Extract common email headers
            subject = msg['subject'] if msg['subject'] else "No Subject"
            from_header = msg['from'] if msg['from'] else "Unknown Sender"
            date_header = msg['date'] if msg['date'] else "Unknown Date"
            
            print(f"Processing ({i+1}/{len(eml_files)}): {subject[:70]}...") # Print progress to console

            # Add a prominent title for each email in the PDF
            story.append(Paragraph(f"--- Email {i+1}: {subject} ---", heading_style))
            story.append(Spacer(1, 0.1 * inch)) # Small space after title
            
            # Add sender and date information
            story.append(Paragraph(f"<b>From:</b> {from_header}", normal_style))
            story.append(Paragraph(f"<b>Date:</b> {date_header}", normal_style))
            story.append(Spacer(1, 0.1 * inch)) # Small space after headers

            # Extract email body content
            body_content = ""
            if msg.is_multipart():
                # Iterate over parts to find the body
                for part in msg.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get('Content-Disposition'))

                    # Prioritize HTML content, then plain text.
                    # Ensure it's not an attachment.
                    if ctype == 'text/html' and 'attachment' not in cdispo:
                        charset = part.get_content_charset()
                        html_body = part.get_payload(decode=True).decode(charset if charset else 'utf-8', errors='ignore')
                        # Use BeautifulSoup to strip HTML tags and get clean text
                        soup = BeautifulSoup(html_body, 'html.parser')
                        body_content = soup.get_text(separator='\n')
                        break # Found HTML, no need to look for plain text now
                    elif ctype == 'text/plain' and 'attachment' not in cdispo:
                        charset = part.get_content_charset()
                        plain_body = part.get_payload(decode=True).decode(charset if charset else 'utf-8', errors='ignore')
                        # Only set plain_body if HTML hasn't been found yet
                        if not body_content: # This ensures HTML preference
                            body_content = plain_body
            else: # Not multipart, assume single part message
                ctype = msg.get_content_type()
                if ctype == 'text/html':
                    charset = msg.get_content_charset()
                    html_body = msg.get_payload(decode=True).decode(charset if charset else 'utf-8', errors='ignore')
                    soup = BeautifulSoup(html_body, 'html.parser')
                    body_content = soup.get_text(separator='\n')
                elif ctype == 'text/plain':
                    charset = msg.get_content_charset()
                    plain_body = msg.get_payload(decode=True).decode(charset if charset else 'utf-8', errors='ignore')
                    body_content = plain_body

            if body_content:
                # Prepare text for ReportLab Paragraph:
                # 1. Replace ampersands, less-than, and greater-than signs to avoid XML parsing issues in ReportLab.
                # 2. Replace newlines with <br/> for proper line breaks in the PDF paragraph.
                formatted_body = body_content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                formatted_body = formatted_body.replace('\n', '<br/>\n')
                story.append(Paragraph(formatted_body, body_style))
            else:
                story.append(Paragraph("<i>No readable body content found for this email.</i>", normal_style))

            # Add a separator between emails for readability
            story.append(Spacer(1, 0.5 * inch)) # Space before divider
            story.append(Paragraph("=" * 100, normal_style)) # A visual divider line
            story.append(Spacer(1, 0.5 * inch)) # Space after divider

        except Exception as e:
            # Catch any errors during processing of a single file
            error_message = f"Error processing '{filename}': {e}"
            print(error_message)
            story.append(Paragraph(f"<b>{error_message}</b>", styles['h5']))
            story.append(Spacer(1, 0.5 * inch))

    # Build the PDF document if there's any content to add
    if story:
        try:
            doc.build(story)
            final_message = f"Successfully created '{pdf_file_path}' with content from {len(eml_files)} emails."
            print(final_message)
            messagebox.showinfo("Conversion Complete", final_message)
        except Exception as e:
            error_message = f"An error occurred while building the PDF document: {e}"
            print(error_message)
            messagebox.showerror("PDF Generation Error", error_message)
    else:
        messagebox.showinfo("No Content", "No content was added to the PDF. Check if .eml files exist and are readable.")
        print("No content was added to the PDF. Check if .eml files exist and are readable.")
    
    root.destroy() # Destroy the hidden root window when done

# --- How to Use ---
# 1. Save this code as a Python file (e.g., `eml_converter_gui.py`).
# 2. Ensure you have the required libraries installed:
#    `pip install reportlab beautifulsoup4`
# 3. Run the script from your terminal: `python eml_converter_gui.py`
#    A folder selection dialog will appear.
#    Select the folder containing your .eml files.
#    The merged PDF will be created in the same directory where you ran the script.

# Call the function to start the process
eml_to_pdf_batch_converter()

