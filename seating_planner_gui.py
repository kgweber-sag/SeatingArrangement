import streamlit as st
import pandas as pd
import io
from datetime import datetime
import tempfile
import base64

from SeatingPlanner import Attendees_from_spreadsheet, SeatingOptimizer, SeatingHistory, TableConstraints
def get_download_link(file_path, link_text):
    """Generate a download link for a file"""
    with open(file_path, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{file_path}">{link_text}</a>'


def generate_markdown(arrangement):
    """Generate markdown text for the seating arrangement"""
    md = ["# Seating Arrangement\n"]
    md.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    # Head table
    md.append("## Head Table\n")
    md.append(f"Number of guests: {len(arrangement[0])}\n")
    md.append("| Name | Gender | Seniority | Field | Assigned |")
    md.append("|------|--------|-----------|--------|----------|")
    for attendee in arrangement[0]:
        assigned = "✓" if attendee.assign_head_table else ""
        md.append(f"| {attendee.name} | {attendee.gender} | {attendee.seniority} | {attendee.field} | {assigned} |")

    # Other tables
    for i, table in enumerate(arrangement[1:], 2):
        md.append(f"\n## Table {i}\n")
        md.append(f"Number of guests: {len(table)}\n")
        md.append("| Name | Gender | Seniority | Field |")
        md.append("|------|--------|-----------|--------|")
        for attendee in table:
            md.append(f"| {attendee.name} | {attendee.gender} | {attendee.seniority} | {attendee.field} |")

    return "\n".join(md)


def display_arrangement(arrangement, show_metrics=True):
    """Display the seating arrangement in Streamlit"""
    # Head table
    st.subheader("Head Table")
    head_df = pd.DataFrame([
        {
            "Name": a.name,
            "Gender": a.gender,
            "Seniority": a.seniority,
            "Field": a.field,
            "Assigned": "✓" if a.assign_head_table else ""
        } for a in arrangement[0]
    ])
    st.dataframe(head_df, hide_index=True)

    # Other tables
    for i, table in enumerate(arrangement[1:], 2):
        st.subheader(f"Table {i}")
        table_df = pd.DataFrame([
            {
                "Name": a.name,
                "Gender": a.gender,
                "Seniority": a.seniority,
                "Field": a.field
            } for a in table
        ])
        st.dataframe(table_df, hide_index=True)


def show_excel_guide():
    """Display the guide for Excel file formatting"""
    st.header("Excel File Format Guide")

    st.write("""
    Your Excel file should contain the following columns:
    """)

    # Show example table
    example_data = {
        'name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
        'gender': ['M', 'F', 'M'],
        'seniority': ['senior', 'junior', 'senior'],
        'division': ['Sales', 'Engineering', 'Marketing'],
        'attending': ['Y', 'Y', 'Y'],
        'head table': [1.0, 0.0, 0.0]
    }
    example_df = pd.DataFrame(example_data)
    st.dataframe(example_df, hide_index=True)

    st.write("""
    ### Column Descriptions

    1. **name** (required)
        - Full name of the attendee
        - Text format

    2. **gender** (required)
        - 'M' or 'F'
        - Used for diversity balancing

    3. **seniority** (required)
        - 'senior' or 'junior'
        - Used for mixing experience levels

    4. **division** (required)
        - Department or team name
        - Used for cross-functional mixing

    5. **attending** (required)
        - 'Y' or 'N'
        - Only 'Y' entries will be included

    6. **head table** (required)
        - 1.0 for assigned head table seats
        - 0.0 for flexible seating

    ### Tips
    - Ensure all required columns are present
    - Check for consistent formatting (especially Y/N and 1.0/0.0)
    - Remove any blank rows
    - Save as .xlsx format
    """)

    # Add download link for template
    st.write("### Download Template")

    # Create template DataFrame
    template_df = pd.DataFrame({
        'name': ['Example Person'],
        'gender': ['F'],
        'seniority': ['senior'],
        'division': ['Department'],
        'attending': ['Y'],
        'head table': [0.0]
    })

    # Convert to Excel bytes
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False)

    st.download_button(
        label="Download Excel Template",
        data=buffer.getvalue(),
        file_name="seating_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def generate_excel(arrangement):
    """Generate a formatted Excel file with the seating arrangement"""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary sheet
        summary_data = {
            'Table': ['Head Table'] + [f'Table {i}' for i in range(2, len(arrangement) + 1)],
            'Number of Guests': [len(table) for table in arrangement]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

        # Head table sheet
        head_data = pd.DataFrame([
            {
                'Name': a.name,
                'Gender': a.gender,
                'Seniority': a.seniority,
                'Field': a.field,
                'Assigned Head Table': '✓' if a.assign_head_table else ''
            } for a in arrangement[0]
        ])
        head_data.to_excel(writer, sheet_name='Head Table', index=False)

        # Other tables
        for i, table in enumerate(arrangement[1:], 2):
            table_data = pd.DataFrame([
                {
                    'Name': a.name,
                    'Gender': a.gender,
                    'Seniority': a.seniority,
                    'Field': a.field
                } for a in table
            ])
            table_data.to_excel(writer, sheet_name=f'Table {i}', index=False)

    return output.getvalue()


def generate_word_doc(arrangement):
    """Generate a Word document with the seating arrangement"""
    from docx import Document
    from docx.shared import Inches

    doc = Document()
    doc.add_heading('Seating Arrangement', 0)
    doc.add_paragraph(f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')

    # Summary table
    doc.add_heading('Summary', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    header_cells = table.rows[0].cells
    header_cells[0].text = 'Table'
    header_cells[1].text = 'Number of Guests'

    for i, group in enumerate(arrangement):
        row_cells = table.add_row().cells
        row_cells[0].text = 'Head Table' if i == 0 else f'Table {i + 1}'
        row_cells[1].text = str(len(group))

    doc.add_paragraph()

    # Head table
    doc.add_heading('Head Table', level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    header_cells = table.rows[0].cells
    for i, header in enumerate(['Name', 'Gender', 'Seniority', 'Field', 'Assigned']):
        header_cells[i].text = header

    for attendee in arrangement[0]:
        row_cells = table.add_row().cells
        row_cells[0].text = attendee.name
        row_cells[1].text = attendee.gender
        row_cells[2].text = attendee.seniority
        row_cells[3].text = attendee.field
        row_cells[4].text = '✓' if attendee.assign_head_table else ''

    # Other tables
    for i, table_group in enumerate(arrangement[1:], 2):
        doc.add_paragraph()
        doc.add_heading(f'Table {i}', level=1)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        for i, header in enumerate(['Name', 'Gender', 'Seniority', 'Field']):
            header_cells[i].text = header

        for attendee in table_group:
            row_cells = table.add_row().cells
            row_cells[0].text = attendee.name
            row_cells[1].text = attendee.gender
            row_cells[2].text = attendee.seniority
            row_cells[3].text = attendee.field

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def generate_pdf(arrangement):
    """Generate a PDF with the seating arrangement"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    # Title
    elements.append(Paragraph('Seating Arrangement', styles['Title']))
    elements.append(Paragraph(f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', styles['Normal']))
    elements.append(Spacer(1, 12))

    # Summary
    elements.append(Paragraph('Summary', styles['Heading1']))
    summary_data = [['Table', 'Number of Guests']]
    for i, group in enumerate(arrangement):
        table_name = 'Head Table' if i == 0 else f'Table {i + 1}'
        summary_data.append([table_name, str(len(group))])

    summary_table = Table(summary_data)
    summary_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 20))

    # Head table
    elements.append(Paragraph('Head Table', styles['Heading1']))
    head_data = [['Name', 'Gender', 'Seniority', 'Field', 'Assigned']]
    for attendee in arrangement[0]:
        head_data.append([
            attendee.name,
            attendee.gender,
            attendee.seniority,
            attendee.field,
            '✓' if attendee.assign_head_table else ''
        ])

    head_table = Table(head_data)
    head_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(head_table)
    elements.append(Spacer(1, 20))

    # Other tables
    for i, table_group in enumerate(arrangement[1:], 2):
        elements.append(Paragraph(f'Table {i}', styles['Heading1']))
        table_data = [['Name', 'Gender', 'Seniority', 'Field']]
        for attendee in table_group:
            table_data.append([
                attendee.name,
                attendee.gender,
                attendee.seniority,
                attendee.field
            ])

        group_table = Table(table_data)
        group_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ]))
        elements.append(group_table)
        elements.append(Spacer(1, 20))

    doc.build(elements)
    return output.getvalue()



def main():
    st.set_page_config(page_title="Seating Arrangement Planner", layout="wide")

    # Create tabs
    tab1, tab2 = st.tabs(["Seating Planner", "File Format Guide"])

    with tab1:
        st.title("Seating Arrangement Planner")

        # Initialize session state
        if 'current_arrangement' not in st.session_state:
            st.session_state.current_arrangement = None
        if 'history' not in st.session_state:
            st.session_state.history = None
        if 'optimizer' not in st.session_state:
            st.session_state.optimizer = None

        # Sidebar for configuration
        with st.sidebar:
            st.header("Configuration")

            # File uploads
            excel_file = st.file_uploader("Upload Attendee Excel File", type=['xlsx'])
            history_file = st.file_uploader("Upload History File (Optional)", type=['json'])

            # Table size configuration
            table_size = st.number_input("Table Size", min_value=4, max_value=12, value=6)

            # Calculate and show recommended number of tables
            if excel_file:
                attendees = Attendees_from_spreadsheet(excel_file)
                n_guests = len(attendees)
                recommended_tables = int(n_guests / table_size + 0.5)

                st.write("### Table Count")
                st.write(f"Total Guests: {n_guests}")
                st.write(f"Recommended Tables: {recommended_tables}")

                # Allow user to override
                num_tables = st.number_input(
                    "Number of Tables (including head table)",
                    min_value=2,
                    value=recommended_tables + 1,  # +1 for head table
                    help="Default is calculated based on guests/table size. Adjust if needed."
                )

            # Action buttons
            generate_button = st.button("Generate New Arrangement")
            regenerate_button = st.button("Regenerate")

        # Main content area
        if excel_file:
            try:
                # Initialize optimizer if needed
                if st.session_state.optimizer is None or history_file:
                    constraints = TableConstraints(min_seats=4, max_seats=table_size)

                    # Handle history file
                    if history_file:
                        # Save uploaded history to temporary file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.json') as tmp_file:
                            tmp_file.write(history_file.getvalue())
                            history_path = tmp_file.name
                    else:
                        history_path = "seating_history.json"

                    st.session_state.history = SeatingHistory(filename=history_path)
                    st.session_state.optimizer = SeatingOptimizer(
                        constraints=constraints,
                        history=st.session_state.history
                    )

                # Generate new arrangement if requested
                if generate_button or regenerate_button or st.session_state.current_arrangement is None:
                    attendees = Attendees_from_spreadsheet(excel_file)

                    st.session_state.current_arrangement = st.session_state.optimizer.optimize_seating(
                        attendees,
                        num_tables=num_tables
                    )

                # Display current arrangement
                if st.session_state.current_arrangement:
                    display_arrangement(st.session_state.current_arrangement)

                    # First show the approval section
                    st.write("### Approve Arrangement")
                    if st.button("Approve & Save to History"):
                        st.session_state.history.save_arrangement(st.session_state.current_arrangement)
                        st.success("Arrangement saved to history!")

                    # Then show export options
                    st.write("### Export Options")
                    col1, col2, col3, col4 = st.columns(4)

                    with col1:
                        if st.button("Export to Excel"):
                            excel_bytes = generate_excel(st.session_state.current_arrangement)
                            st.download_button(
                                "Download Excel File",
                                excel_bytes,
                                file_name="seating_arrangement.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                    with col2:
                        if st.button("Export to Word"):
                            word_bytes = generate_word_doc(st.session_state.current_arrangement)
                            st.download_button(
                                "Download Word Document",
                                word_bytes,
                                file_name="seating_arrangement.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

                    with col3:
                        if st.button("Export to PDF"):
                            pdf_bytes = generate_pdf(st.session_state.current_arrangement)
                            st.download_button(
                                "Download PDF",
                                pdf_bytes,
                                file_name="seating_arrangement.pdf",
                                mime="application/pdf"
                            )

                    with col4:
                        if st.session_state.history and st.session_state.history.filename:
                            st.download_button(
                                "Download History",
                                open(st.session_state.history.filename, 'rb'),
                                file_name="seating_history.json",
                                mime="application/json"
                            )

            except Exception as e:
                st.error(f"Error: {str(e)}")
        else:
            st.info("Please upload an Excel file with attendee data to begin.")

    with tab2:
        show_excel_guide()


if __name__ == "__main__":
    main()