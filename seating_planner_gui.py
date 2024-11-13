
import streamlit as st
import pandas as pd
import json
from datetime import datetime
import tempfile
import base64

from SeatingPlanner import  TableConstraints, SeatingHistory, SeatingOptimizer, Attendees_from_spreadsheet

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


def main():
    st.set_page_config(page_title="Seating Arrangement Planner", layout="wide")
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
                n_guests = len(attendees)
                n_tables = int(n_guests / table_size + 0.5)

                st.session_state.current_arrangement = st.session_state.optimizer.optimize_seating(
                    attendees,
                    num_tables=n_tables + 1
                )

            # Display current arrangement
            if st.session_state.current_arrangement:
                display_arrangement(st.session_state.current_arrangement)

                # Action buttons for current arrangement
                col1, col2, col3 = st.columns(3)

                with col1:
                    if st.button("Approve & Save"):
                        st.session_state.history.save_arrangement(st.session_state.current_arrangement)
                        st.success("Arrangement saved to history!")

                with col2:
                    if st.button("Export to Markdown"):
                        md_content = generate_markdown(st.session_state.current_arrangement)
                        st.download_button(
                            "Download Markdown",
                            md_content,
                            file_name="seating_arrangement.md",
                            mime="text/markdown"
                        )

                with col3:
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


if __name__ == "__main__":
    main()