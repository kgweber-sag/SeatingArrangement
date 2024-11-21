# Seating Arrangement Planner

This project is a Seating Arrangement Planner application built with Python and Streamlit. It allows users to upload
attendee lists, configure table sizes, and generate optimized seating arrangements. The application also supports
exporting the seating plans to Word and PDF formats.

## Features

- Upload attendee Excel files
- Configure table sizes and diversity weights
- Generate optimized seating arrangements
- Export seating plans to Word and PDF
- Maintain seating history

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/kgweber-sag/SeatingArrangement.git
    cd SeatingArrangement
    ```

2. Create a virtual environment and activate it:
    ```sh
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. Run the Streamlit application:
    ```sh
    streamlit run seating_planner_gui.py
    ```
2. Open your web browser and go to `http://localhost:8501`.
3. Upload the attendee Excel file and configure the event date, table sizes and diversity weights.
4. Generate the seating arrangement
5. Regenerate as many times as you wish
6. Approve and download a new copy of the history with the proposed seating configuration appended.
7. Optionally, export the seating plan to Word or PDF.

## Project Structure

- `seating_planner_gui.py`: Main application file.
- `seating_planner.py`: Seating arrangement planner module.
- `requirements.txt`: List of dependencies.
- `example_attendee_list.xlsx`: Example attendee list.
- `README.md`: Project documentation.

## Dependencies

- pandas==2.2.2
- openpyxl==3.1.5
- streamlit==1.38.0
- python-docx==1.1.2
- reportlab
- pytest

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.