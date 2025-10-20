# Microwave DAQ Testing System

Author: ziad emad allam

## Overview
This project is a professional automated testing system for microwave ovens. It uses Python, PyQt5 for the GUI, and DAQ hardware for real signal acquisition. The system is designed for industrial-level analysis, including sector-based Defrost mode, and generates detailed Excel reports.

## File Structure & Deep Explanation
- `main.py`: The main entry point and GUI controller. It:
  - Initializes the application and main window.
  - Manages user interaction, including starting/stopping tests, saving data, and displaying real-time signal status.
  - Integrates the state machine to control microwave operation modes.
  - Handles event logging, warnings, and updates all UI elements (icons, graphs, statistics).
  - Calls DAQHandler for data acquisition and analysis, and ExcelWriter for report generation.
  - Implements DefrostDialog for sector calculation and user input in Defrost mode.

- `daq_handler.py`: The DAQ logic and analysis engine. It:
  - Connects to the DAQ device and reads all signal channels (Microwave, Grill, Lamp, Door SW, Buzzer).
  - Buffers and timestamps all acquired data for later analysis.
  - Calculates power percentages for each signal using configurable thresholds and sample windows.
  - Implements sector-based Defrost logic: calculates sector timings and expected power for each sector based on weight, and analyzes Pass/Fail for each sector individually.
  - Provides industrial Pass/Fail analysis for all modes, including detailed results and warnings (e.g., door opens, MW+Grill overlap).
  - Exports all recorded data for reporting.

- `config.py`: Central configuration and constants. It:
  - Defines device/channel mappings, voltage thresholds, and signal names.
  - Sets all tolerances, Defrost sector definitions (percentages, expected power, timings), and Excel columns.
  - Contains all icon symbols and theme colors for the GUI.

- `excel_writer.py`: Excel report generator. It:
  - Creates and manages Excel workbooks.
  - Writes all raw test data, summary statistics, and Pass/Fail results (including sector results for Defrost mode).
  - Ensures reports are formatted and saved correctly for industrial documentation.

- `state_machine.py`: Microwave state machine. It:
  - Defines all operational states (Idle, Running, Paused, etc.) and transitions.
  - Ensures correct sequencing and logic for each microwave mode, including Defrost and Child Lock.

- `__pycache__/`: Python bytecode cache. Not user-editable, auto-generated for faster execution.
- `.venv/`: Python virtual environment. Contains all installed libraries for this project only.

## How to Run
1. **Environment Setup**
   - Install Python 3.8+.
   - Create and activate a virtual environment:
     ```powershell
     python -m venv .venv
     .venv\Scripts\activate
     ```
2. **Install Required Libraries**
   - Install dependencies:
     ```powershell
     pip install -r requirements.txt
     ```
3. **Connect DAQ Hardware**
   - Connect and configure your DAQ device as specified in `config.py`.
4. **Run the Application**
   - Start the GUI:
     ```powershell
     python main.py
     ```

## Required Libraries
- PyQt5: For GUI components and dialogs.
- nidaqmx: For DAQ device communication and data acquisition.
- openpyxl: For Excel report generation and saving.
- pyqtgraph: For real-time signal plotting and visualization.
- numpy: For efficient data handling and calculations.

All libraries are listed in `requirements.txt`.

## Features (Deep)
- **Real-Time Signal Monitoring**: All five signals are displayed live with icons, graphs, and statistics. Signal status is updated every second.
- **Industrial Pass/Fail Analysis**: All modes are analyzed using configurable tolerances. Defrost mode uses sector-based analysis, checking each sector's measured power against expected values.
- **Defrost Sector Logic**: Defrost mode divides the test into three sectors, each with its own timing and expected power. The system calculates sector boundaries based on sample weight and analyzes each sector individually for Pass/Fail.
- **Automated Excel Reporting**: All test data, statistics, and results (including sector results) are saved in a formatted Excel file for documentation and review.
- **Professional GUI**: Modern design, event logging, warnings, and user-friendly controls. All thresholds, icons, and colors are configurable in `config.py`.
- **State Machine Control**: Ensures correct operation and transitions for all microwave modes, including safety features like Child Lock and overlap detection.

## Notes
- Code is modular, clean, and documented for easy maintenance and extension.
- For troubleshooting, check DAQ connection, library installation, and configuration in `config.py`.
- The system is designed for industrial and laboratory environments, with a focus on reliability and accuracy.

---
Project by ziad emad allam
