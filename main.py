"""
Main GUI Application for Microwave DAQ Testing System
Full Featured - Optimized for 15.6" Laptop - Single Page Layout
Door SW REVERSED: 0V=Closed, 5V=Open
With Complete Pass/Fail Analysis & All Features
"""

import sys
import os
import subprocess
import platform
import traceback
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QGroupBox, 
                             QGridLayout, QFileDialog, QMessageBox, QProgressBar,
                             QDialog, QLineEdit, QTextEdit, QComboBox, QFrame,
                             QSpacerItem, QSizePolicy, QScrollArea, QSplitter)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont, QPalette, QColor
import pyqtgraph as pg
from datetime import datetime
import numpy as np
import logging

import config
from daq_handler import DAQHandler
from excel_writer import ExcelWriter
from state_machine import MicrowaveStateMachine, MicrowaveState


# ==================== MODERN UI COMPONENTS ====================

class ModernButton(QPushButton):
    """Modern styled button"""
    def __init__(self, text, icon="", color="primary", parent=None):
        super().__init__(parent)
        self.setText(text)
        self.setObjectName(f"btn_{color}")
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(35)
        self.setFont(QFont("Segoe UI", 9))


class StatusIndicator(QLabel):
    """Status indicator"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setFixedSize(10, 10)
    
    def set_status(self, active=False):
        if active:
            self.setStyleSheet("background-color: #4CAF50; border-radius: 5px;")
        else:
            self.setStyleSheet("background-color: #757575; border-radius: 5px;")


# ==================== DEFROST DIALOG ====================

class DefrostDialog(QDialog):
    """Dialog for Defrost mode weight input with sector calculation"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Defrost Test Configuration")
        self.setModal(True)
        self.setFixedSize(450, 400)
        
        self.weight = None
        self.sectors = None
        
        self._setup_ui()
        self._apply_style()
    
    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)
        
        # Title
        title = QLabel("Defrost Test Setup")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Weight input section
        input_group = QGroupBox("Weight Configuration")
        input_layout = QGridLayout()
        input_layout.setSpacing(10)
        
        weight_label = QLabel("Enter Weight (gm):")
        weight_label.setFont(QFont("Segoe UI", 10))
        input_layout.addWidget(weight_label, 0, 0)
        
        self.weight_input = QLineEdit()
        self.weight_input.setPlaceholderText("100-2000 gm")
        self.weight_input.setFont(QFont("Segoe UI", 11))
        self.weight_input.setMinimumHeight(35)
        input_layout.addWidget(self.weight_input, 0, 1)
        
        range_label = QLabel("Range: 100-2000 gm, Step: 100 gm")
        range_label.setFont(QFont("Segoe UI", 8))
        range_label.setStyleSheet("color: #B0B0B0;")
        input_layout.addWidget(range_label, 1, 0, 1, 2)
        
        self.calc_button = ModernButton("Calculate Sectors", "", "primary")
        self.calc_button.clicked.connect(self.calculate_sectors)
        input_layout.addWidget(self.calc_button, 2, 0, 1, 2)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # Results section
        results_group = QGroupBox("Calculated Sectors")
        results_layout = QVBoxLayout()
        results_layout.setSpacing(5)
        
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setMinimumHeight(150)
        self.results_text.setFont(QFont("Consolas", 9))
        results_layout.addWidget(self.results_text)
        
        results_group.setLayout(results_layout)
        layout.addWidget(results_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        self.ok_button = ModernButton("Start Recording", "", "success")
        self.ok_button.setEnabled(False)
        self.ok_button.clicked.connect(self.accept)
        
        self.cancel_button = ModernButton("Cancel", "", "danger")
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def calculate_sectors(self):
        try:
            weight = int(self.weight_input.text())
            
            if weight < 100 or weight > 2000:
                QMessageBox.warning(self, "Invalid Weight", "Weight must be between 100-2000 gm")
                return
            
            # Create temporary DAQ handler for calculation
            temp_handler = DAQHandler()
            sectors, time_msg = temp_handler.calculate_defrost_sectors(weight)
            
            if sectors is None:
                QMessageBox.warning(self, "Error", time_msg)
                return
            
            self.weight = weight
            self.sectors = sectors
            
            # Display results
            result_text = f"Weight: {weight} gm\n"
            result_text += f"{time_msg}\n"
            result_text += f"{'='*45}\n\n"
            
            for i, sector in enumerate(sectors, 1):
                start_min = int(sector['start_time'] // 60)
                start_sec = int(sector['start_time'] % 60)
                end_min = int(sector['end_time'] // 60)
                end_sec = int(sector['end_time'] % 60)
                
                result_text += f"{sector['name']}:\n"
                result_text += f"  Time Range: {start_min:02d}:{start_sec:02d} - {end_min:02d}:{end_sec:02d}\n"
                result_text += f"  Duration: {sector['duration']:.1f} seconds\n"
                result_text += f"  Expected Power: {sector['expected_power']}%\n"
                result_text += f"  Pattern: ON={sector['on_time']}s, OFF={sector['off_time']}s\n"
                result_text += f"  Period: {sector['period']}s\n\n"
            
            self.results_text.setText(result_text)
            self.ok_button.setEnabled(True)
            
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number!")
    
    def _apply_style(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #1a1a1a;
            }
            QLabel {
                color: #FFFFFF;
            }
            QGroupBox {
                background-color: #252525;
                border: 1px solid #3a3a3a;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 15px;
                font-weight: bold;
                color: #FFFFFF;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QLineEdit {
                background-color: #2d2d2d;
                color: #FFFFFF;
                border: 2px solid #4a4a4a;
                border-radius: 6px;
                padding: 8px;
            }
            QLineEdit:focus {
                border-color: #4CAF50;
            }
            QTextEdit {
                background-color: #1e1e1e;
                color: #4FC3F7;
                border: 1px solid #3a3a3a;
                border-radius: 6px;
                padding: 10px;
            }
        """)


# ==================== MODE CONFIGURATION DATA ====================

MODE_CONFIGS = {
    "-- Select Test Mode --": {
        "type": "none",
        "description": "Please select a test mode to begin",
        "details": ""
    },
    "Manual: Microwave": {
        "type": "manual_mw",
        "description": "Manual microwave test with adjustable power",
        "details": "Select power level from 10P to 100P. Power is controlled via PWM duty cycle.",
        "requires": ["power"],
        "power_range": (10, 100, 10),
        "expected_power": "variable"
    },
    "Manual: Grill": {
        "type": "manual_grill",
        "description": "Manual grill test at 100% power",
        "details": "Grill operates at full power (100%). No power adjustment available.",
        "requires": [],
        "expected_power": 100
    },
    "Combination: C1 (20% MW / 80% Grill)": {
        "type": "combination",
        "mode": "C1",
        "description": "Combination mode C1",
        "details": "20% MW + 80% Grill (alternating, no overlap). Suitable for liquid vegetables.",
        "requires": [],
        "expected_mw": 20,
        "expected_grill": 80
    },
    "Combination: C2 (40% MW / 60% Grill)": {
        "type": "combination",
        "mode": "C2",
        "description": "Combination mode C2",
        "details": "40% MW + 60% Grill (alternating, no overlap). Suitable for warming.",
        "requires": [],
        "expected_mw": 40,
        "expected_grill": 60
    },
    "Defrost": {
        "type": "defrost",
        "description": "Defrost mode with 3 power sectors",
        "details": "Weight-based defrost with 3 sectors at different power levels (36.7%, 23.3%, 30%).",
        "requires": ["weight"],
        "weight_range": (100, 2000, 100),
        "sectors": 3
    },
    "Auto Menu: Popcorn": {
        "type": "auto",
        "menu": "popcorn",
        "description": "Auto cook - Popcorn",
        "details": "100% MW power. Time calculated based on weight (50-150 gm).",
        "requires": ["weight"],
        "weight_range": (50, 150, 50),
        "expected_power": 100
    },
    "Auto Menu: Meat": {
        "type": "auto",
        "menu": "meat",
        "description": "Auto cook - Meat",
        "details": "100% MW power. Supports 100-1000 gm with 3 weight ranges.",
        "requires": ["weight"],
        "weight_range": (100, 1000, 50),
        "expected_power": 100
    },
    "Auto Menu: Pizza": {
        "type": "auto",
        "menu": "pizza",
        "description": "Auto cook - Pizza",
        "details": "100% MW power. Optimized for 100-900 gm pizza.",
        "requires": ["weight"],
        "weight_range": (100, 900, 50),
        "expected_power": 100
    },
    "Auto Menu: Chicken": {
        "type": "auto",
        "menu": "chicken",
        "description": "Auto cook - Chicken",
        "details": "53% MW + 47% Grill (alternating). Includes midtime pause for turning food.",
        "requires": ["weight"],
        "weight_range": (50, 1500, 50),
        "expected_mw": 53,
        "expected_grill": 47,
        "special": "midtime_pause"
    },
    "Auto Menu: Rice": {
        "type": "auto",
        "menu": "rice",
        "description": "Auto cook - Rice",
        "details": "100% MW power. Supports 100-800 gm (0.5-4 cups).",
        "requires": ["weight"],
        "weight_range": (100, 800, 50),
        "expected_power": 100
    },
    "Auto Menu: Beverages": {
        "type": "auto",
        "menu": "beverages",
        "description": "Auto cook - Beverages",
        "details": "100% MW power. Measured in ml (150-600 ml).",
        "requires": ["weight"],
        "weight_range": (150, 600, 150),
        "unit": "ml",
        "expected_power": 100
    },
    "Auto Menu: Pasta": {
        "type": "auto",
        "menu": "pasta",
        "description": "Auto cook - Pasta",
        "details": "80% MW power. Optimized for 50-350 gm pasta.",
        "requires": ["weight"],
        "weight_range": (50, 350, 50),
        "expected_power": 80
    },
    "Auto Menu: Fish": {
        "type": "auto",
        "menu": "fish",
        "description": "Auto cook - Fish",
        "details": "77% MW power. Supports 200-1000 gm with 100 gm steps.",
        "requires": ["weight"],
        "weight_range": (200, 1000, 100),
        "expected_power": 77
    }
    ,
    "Normal": {
        "type": "normal",
        "description": "Normal mode: calculates idle/silence duration.",
        "details": "Tracks how long the system remains idle (OFF) during the test.",
        "requires": []
    }
}


# ==================== MAIN WINDOW ====================

class MainWindow(QMainWindow):
    def update_signal_icons(self):
        """Update icons state (solid/blink/off) and always show the correct icon for each channel (Unicode icons)"""
        channel_icons = {
            'Microwave': '⬛',
            'Grill': '♨',
            'Lamp': '💡',
            'Door SW': '⎔',
            'Buzzer': '🔊'
        }
        for channel, widget in self.signal_widgets.items():
            icon = channel_icons.get(channel, "")
            widget['icon'].setText(icon)
            widget['icon'].setFont(QFont("Segoe UI", 28, QFont.Bold))
            status = widget['status'].text()
            # Color logic
            if status == "ON":
                widget['icon'].setStyleSheet("color: #4CAF50;")
            elif status == "BEEP":
                widget['icon'].setStyleSheet("color: #FF9800;")
            elif status == "OFF":
                widget['icon'].setStyleSheet("color: #757575;")
            else:
                widget['icon'].setStyleSheet("color: #757575;")
    """Main Application Window - Full Featured for 15.6 inch laptop"""
    def __init__(self):
        super().__init__()
        
        self.daq = DAQHandler()
        self.excel_writer = None
        
        self.current_mode = None
        self.current_config = None
        
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_display)
        
        self.state_machine = MicrowaveStateMachine(sleep_timeout=900)  # 15 min
        
        logging.basicConfig(
            filename='microwave_events.log',
            level=logging.INFO,
            format='%(asctime)s %(levelname)s: %(message)s'
        )
        self.last_logged_state = None
        
        # Child Lock variables
        self.child_lock_active = False
        self.child_lock_timer = None
        
        self._setup_ui()
        self._apply_modern_theme()
        self._connect_daq()
    
    def keyPressEvent(self, event):
        # Child Lock: Detect Start+Cancel combo
        if event.key() == Qt.Key_S and event.modifiers() & Qt.ControlModifier:
            # Ctrl+S simulates Start+Cancel for demo
            if not self.child_lock_active:
                self._activate_child_lock()
            else:
                self._deactivate_child_lock()

    def play_beep(self, count=1, long=False):
        """Play beep sound according to spec (count, long beep)"""
        for _ in range(count):
            QApplication.beep()
        if long:
            # Simulate long beep by holding sound (not natively supported)
            QApplication.beep()
            QTimer.singleShot(700, QApplication.beep)

    def _activate_child_lock(self):
        self.child_lock_active = True
        self.rec_status_label.setText("Child Lock Active")
        self.rec_status_label.setStyleSheet("color: #FFD600;")
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.mode_selector.setEnabled(False)
        QApplication.beep()
        QApplication.beep()
        self.play_beep(count=2)
        self.warnings_text.append("[Child Lock] Activated. All controls disabled.")
        # Optionally show lock icon

    def _deactivate_child_lock(self):
        self.child_lock_active = False
        self.rec_status_label.setText("Ready")
        self.rec_status_label.setStyleSheet("color: #4CAF50;")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(True)
        self.save_button.setEnabled(True)
        self.mode_selector.setEnabled(True)
        QApplication.beep()
        self.play_beep(long=True)
        self.warnings_text.append("[Child Lock] Deactivated. Controls enabled.")
        # Optionally hide lock icon
    
    def _setup_ui(self):
        """Setup full-featured compact UI for 15.6 inch laptop"""
        self.setWindowTitle("Microwave DAQ Testing System - Tornado")
        self.setGeometry(50, 50, 1400, 900)
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setSpacing(6)
        main_layout.setContentsMargins(8, 8, 8, 8)
        central_widget.setLayout(main_layout)
        
        # Header
        header = self._create_header()
        main_layout.addWidget(header)
        
        # Top section: Mode + Control
        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        
        # Left: Mode selection
        mode_section = self._create_mode_section()
        top_row.addWidget(mode_section, 2)
        
        # Right: Control panel
        control_section = self._create_control_section()
        top_row.addWidget(control_section, 3)
        
        main_layout.addLayout(top_row)
        
        # Middle: Signals + Graphs
        middle_splitter = QSplitter(Qt.Horizontal)
        
        # Left: Signals
        signals_widget = self._create_signals_section()
        middle_splitter.addWidget(signals_widget)
        
        # Right: Graphs
        graphs_widget = self._create_graphs_section()
        middle_splitter.addWidget(graphs_widget)
        
        middle_splitter.setStretchFactor(0, 1)
        middle_splitter.setStretchFactor(1, 2)
        
        main_layout.addWidget(middle_splitter, 1)
        
        # Bottom: Warnings + Stats
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(8)
        
        warnings_widget = self._create_warnings_section()
        bottom_row.addWidget(warnings_widget, 2)
        
        stats_widget = self._create_stats_section()
        bottom_row.addWidget(stats_widget, 1)
        
        main_layout.addLayout(bottom_row)
    
    def _create_header(self):
        """Create compact header"""
        header = QFrame()
        header.setObjectName("header")
        header.setMaximumHeight(45)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(12, 5, 12, 5)
        
        title = QLabel("Microwave DAQ Testing System - Tornado")
        title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        layout.addWidget(title)
        
        layout.addStretch()
        
        # Current State Display
        self.state_label = QLabel("State: IDLE")
        self.state_label.setFont(QFont("Segoe UI", 9, QFont.Bold))
        self.state_label.setStyleSheet("color: #FFD600; padding: 4px;")
        layout.addWidget(self.state_label)
        
        layout.addSpacing(15)
        
        # DAQ Status
        self.daq_indicator = StatusIndicator()
        layout.addWidget(self.daq_indicator)
        self.daq_status_label = QLabel("DAQ")
        self.daq_status_label.setFont(QFont("Segoe UI", 8))
        layout.addWidget(self.daq_status_label)
        
        layout.addSpacing(15)
        
        # Recording Status
        self.rec_indicator = StatusIndicator()
        layout.addWidget(self.rec_indicator)
        self.rec_status_label = QLabel("Ready")
        self.rec_status_label.setFont(QFont("Segoe UI", 8))
        layout.addWidget(self.rec_status_label)
        
        header.setLayout(layout)
        return header
    
    def _create_mode_section(self):
        """Create mode selection section"""
        group = QGroupBox("Test Mode")
        layout = QVBoxLayout()
        layout.setSpacing(6)
        
        # Mode selector
        self.mode_selector = QComboBox()
        self.mode_selector.setMinimumHeight(35)
        self.mode_selector.setFont(QFont("Segoe UI", 9))
        for mode_name in MODE_CONFIGS.keys():
            self.mode_selector.addItem(mode_name)
        self.mode_selector.currentTextChanged.connect(self._on_mode_changed)
        layout.addWidget(self.mode_selector)
        
        # Mode description
        self.mode_desc_label = QLabel("Select a test mode")
        self.mode_desc_label.setWordWrap(True)
        self.mode_desc_label.setFont(QFont("Segoe UI", 8))
        self.mode_desc_label.setStyleSheet("color: #B0B0B0; padding: 5px;")
        self.mode_desc_label.setMaximumHeight(40)
        layout.addWidget(self.mode_desc_label)
        
        # Config inputs
        self.config_widget = QWidget()
        self.config_layout = QVBoxLayout()
        self.config_layout.setContentsMargins(0, 0, 0, 0)
        self.config_layout.setSpacing(4)
        self.config_widget.setLayout(self.config_layout)
        layout.addWidget(self.config_widget)
        
        # Expected results
        self.expected_label = QLabel("")
        self.expected_label.setWordWrap(True)
        self.expected_label.setFont(QFont("Segoe UI", 8))
        self.expected_label.setStyleSheet("background: #1e3a1e; color: #66BB6A; padding: 6px; border-radius: 4px;")
        self.expected_label.setMaximumHeight(50)
        self.expected_label.hide()
        layout.addWidget(self.expected_label)
        
        layout.addStretch()
        group.setLayout(layout)
        return group
    
    def _create_control_section(self):
        """Create control section"""
        group = QGroupBox("Control Panel")
        layout = QVBoxLayout()
        layout.setSpacing(6)
        
        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(6)
        
        self.start_button = ModernButton("Start", "", "success")
        self.start_button.clicked.connect(self.start_recording)
        btn_layout.addWidget(self.start_button)
        
        self.stop_button = ModernButton("Stop", "", "danger")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_recording)
        btn_layout.addWidget(self.stop_button)
        
        self.save_button = ModernButton("Save", "", "primary")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_data)
        btn_layout.addWidget(self.save_button)
        
        layout.addLayout(btn_layout)
        
        # Info display
        info_grid = QGridLayout()
        info_grid.setSpacing(6)
        
        # Duration
        dur_lbl = QLabel("Duration:")
        dur_lbl.setFont(QFont("Segoe UI", 8, QFont.Bold))
        info_grid.addWidget(dur_lbl, 0, 0)
        
        self.duration_display = QLabel("00:00:00")
        self.duration_display.setFont(QFont("Consolas", 11, QFont.Bold))
        self.duration_display.setStyleSheet("color: #4FC3F7;")
        info_grid.addWidget(self.duration_display, 0, 1)
        
        # Samples
        samp_lbl = QLabel("Samples:")
        samp_lbl.setFont(QFont("Segoe UI", 8, QFont.Bold))
        info_grid.addWidget(samp_lbl, 0, 2)
        
        self.samples_display = QLabel("0")
        self.samples_display.setFont(QFont("Consolas", 11, QFont.Bold))
        self.samples_display.setStyleSheet("color: #4FC3F7;")
        info_grid.addWidget(self.samples_display, 0, 3)
        
        # Result
        result_lbl = QLabel("Test Result:")
        result_lbl.setFont(QFont("Segoe UI", 8, QFont.Bold))
        info_grid.addWidget(result_lbl, 1, 0)
        
        self.result_display = QLabel("N/A")
        self.result_display.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.result_display.setStyleSheet("color: #909090;")
        info_grid.addWidget(self.result_display, 1, 1, 1, 3)
        
        layout.addLayout(info_grid)
        
        group.setLayout(layout)
        return group
    
    def _create_signals_section(self):
        """Create signals section"""
        group = QGroupBox("Real-time Signals")
        layout = QGridLayout()
        layout.setSpacing(6)
        layout.setContentsMargins(6, 6, 6, 6)
        
        self.signal_widgets = {}
        
        # Icon mapping for each channel (real-world inspired)
        channel_icons = {
            'Microwave': '⬛',   # Black square (device)
            'Grill': '♨',       # Hot springs (grill/heat)
            'Lamp': '💡',        # Light bulb
            'Door SW': '⎔',     # Open door/tech symbol
            'Buzzer': '🔊'       # Speaker
        }
        
        row = 0
        for channel in config.CHANNELS.keys():
            # Icon
            icon = QLabel(channel_icons.get(channel, ""))
            icon.setFont(QFont("Arial", 18))
            layout.addWidget(icon, row, 0)

            # Name
            name = QLabel(channel)
            name.setFont(QFont("Segoe UI", 8, QFont.Bold))
            layout.addWidget(name, row, 1)

            # Voltage
            voltage = QLabel("0.00V")
            voltage.setFont(QFont("Consolas", 9))
            voltage.setStyleSheet("color: #4FC3F7;")
            layout.addWidget(voltage, row, 2)

            # Status
            status = QLabel("OFF")
            status.setFont(QFont("Segoe UI", 8))
            layout.addWidget(status, row, 3)

            # Power for MW/Grill
            if channel in ['Microwave', 'Grill']:
                power = QLabel("0%")
                power.setFont(QFont("Segoe UI", 8, QFont.Bold))
                power.setStyleSheet("color: #66BB6A;")
                layout.addWidget(power, row, 4)

                progress = QProgressBar()
                progress.setMaximum(100)
                progress.setTextVisible(False)
                progress.setMaximumHeight(4)
                layout.addWidget(progress, row + 1, 0, 1, 5)
                self.signal_widgets[channel] = {
                    'icon': icon, 'voltage': voltage, 'status': status,
                    'power': power, 'progress': progress
                }
                row += 1
            else:
                self.signal_widgets[channel] = {
                    'icon': icon, 'voltage': voltage, 'status': status
                }

            row += 1
        
        group.setLayout(layout)
        return group
    
    def _create_graphs_section(self):
        """Create all 5 graphs section"""
        group = QGroupBox("Live Graphs (Last 60 seconds)")
        layout = QVBoxLayout()
        layout.setSpacing(3)
        layout.setContentsMargins(4, 4, 4, 4)
        
        self.graph_widgets = {}
        self.graph_curves = {}
        self.graph_data_x = {}
        self.graph_data_y = {}
        
        for channel in config.CHANNELS.keys():
            plot_widget = pg.PlotWidget()
            plot_widget.setBackground('#1e1e1e')
            plot_widget.setMaximumHeight(95)
            plot_widget.setLabel('left', 'V', **{'color': '#FFF', 'font-size': '7pt'})
            plot_widget.setYRange(config.VOLTAGE_MIN, config.VOLTAGE_MAX)
            plot_widget.setTitle(f"<span style='color: #FFF; font-size: 8pt; font-weight: bold'>{channel}</span>")
            plot_widget.showGrid(x=True, y=True, alpha=0.2)
            
            color = config.GRAPH_COLORS.get(channel, '#FFFFFF')
            curve = plot_widget.plot(pen=pg.mkPen(color, width=1.5))
            
            self.graph_widgets[channel] = plot_widget
            self.graph_curves[channel] = curve
            self.graph_data_x[channel] = []
            self.graph_data_y[channel] = []
            
            layout.addWidget(plot_widget)
        
        group.setLayout(layout)
        return group
    
    def _create_warnings_section(self):
        """Create warnings section"""
        group = QGroupBox("Warnings & Alerts")
        layout = QVBoxLayout()
        layout.setContentsMargins(4, 4, 4, 4)
        
        self.warnings_text = QTextEdit()
        self.warnings_text.setReadOnly(True)
        self.warnings_text.setMaximumHeight(55)
        self.warnings_text.setFont(QFont("Consolas", 8))
        layout.addWidget(self.warnings_text)
        
        group.setLayout(layout)
        return group
    
    def _create_stats_section(self):
        """Create statistics section"""
        group = QGroupBox("Statistics")
        layout = QGridLayout()
        layout.setSpacing(4)
        layout.setContentsMargins(6, 6, 6, 6)
        
        self.stats_widgets = {}
        
        stats_data = [
            ('mw_power', 'MW Power:', '0%', 0, 0),
            ('grill_power', 'Grill Power:', '0%', 0, 1),
            ('door_opens', 'Door Opens:', '0', 1, 0),
            ('mw_expected', 'Expected MW:', 'N/A', 1, 1),
        ]
        
        for key, label, default, row, col in stats_data:
            lbl = QLabel(label)
            lbl.setFont(QFont("Segoe UI", 7))
            layout.addWidget(lbl, row * 2, col)
            
            val = QLabel(default)
            val.setFont(QFont("Consolas", 9, QFont.Bold))
            val.setStyleSheet("color: #4FC3F7;")
            layout.addWidget(val, row * 2 + 1, col)
            
            self.stats_widgets[key] = val
        
        group.setLayout(layout)
        return group
    
    def _on_mode_changed(self, mode_name):
        """Handle mode change"""
        self.current_mode = mode_name
        self.current_config = MODE_CONFIGS.get(mode_name, {})
        
        # Update description
        desc = self.current_config.get('description', '')
        details = self.current_config.get('details', '')
        self.mode_desc_label.setText(f"{desc}\n{details}")
        
        # Clear config
        while self.config_layout.count():
            child = self.config_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        requires = self.current_config.get('requires', [])
        
        # Update expected results
        self._update_expected_display()
        
        if not requires:
            return
        
        # Weight input
        if 'weight' in requires:
            lbl = QLabel("Weight/Amount:")
            lbl.setFont(QFont("Segoe UI", 8))
            self.config_layout.addWidget(lbl)
            
            self.weight_input = QLineEdit()
            unit = self.current_config.get('unit', 'gm')
            w_range = self.current_config.get('weight_range', (0, 1000))
            self.weight_input.setPlaceholderText(f"{w_range[0]}-{w_range[1]} {unit}")
            self.weight_input.setMinimumHeight(28)
            self.config_layout.addWidget(self.weight_input)
        
        # Power input
        if 'power' in requires:
            lbl = QLabel("Power Level:")
            lbl.setFont(QFont("Segoe UI", 8))
            self.config_layout.addWidget(lbl)
            
            self.power_selector = QComboBox()
            self.power_selector.setMinimumHeight(28)
            
            power_range = self.current_config.get('power_range', (10, 100, 10))
            for p in range(power_range[0], power_range[1] + 1, power_range[2]):
                self.power_selector.addItem(f"{p}P")
            
            self.config_layout.addWidget(self.power_selector)
    
    def _update_expected_display(self):
        """Update expected results display"""
        if not self.current_config or self.current_config.get('type') == 'none':
            self.expected_label.hide()
            return
        
        text = "Expected: "
        
        if 'expected_power' in self.current_config:
            power = self.current_config['expected_power']
            if power != "variable":
                text += f"MW={power}%"
        
        if 'expected_mw' in self.current_config:
            mw = self.current_config['expected_mw']
            grill = self.current_config.get('expected_grill', 0)
            text += f"MW={mw}%, Grill={grill}%"
        
        if text != "Expected: ":
            self.expected_label.setText(text)
            self.expected_label.show()
        else:
            self.expected_label.hide()
    
    def _apply_modern_theme(self):
        """Apply dark professional theme"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1a1a1a;
            }
            QWidget {
                color: #FFFFFF;
            }
            #header {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #1e3a5f, stop:1 #2a5298);
                border-radius: 6px;
            }
            QGroupBox {
                background-color: #252525;
                border: 1px solid #3a3a3a;
                border-radius: 6px;
                margin-top: 6px;
                font-weight: bold;
                padding-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
            }
            QComboBox {
                background-color: #2d2d2d;
                color: #FFFFFF;
                border: 2px solid #4a4a4a;
                border-radius: 5px;
                padding: 4px;
            }
            QComboBox:hover {
                border-color: #2196F3;
            }
            QComboBox QAbstractItemView {
                background-color: #2d2d2d;
                color: #FFFFFF;
                selection-background-color: #2196F3;
            }
            QLineEdit {
                background-color: #2d2d2d;
                color: #FFFFFF;
                border: 2px solid #4a4a4a;
                border-radius: 5px;
                padding: 4px;
            }
            QLineEdit:focus {
                border-color: #4CAF50;
            }
            #btn_primary {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2196F3, stop:1 #1976D2);
                color: white;
                border: none;
                border-radius: 5px;
                font-weight: bold;
            }
            #btn_primary:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #42A5F5, stop:1 #2196F3);
            }
            #btn_primary:disabled {
                background: #424242;
                color: #757575;
            }
            #btn_success {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #4CAF50, stop:1 #388E3C);
                color: white;
                border: none;
                border-radius: 5px;
                font-weight: bold;
            }
            #btn_success:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #66BB6A, stop:1 #4CAF50);
            }
            #btn_success:disabled {
                background: #424242;
                color: #757575;
            }
            #btn_danger {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f44336, stop:1 #D32F2F);
                color: white;
                border: none;
                border-radius: 5px;
                font-weight: bold;
            }
            #btn_danger:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #EF5350, stop:1 #f44336);
            }
            #btn_danger:disabled {
                background: #424242;
                color: #757575;
            }
            QProgressBar {
                background-color: #1e1e1e;
                border: none;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #4CAF50, stop:1 #66BB6A);
                border-radius: 2px;
            }
            QTextEdit {
                background-color: #1e1e1e;
                color: #FFEB3B;
                border: 1px solid #3a3a3a;
                border-radius: 5px;
                font-family: Consolas;
            }
            QSplitter::handle {
                background-color: #3a3a3a;
            }
        """)
    
    def _connect_daq(self):
        """Connect to DAQ"""
        success, message = self.daq.connect()
        
        if success:
            self.daq_indicator.set_status(True)
            self.daq_status_label.setText("Connected")
            self.daq_status_label.setStyleSheet("color: #4CAF50;")
        else:
            self.daq_indicator.set_status(False)
            self.daq_status_label.setText("Disconnected")
            self.daq_status_label.setStyleSheet("color: #f44336;")
            QMessageBox.critical(self, "DAQ Error", message)
    
    def start_recording(self):
        """Start recording"""
        if not self.daq.is_connected:
            QMessageBox.warning(self, "Error", "DAQ not connected!")
            return
        
        if self.current_mode == "-- Select Test Mode --":
            QMessageBox.warning(self, "Error", "Please select a test mode!")
            return
        
        # Set expected values
        if 'expected_power' in self.current_config:
            power = self.current_config['expected_power']
            if power != "variable":
                self.daq.expected_mw_power = power
                self.daq.expected_grill_power = None
        elif 'expected_mw' in self.current_config:
            self.daq.expected_mw_power = self.current_config['expected_mw']
            self.daq.expected_grill_power = self.current_config.get('expected_grill', None)
        else:
            self.daq.expected_mw_power = None
            self.daq.expected_grill_power = None
        # Reset idle tracking for Normal mode
        if self.current_config.get('type') == 'normal':
            self.idle_time = 0
            self.last_sample_time = None
            self.last_active = True

        # Handle Defrost
        if self.current_config.get('type') == 'defrost':
            dialog = DefrostDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                self.daq.defrost_mode = True
                self.daq.defrost_weight = dialog.weight
                self.daq.defrost_sectors = dialog.sectors
            else:
                return
        
        # Start
        success, message = self.daq.start_recording()
        
        if success:
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            self.save_button.setEnabled(False)
            self.mode_selector.setEnabled(False)
            self.rec_indicator.set_status(True)
            self.rec_status_label.setText("Recording...")
            self.rec_status_label.setStyleSheet("color: #4CAF50;")
            for channel in config.CHANNELS.keys():
                self.graph_data_x[channel] = []
                self.graph_data_y[channel] = []
            self.warnings_text.clear()
            self.result_display.setText("Testing...")
            self.result_display.setStyleSheet("color: #FFA726;")
            if self.daq.expected_mw_power:
                self.stats_widgets['mw_expected'].setText(f"{self.daq.expected_mw_power}%")
            # Set timer to 200ms for 5Hz sampling
            self.update_timer.start(200)
    
    def stop_recording(self):
        """Stop recording"""
        self.update_timer.stop()
        self.daq.stop_recording()
        
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(True)
        self.mode_selector.setEnabled(True)
        
        self.rec_indicator.set_status(False)
        self.rec_status_label.setText("Stopped")
        self.rec_status_label.setStyleSheet("color: #FFA726;")
        
        # Analyze
        results = self.daq.analyze_pass_fail()
        overall = results.get('overall_result', 'N/A')
        if self.current_config and self.current_config.get('type') == 'normal':
            idle_sec = int(self.idle_time)
            idle_min = idle_sec // 60
            idle_rem_sec = idle_sec % 60
            self.result_display.setText(f"Idle Time: {idle_min:02d}:{idle_rem_sec:02d}")
            self.result_display.setStyleSheet("color: #2196F3; font-weight: bold;")
        else:
            self.result_display.setText(overall)
            if overall == 'PASS':
                self.result_display.setStyleSheet("color: #4CAF50; font-weight: bold;")
            elif overall == 'FAIL':
                self.result_display.setStyleSheet("color: #f44336; font-weight: bold;")
        
        # End of Cooking: 3 beeps
        self.play_beep(count=3)
    
    def update_display(self):
        """Update display
        Special logic:
        - NO OVERLAP: MW and Grill never run simultaneously in combination modes (Chicken, C1, C2)
        - Chicken Midtime Alert: Beep, pause, wait for user, resume
        - Child Lock: Disable controls, show status, beep
        - Unified warnings and beeps according to spec
        - Icon state logic: solid/blink/off
        - Always show icons for all signals
        """
        sample_data, error = self.daq.read_sample()
        if error:
            logging.error(f"DAQ Error: {error}")
            self.update_timer.stop()
            self.stop_recording()
            QMessageBox.critical(self, "DAQ Error", error)
            return
        if not sample_data:
            return
        voltages = sample_data['voltages']
        powers = sample_data['powers']
        warnings = sample_data['warnings']
        elapsed = sample_data['elapsed']
        # Map DAQ voltages to logical signals
        daq_signals = {
            'door_open': voltages.get('Door SW', 5.0) >= config.ON_THRESHOLD,
            'start_pressed': voltages.get('Buzzer', 0) >= config.ON_THRESHOLD,  # Example: map Buzzer to Start
            'cancel_pressed': False,  # Add mapping if available
            'knob_turned': False,     # Add mapping if available
            'lock_combo': False,      # Add mapping if available
            'unlock_combo': False     # Add mapping if available
        }
        state = self.state_machine.update(daq_signals)
        self.state_label.setText(f"State: {state.name}")
        # NO OVERLAP LOGIC: Ensure MW and Grill never run simultaneously in combination modes
        if self.current_config and self.current_config.get('type') in ['combination', 'auto']:
            if self.current_config.get('mode') in ['C1', 'C2'] or self.current_config.get('menu') == 'chicken':
                mw_on = voltages.get('Microwave', 0) >= config.ON_THRESHOLD
                grill_on = voltages.get('Grill', 0) >= config.ON_THRESHOLD
                if mw_on and grill_on:
                    warnings.append("NO OVERLAP: MW and Grill should not run simultaneously!")
        # Chicken Midtime Alert logic REMOVED (no pause, no midtime warning)
        # Log state transitions
        if state != self.last_logged_state:
            logging.info(f"State changed to: {state.name}")
            self.last_logged_state = state
        # Log important DAQ events
        if daq_signals['door_open']:
            logging.info("Door is open")
        if daq_signals['start_pressed']:
            logging.info("Start button pressed")
        # Track idle time for Normal mode
        if self.current_config and self.current_config.get('type') == 'normal':
            mw_on = voltages.get('Microwave', 0) >= config.ON_THRESHOLD
            grill_on = voltages.get('Grill', 0) >= config.ON_THRESHOLD
            active = mw_on or grill_on
            if self.last_sample_time is not None:
                dt = elapsed - self.last_sample_time
                if not active:
                    self.idle_time += dt
            self.last_sample_time = elapsed
            self.last_active = active
        # Update signals (status text, style, but always show icon)
        for channel, voltage in voltages.items():
            widget = self.signal_widgets[channel]
            widget['voltage'].setText(f"{voltage:.2f}V")
            is_on = voltage >= config.ON_THRESHOLD
            # Set status and style
            if channel == 'Door SW':
                if 0 <= voltage < 0.5:
                    widget['status'].setText("ON")
                    widget['status'].setStyleSheet("color: #4CAF50;")
                elif 4.5 <= voltage <= 5.0:
                    widget['status'].setText("OFF")
                    widget['status'].setStyleSheet("color: #f44336;")
                else:
                    widget['status'].setText("Unknown")
                    widget['status'].setStyleSheet("color: #757575;")
            elif channel == 'Lamp':
                if is_on:
                    widget['status'].setText("ON")
                    widget['status'].setStyleSheet("color: #FFEB3B;")
                else:
                    widget['status'].setText("OFF")
                    widget['status'].setStyleSheet("color: #757575;")
            elif channel == 'Buzzer':
                if is_on:
                    widget['status'].setText("BEEP")
                    widget['status'].setStyleSheet("color: #FF9800;")
                else:
                    widget['status'].setText("OFF")
                    widget['status'].setStyleSheet("color: #757575;")
            else:  # MW/Grill
                if is_on:
                    widget['status'].setText("ON")
                    widget['status'].setStyleSheet("color: #f44336;")
                else:
                    widget['status'].setText("OFF")
                    widget['status'].setStyleSheet("color: #757575;")
                if 'power' in widget and 'progress' in widget:
                    power = powers.get(channel, 0)
                    widget['power'].setText(f"{power:.1f}%")
                    widget['progress'].setValue(int(power))
        # Always update icons for all signals
        self.update_signal_icons()
        # Update warnings
        if warnings:
            current_time = datetime.now().strftime("%H:%M:%S")
            for warning in warnings:
                self.warnings_text.append(f"[{current_time}] {warning}")
        # Update graphs
        for channel in config.CHANNELS.keys():
            self.graph_data_x[channel].append(elapsed)
            self.graph_data_y[channel].append(voltages[channel])
            while self.graph_data_x[channel] and self.graph_data_x[channel][0] < elapsed - config.GRAPH_WINDOW_SIZE:
                self.graph_data_x[channel].pop(0)
                self.graph_data_y[channel].pop(0)
            self.graph_curves[channel].setData(self.graph_data_x[channel], self.graph_data_y[channel])
            if elapsed > config.GRAPH_WINDOW_SIZE:
                self.graph_widgets[channel].setXRange(elapsed - config.GRAPH_WINDOW_SIZE, elapsed)
            else:
                self.graph_widgets[channel].setXRange(0, config.GRAPH_WINDOW_SIZE)
        # Update stats
        stats = self.daq.get_statistics()
        duration = stats['duration']
        hours = int(duration // 3600)
        minutes = int((duration % 3600) // 60)
        seconds = int(duration % 60)
        self.duration_display.setText(f"{hours:02d}:{minutes:02d}:{seconds:02d}")
        self.samples_display.setText(f"{stats['sample_count']}")
        self.stats_widgets['mw_power'].setText(f"{stats['mw_avg_power']:.1f}%")
        self.stats_widgets['grill_power'].setText(f"{stats['grill_avg_power']:.1f}%")
        self.stats_widgets['door_opens'].setText(f"{stats['door_opens']}")
    
    # _resume_after_midtime removed (midtime pause feature disabled)
    
    def save_data(self):
        """Save with Pass/Fail"""
        safe_mode_name = self.current_mode.replace(':', '').replace('/', '-').replace(' ', '_')
        default_name = f"{safe_mode_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        filename, _ = QFileDialog.getSaveFileName(self, "Save Test Results", default_name, "Excel Files (*.xlsx)")
        
        if not filename:
            return
        
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        
        try:
            self.excel_writer = ExcelWriter()
            self.excel_writer.create_workbook()
            
            data = self.daq.get_all_data()
            
            if len(data) == 0:
                QMessageBox.warning(self, "No Data", "No data to save!")
                return
            
            self.excel_writer.write_data(data)
            
            stats = self.daq.get_statistics()
            duration = stats['duration']
            hours = int(duration // 3600)
            minutes = int((duration % 3600) // 60)
            seconds = int(duration % 60)
            
            test_info = {
                'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'mode': self.current_mode,
                'duration': f"{hours:02d}:{minutes:02d}:{seconds:02d}",
                'samples': stats['sample_count']
            }
            
            # Only add defrost_sectors if current mode is Defrost
            if self.current_config and self.current_config.get('type') == 'defrost':
                test_info['defrost_sectors'] = self.daq.defrost_sectors
            
            pass_fail_results = self.daq.analyze_pass_fail()
            
            self.excel_writer.add_summary_sheet(stats, test_info, pass_fail_results)
            
            success, message = self.excel_writer.save(filename)
            
            if success:
                overall_result = pass_fail_results.get('overall_result', 'N/A')
                
                msg = f"═══════════════════════════════\n"
                msg += f"  TEST RESULT: {overall_result}\n"
                msg += f"═════════════════════════════\n\n"
                msg += f"File: {filename}\n\n"
                msg += "Details:\n"
                for detail in pass_fail_results.get('details', []):
                    msg += f"{detail}\n"
                msg += f"\nOpen file location?"
                
                msgbox = QMessageBox(self)
                msgbox.setWindowTitle("Test Complete")
                msgbox.setText(msg)
                
                if overall_result == 'PASS':
                    msgbox.setIcon(QMessageBox.Information)
                else:
                    msgbox.setIcon(QMessageBox.Warning)
                
                msgbox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                
                if msgbox.exec_() == QMessageBox.Yes:
                    folder = os.path.dirname(filename)
                    
                    if platform.system() == 'Windows':
                        subprocess.Popen(f'explorer /select,"{filename}"')
                    elif platform.system() == 'Darwin':
                        subprocess.Popen(['open', '-R', filename])
                    else:
                        subprocess.Popen(['xdg-open', folder])
            else:
                QMessageBox.critical(self, "Save Failed", message)
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {str(e)}")
            traceback.print_exc()
    
    def closeEvent(self, event):
        """Handle close"""
        self.update_timer.stop()
        self.daq.disconnect()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 9))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()