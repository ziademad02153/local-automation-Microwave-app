"""
Configuration file for Microwave DAQ Testing System
Contains all device settings, channel mappings, and constants
"""

# ==================== DEVICE CONFIGURATION ====================
DEVICE_NAME = "cDAQ1Mod1"

# Channel Mapping (Physical connections)
CHANNELS = {
    'Door SW': 'ai0',      # Door switch (REVERSED: 0V=Closed, 5V=Open)
    'Lamp': 'ai1',         # Cavity lamp
    'Microwave': 'ai2',    # MW relay (Power calculation)
    'Grill': 'ai3',        # Grill relay (Power calculation)
    'Buzzer': 'ai4'        # Buzzer
}

# Voltage settings
VOLTAGE_MIN = -1.0  # Volts
VOLTAGE_MAX = 6.0   # Volts
VOLTAGE_RANGE_MIN = 0.0  # For DAQ configuration
VOLTAGE_RANGE_MAX = 5.0  # For DAQ configuration

# ==================== SIGNAL THRESHOLDS ====================
ON_THRESHOLD = 4.6  # Voltage >= 4.6V considered as ON
OUT_OF_RANGE_HIGH = 5.5  # Warning if voltage > 5.5V
OUT_OF_RANGE_LOW = -0.5  # Warning if voltage < -0.5V

# Door is REVERSED
DOOR_CLOSED_THRESHOLD = 0.5  # Door voltage < 0.5V = CLOSED

# ==================== SAMPLING CONFIGURATION ====================
SAMPLING_RATE = 1  # Hz (1 sample per second)
POWER_CALC_WINDOW = 100  # Calculate power based on last 100 samples

# ==================== WARNING SETTINGS ====================
OVERLAP_TOLERANCE = 2  # Seconds - MW+Grill overlap > 2s triggers warning

# ==================== GRAPH SETTINGS ====================
GRAPH_WINDOW_SIZE = 60  # Seconds (display last 60 seconds)
GRAPH_UPDATE_INTERVAL = 1000  # Milliseconds (1 second)

# Graph colors (for 5 separate graphs)
GRAPH_COLORS = {
    'Door SW': '#00FF00',      # Green
    'Lamp': '#FFFF00',         # Yellow
    'Microwave': '#FF0000',    # Red
    'Grill': '#FF8800',        # Orange
    'Buzzer': '#00FFFF'        # Cyan
}

# ==================== GUI THEME (Dark Professional) ====================
THEME = {
    'background': '#2b2b2b',
    'text': '#ffffff',
    'panel_bg': '#3c3c3c',
    'border': '#555555',
    'button_start': '#4CAF50',
    'button_stop': '#f44336',
    'button_save': '#2196F3',
    'warning': '#FF9800',
    'error': '#F44336',
    'success': '#4CAF50'
}

# ==================== PASS/FAIL TOLERANCE ====================
PASS_FAIL_TOLERANCE = 5  # Percentage tolerance for power levels

# ==================== DEFROST CALCULATION ====================
# Defrost formula constants
DEFROST_CONFIG = {
    'weight_step': 100,  # grams
    'constant_factor': 2.05,  # minutes per (weight/100)
    
    # Sector percentages and power
    'sectors': [
        {
            'name': 'Sector 1',
            'percentage': 0.14,  # 14% of total time
            'power': 36.7,       # Expected power %
            'on_time': 11,       # seconds
            'off_time': 19,      # seconds
            'period': 30         # seconds
        },
        {
            'name': 'Sector 2',
            'percentage': 0.50,  # 50% of total time
            'power': 23.3,
            'on_time': 7,
            'off_time': 23,
            'period': 30
        },
        {
            'name': 'Sector 3',
            'percentage': 0.36,  # 36% of total time
            'power': 30.0,
            'on_time': 9,
            'off_time': 21,
            'period': 30
        }
    ],
    
    'weight_range': (100, 2000)  # grams (min, max)
}

# ==================== EXCEL SETTINGS ====================
EXCEL_COLUMNS = [
    'H', 'Min', 'Sec', 'ms',
    'Microwave', 'Lamp', 'Door_SW', 'Buzzer', 'Grill',
    'MW_Power%', 'Grill_Power%'
]

# ==================== ICON SYMBOLS ====================
ICONS = {
    'ON': '🔴',
    'OFF': '⚫',
    'CLOSED': '🟢',
    'OPEN': '🔴',
    'LAMP_ON': '💡',
    'LAMP_OFF': '⚫',
    'BUZZER_ON': '🔊',
    'BUZZER_OFF': '⚫'
}