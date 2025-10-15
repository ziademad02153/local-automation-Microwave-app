"""
DAQ Handler Module - UPDATED
Handles all DAQ communication, data acquisition, and power calculations
Door SW is REVERSED: 0V = Closed, 5V = Open
"""

import nidaqmx
from nidaqmx.constants import TerminalConfiguration
import numpy as np
from datetime import datetime, timedelta
from collections import deque
import config

class DAQHandler:
    def __init__(self):
        self.task = None
        self.is_connected = False
        self.is_recording = False
        
        # Data buffers for each channel
        self.data_buffers = {channel: deque(maxlen=10000) for channel in config.CHANNELS.keys()}
        self.timestamps = deque(maxlen=10000)
        
        # Statistics
        self.start_time = None
        self.sample_count = 0
        self.door_open_count = 0
        self.last_door_state = False
        
        # Overlap detection
        self.overlap_start_time = None
        self.overlap_duration = 0
        
        # Defrost mode
        self.defrost_mode = False
        self.defrost_weight = 0
        self.defrost_sectors = []
        
        # For Pass/Fail analysis
        self.test_mode = None
        self.expected_mw_power = None
        self.expected_grill_power = None
        
    def connect(self):
        """Connect to DAQ device"""
        try:
            self.task = nidaqmx.Task()
            
            # Add all channels
            for name, channel in config.CHANNELS.items():
                self.task.ai_channels.add_ai_voltage_chan(
                    f"{config.DEVICE_NAME}/{channel}",
                    name_to_assign_to_channel=name,
                    terminal_config=TerminalConfiguration.RSE,
                    min_val=config.VOLTAGE_RANGE_MIN,
                    max_val=config.VOLTAGE_RANGE_MAX
                )
            
            self.is_connected = True
            return True, "DAQ Connected Successfully"
            
        except Exception as e:
            self.is_connected = False
            return False, f"DAQ Connection Failed: {str(e)}"
    
    def disconnect(self):
        """Disconnect from DAQ"""
        try:
            if self.task:
                self.task.close()
                self.task = None
            self.is_connected = False
            return True, "DAQ Disconnected"
        except Exception as e:
            return False, f"Error disconnecting: {str(e)}"
    
    def start_recording(self):
        """Start data recording"""
        if not self.is_connected:
            return False, "DAQ not connected"
        
        self.is_recording = True
        self.start_time = datetime.now()
        self.sample_count = 0
        self.door_open_count = 0
        
        # Clear buffers
        for buffer in self.data_buffers.values():
            buffer.clear()
        self.timestamps.clear()
        
        return True, "Recording Started"
    
    def stop_recording(self):
        """Stop data recording"""
        self.is_recording = False
        return True, "Recording Stopped"
    
    def read_sample(self):
        """Read one sample from all channels"""
        if not self.is_connected or not self.is_recording:
            return None, "Not recording"
        
        try:
            # Read data from all channels
            data = self.task.read()
            timestamp = datetime.now()
            
            # Store in buffers
            self.timestamps.append(timestamp)
            for i, channel_name in enumerate(config.CHANNELS.keys()):
                self.data_buffers[channel_name].append(data[i])
            
            self.sample_count += 1
            
            # Create data dictionary
            sample_data = {
                'timestamp': timestamp,
                'elapsed': (timestamp - self.start_time).total_seconds(),
                'voltages': {}
            }
            
            for i, channel_name in enumerate(config.CHANNELS.keys()):
                sample_data['voltages'][channel_name] = data[i]
            
            # Check for warnings
            warnings = self._check_warnings(sample_data)
            sample_data['warnings'] = warnings
            
            # Calculate power percentages
            powers = self._calculate_powers()
            sample_data['powers'] = powers
            
            return sample_data, None
            
        except Exception as e:
            self.is_connected = False
            return None, f"DAQ Error: {str(e)}"
    
    def _check_warnings(self, sample_data):
        """Check for warning conditions"""
        warnings = []
        voltages = sample_data['voltages']
        
        # Check out of range
        for channel, voltage in voltages.items():
            if voltage > config.OUT_OF_RANGE_HIGH or voltage < config.OUT_OF_RANGE_LOW:
                warnings.append(f"⚠️ {channel} out of range: {voltage:.2f}V")
        
        # Check door opened (REVERSED LOGIC)
        door_voltage = voltages.get('Door SW', 5.0)
        door_open = door_voltage >= config.ON_THRESHOLD  # HIGH = OPEN
        
        if door_open and not self.last_door_state:
            self.door_open_count += 1
            warnings.append("⚠️ Door opened during test")
        
        self.last_door_state = door_open
        
        # Check MW + Grill overlap
        mw_voltage = voltages.get('Microwave', 0)
        grill_voltage = voltages.get('Grill', 0)
        
        mw_on = mw_voltage >= config.ON_THRESHOLD
        grill_on = grill_voltage >= config.ON_THRESHOLD
        
        if mw_on and grill_on:
            if self.overlap_start_time is None:
                self.overlap_start_time = sample_data['timestamp']
            else:
                self.overlap_duration = (sample_data['timestamp'] - self.overlap_start_time).total_seconds()
                
                if self.overlap_duration > config.OVERLAP_TOLERANCE:
                    warnings.append(f"⚠️ MW + Grill overlap detected: {self.overlap_duration:.1f}s")
        else:
            self.overlap_start_time = None
            self.overlap_duration = 0
        
        return warnings
    
    def _calculate_powers(self):
        """Calculate power percentages based on last N samples"""
        powers = {}
        
        # Calculate for Microwave
        mw_buffer = list(self.data_buffers['Microwave'])
        if len(mw_buffer) > 0:
            window = mw_buffer[-config.POWER_CALC_WINDOW:] if len(mw_buffer) >= config.POWER_CALC_WINDOW else mw_buffer
            on_count = sum(1 for v in window if v >= config.ON_THRESHOLD)
            powers['Microwave'] = (on_count / len(window)) * 100 if len(window) > 0 else 0
            powers['Microwave_samples'] = len(window)
        else:
            powers['Microwave'] = 0
            powers['Microwave_samples'] = 0
        
        # Calculate for Grill
        grill_buffer = list(self.data_buffers['Grill'])
        if len(grill_buffer) > 0:
            window = grill_buffer[-config.POWER_CALC_WINDOW:] if len(grill_buffer) >= config.POWER_CALC_WINDOW else grill_buffer
            on_count = sum(1 for v in window if v >= config.ON_THRESHOLD)
            powers['Grill'] = (on_count / len(window)) * 100 if len(window) > 0 else 0
            powers['Grill_samples'] = len(window)
        else:
            powers['Grill'] = 0
            powers['Grill_samples'] = 0
        
        return powers
    
    def calculate_defrost_sectors(self, weight_grams):
        """Calculate defrost sector timings based on weight"""
        if weight_grams < config.DEFROST_CONFIG['weight_range'][0] or \
           weight_grams > config.DEFROST_CONFIG['weight_range'][1]:
            return None, f"Weight must be between {config.DEFROST_CONFIG['weight_range'][0]}-{config.DEFROST_CONFIG['weight_range'][1]}g"
        
        # Calculate total time
        weight_step = config.DEFROST_CONFIG['weight_step']
        constant = config.DEFROST_CONFIG['constant_factor']
        total_time_minutes = constant * (weight_grams / weight_step)
        total_time_seconds = total_time_minutes * 60
        
        # Calculate sector boundaries
        sectors = []
        cumulative_time = 0
        
        for sector_config in config.DEFROST_CONFIG['sectors']:
            sector_duration = total_time_seconds * sector_config['percentage']
            
            sector = {
                'name': sector_config['name'],
                'start_time': cumulative_time,
                'end_time': cumulative_time + sector_duration,
                'duration': sector_duration,
                'expected_power': sector_config['power'],
                'on_time': sector_config['on_time'],
                'off_time': sector_config['off_time'],
                'period': sector_config['period']
            }
            
            sectors.append(sector)
            cumulative_time += sector_duration
        
        self.defrost_mode = True
        self.defrost_weight = weight_grams
        self.defrost_sectors = sectors
        
        return sectors, f"Total time: {int(total_time_minutes)}:{int((total_time_minutes % 1) * 60):02d}"
    
    def get_current_defrost_sector(self, elapsed_time):
        """Get current defrost sector based on elapsed time"""
        if not self.defrost_mode:
            return None
        
        for sector in self.defrost_sectors:
            if sector['start_time'] <= elapsed_time < sector['end_time']:
                return sector
        
        return None
    
    def get_statistics(self):
        """Get recording statistics"""
        if self.start_time is None:
            duration = 0
        else:
            duration = (datetime.now() - self.start_time).total_seconds()
        
        powers = self._calculate_powers()
        
        stats = {
            'duration': duration,
            'sample_count': self.sample_count,
            'door_opens': self.door_open_count,
            'mw_avg_power': powers.get('Microwave', 0),
            'grill_avg_power': powers.get('Grill', 0),
            'mw_samples': powers.get('Microwave_samples', 0),
            'grill_samples': powers.get('Grill_samples', 0)
        }
        
        return stats
    
    def analyze_pass_fail(self):
        """Analyze test results for Pass/Fail - NEW"""
        stats = self.get_statistics()
        
        results = {
            'overall_result': 'PASS',
            'mw_result': 'N/A',
            'grill_result': 'N/A',
            'mw_measured': stats['mw_avg_power'],
            'grill_measured': stats['grill_avg_power'],
            'mw_expected': self.expected_mw_power,
            'grill_expected': self.expected_grill_power,
            'details': []
        }
        
        # Check MW power
        if self.expected_mw_power is not None and self.expected_mw_power != 'variable':
            tolerance = config.PASS_FAIL_TOLERANCE
            lower = self.expected_mw_power - tolerance
            upper = self.expected_mw_power + tolerance
            
            if lower <= stats['mw_avg_power'] <= upper:
                results['mw_result'] = 'PASS'
                results['details'].append(f"✅ MW Power: {stats['mw_avg_power']:.1f}% (Expected: {self.expected_mw_power}% ±{tolerance}%)")
            else:
                results['mw_result'] = 'FAIL'
                results['overall_result'] = 'FAIL'
                results['details'].append(f"❌ MW Power: {stats['mw_avg_power']:.1f}% (Expected: {self.expected_mw_power}% ±{tolerance}%)")
        
        # Check Grill power
        if self.expected_grill_power is not None:
            tolerance = config.PASS_FAIL_TOLERANCE
            lower = self.expected_grill_power - tolerance
            upper = self.expected_grill_power + tolerance
            
            if lower <= stats['grill_avg_power'] <= upper:
                results['grill_result'] = 'PASS'
                results['details'].append(f"✅ Grill Power: {stats['grill_avg_power']:.1f}% (Expected: {self.expected_grill_power}% ±{tolerance}%)")
            else:
                results['grill_result'] = 'FAIL'
                results['overall_result'] = 'FAIL'
                results['details'].append(f"❌ Grill Power: {stats['grill_avg_power']:.1f}% (Expected: {self.expected_grill_power}% ±{tolerance}%)")
        
        # Check door opens
        if stats['door_opens'] > 0:
            results['details'].append(f"⚠️ Door was opened {stats['door_opens']} time(s) during test")
        
        return results
    
    def get_all_data(self):
        """Get all recorded data for export"""
        data = []
        
        # Map channel names to Excel column names
        channel_to_excel = {
            'Door SW': 'Door_SW',
            'Lamp': 'Lamp',
            'Microwave': 'Microwave',
            'Grill': 'Grill',
            'Buzzer': 'Buzzer'
        }
        
        for i in range(len(self.timestamps)):
            elapsed = (self.timestamps[i] - self.start_time).total_seconds()
            hours = int(elapsed // 3600)
            minutes = int((elapsed % 3600) // 60)
            seconds = int(elapsed % 60)
            milliseconds = int((elapsed % 1) * 1000)
            
            row = {
                'H': hours,
                'Min': minutes,
                'Sec': seconds,
                'ms': milliseconds
            }
            
            # Add voltage data with Excel-compatible names
            for channel_name, excel_name in channel_to_excel.items():
                row[excel_name] = self.data_buffers[channel_name][i]
            
            # Calculate power for this point (rolling window)
            if i >= config.POWER_CALC_WINDOW:
                mw_window = list(self.data_buffers['Microwave'])[i-config.POWER_CALC_WINDOW:i]
                grill_window = list(self.data_buffers['Grill'])[i-config.POWER_CALC_WINDOW:i]
            else:
                mw_window = list(self.data_buffers['Microwave'])[:i+1]
                grill_window = list(self.data_buffers['Grill'])[:i+1]
            
            mw_on = sum(1 for v in mw_window if v >= config.ON_THRESHOLD)
            grill_on = sum(1 for v in grill_window if v >= config.ON_THRESHOLD)
            
            row['MW_Power%'] = (mw_on / len(mw_window)) * 100 if len(mw_window) > 0 else 0
            row['Grill_Power%'] = (grill_on / len(grill_window)) * 100 if len(grill_window) > 0 else 0
            
            data.append(row)
        
        return data