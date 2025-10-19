"""
Microwave State Machine Module
A generic state machine for managing microwave states based on DAQ signals.
No emojis, clean code, English only, global best practices.
"""

from enum import Enum, auto
import time

class MicrowaveState(Enum):
    IDLE = auto()
    RUN = auto()
    PAUSE = auto()
    SLEEP = auto()
    LOCKED = auto()

class MicrowaveStateMachine:
    def __init__(self, sleep_timeout=900):
        self.state = MicrowaveState.IDLE
        self.last_interaction = time.time()
        self.sleep_timeout = sleep_timeout  # seconds (default 15 min)
        self.locked = False

    def update(self, daq_signals):
        """
        Update state based on DAQ signals.
        daq_signals: dict with keys like 'door_open', 'start_pressed', 'cancel_pressed', 'knob_turned', etc.
        """
        now = time.time()
        if self.locked:
            self.state = MicrowaveState.LOCKED
            if daq_signals.get('unlock_combo', False):
                self.locked = False
                self.state = MicrowaveState.IDLE
                self.last_interaction = now
            return self.state

        # Any user interaction resets sleep timer
        if any([daq_signals.get('start_pressed'), daq_signals.get('cancel_pressed'), daq_signals.get('knob_turned'), daq_signals.get('door_open')]):
            self.last_interaction = now
            if self.state == MicrowaveState.SLEEP:
                self.state = MicrowaveState.IDLE

        # Lock combo
        if daq_signals.get('lock_combo', False):
            self.locked = True
            self.state = MicrowaveState.LOCKED
            return self.state

        # Sleep mode
        if (now - self.last_interaction) > self.sleep_timeout:
            self.state = MicrowaveState.SLEEP
            return self.state

        # State transitions
        if self.state == MicrowaveState.IDLE:
            if daq_signals.get('start_pressed') and not daq_signals.get('door_open'):
                self.state = MicrowaveState.RUN
        elif self.state == MicrowaveState.RUN:
            if daq_signals.get('door_open'):
                self.state = MicrowaveState.PAUSE
            elif daq_signals.get('cancel_pressed'):
                self.state = MicrowaveState.IDLE
        elif self.state == MicrowaveState.PAUSE:
            if not daq_signals.get('door_open') and daq_signals.get('start_pressed'):
                self.state = MicrowaveState.RUN
            elif daq_signals.get('cancel_pressed'):
                self.state = MicrowaveState.IDLE
        elif self.state == MicrowaveState.SLEEP:
            if any([daq_signals.get('start_pressed'), daq_signals.get('cancel_pressed'), daq_signals.get('knob_turned'), daq_signals.get('door_open')]):
                self.state = MicrowaveState.IDLE
                self.last_interaction = now
        return self.state

    def is_locked(self):
        return self.locked

    def get_state(self):
        return self.state
