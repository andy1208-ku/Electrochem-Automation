#!/usr/bin/env python3
"""
run_measurement_serial.py - Execute MethodSCRIPT measurements via serial
"""

import sys
import os
import csv
import json
import time
from datetime import datetime
from pathlib import Path
import serial
import serial.tools.list_ports

class SerialMeasurementRunner:
    def __init__(self, script_path, save_data=True):
        self.script_path = Path(script_path)
        self.save_data = save_data
        self.data_points = []
        self.connection = None
        
        # Create data directory structure
        self.data_base_path = Path("measurement_data")
        self.data_base_path.mkdir(exist_ok=True)
        
        date_folder = datetime.now().strftime('%Y-%m-%d')
        self.data_folder = self.data_base_path / date_folder
        self.data_folder.mkdir(exist_ok=True)
    
    def find_device_port(self):
        """Auto-detect EmStat Pico serial port"""
        print("Scanning for devices...")
        ports = serial.tools.list_ports.comports(include_links=False)
        
        candidates = []
        for port in ports:
            print(f"  Found port: {port.description} ({port.device})")
            # Check for known device names
            if any(name in port.description for name in 
                   ['ESPicoDev', 'EmStat', 'USB Serial Port', 'FTDI']):
                candidates.append(port.device)
        
        if not candidates:
            print("ERROR: No device found")
            return None
        elif len(candidates) > 1:
            print(f"Multiple devices found: {candidates}")
            print(f"Using first device: {candidates[0]}")
        
        return candidates[0]
    
    def connect(self, port=None):
        """Connect to device via serial"""
        if port is None:
            port = self.find_device_port()
            if port is None:
                return False
        
        try:
            print(f"Connecting to {port}...")
            self.connection = serial.Serial(
                port=port,
                baudrate=230400,
                timeout=1,
                write_timeout=1
            )
            time.sleep(2)  # Give device time to initialize
            
            # Clear any pending data
            self.connection.reset_input_buffer()
            self.connection.reset_output_buffer()
            
            # Send version query to test connection
            self.connection.write(b't\n')
            response = self.connection.readline()
            if response:
                print(f"Device responded: {response.decode('utf-8', errors='ignore').strip()}")
                return True
            else:
                print("No response from device")
                return False
                
        except Exception as e:
            print(f"Connection failed: {e}")
            return False
    
    def run_script(self, script):
        """Send MethodSCRIPT to device and collect data"""
        if not self.connection:
            print("ERROR: Not connected to device")
            return False
        
        try:
            print("Sending script to device...")
            
            # Send the script line by line
            lines = script.strip().split('\n')
            for line in lines:
                self.connection.write((line + '\n').encode('utf-8'))
                time.sleep(0.01)  # Small delay between lines
            
            # Send empty line to signal end of script
            self.connection.write(b'\n')
            
            print("Script sent. Collecting data...")
            print("-" * 40)
            
            # Collect response data
            timeout_counter = 0
            max_timeout = 600  # 10 minutes max
            
            while timeout_counter < max_timeout:
                try:
                    line = self.connection.readline()
                    if not line:
                        timeout_counter += 1
                        continue
                    
                    # Decode and process line
                    text = line.decode('utf-8', errors='ignore').strip()
                    if not text:
                        continue
                    
                    print(text)  # Show raw output
                    
                    # Parse data lines (format: "P<value>;DA<value>;BA<value>")
                    if text.startswith('P'):
                        self.parse_data_line(text)
                    
                    # Check for end markers
                    if text in ['*', 'Measurement completed', 'Script completed']:
                        print("\nMeasurement completed")
                        break
                    
                    # Check for errors
                    if text.startswith('!'):
                        print(f"Device error: {text}")
                        if "abort" in text.lower():
                            break
                    
                    timeout_counter = 0  # Reset timeout on valid data
                    
                except serial.SerialTimeoutException:
                    timeout_counter += 1
                    continue
            
            if timeout_counter >= max_timeout:
                print("Timeout waiting for measurement completion")
            
            return True
            
        except Exception as e:
            print(f"Error running script: {e}")
            return False
    
    def parse_data_line(self, line):
        """Parse a data line from the device"""
        # Example format: "P-500m;DA-2.345e-6;BA100u"
        try:
            parts = line.split(';')
            data_point = {}
            
            for part in parts:
                if part.startswith('P'):
                    # Potential in mV
                    value = self.parse_value(part[1:])
                    data_point['potential'] = value / 1000.0  # Convert to V
                elif part.startswith('DA') or part.startswith('BA'):
                    # Current (DA = measured, BA = range)
                    if part.startswith('DA'):
                        value = self.parse_value(part[2:])
                        data_point['current'] = value * 1e6  # Convert to µA
                        
            if 'potential' in data_point and 'current' in data_point:
                self.data_points.append(data_point)
                
        except Exception as e:
            # Ignore parse errors for non-data lines
            pass
    
    def parse_value(self, text):
        """Parse a value with SI suffix (m, u, n, etc.)"""
        multipliers = {
            'n': 1e-9, 'u': 1e-6, 'm': 1e-3,
            '': 1, 'k': 1e3, 'M': 1e6
        }
        
        # Handle scientific notation
        if 'e' in text.lower():
            return float(text)
        
        # Find suffix
        for suffix in multipliers:
            if text.endswith(suffix):
                if suffix:
                    number = text[:-len(suffix)]
                else:
                    number = text
                try:
                    return float(number) * multipliers[suffix]
                except:
                    return 0.0
        
        # Try direct conversion
        try:
            return float(text)
        except:
            return 0.0
    
    def save_data(self):
        """Save collected data to CSV"""
        if not self.data_points:
            print("No data to save")
            return None
        
        # Generate filename
        base_name = self.script_path.stem
        timestamp = datetime.now().strftime('%H%M%S')
        csv_filename = self.data_folder / f"{base_name}_{timestamp}.csv"
        
        # Save CSV
        with open(csv_filename, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['potential', 'current'])
            writer.writerow({'potential': 'Potential (V)', 'current': 'Current (µA)'})
            writer.writerows(self.data_points)
        
        print(f"\nData saved to: {csv_filename}")
        return csv_filename
    
    def disconnect(self):
        """Close serial connection"""
        if self.connection:
            try:
                self.connection.close()
                print("Disconnected from device")
            except:
                pass
    
    def run(self):
        """Main execution"""
        print("=" * 60)
        print("Serial MethodSCRIPT Measurement")
        print(f"Script: {self.script_path}")
        print("=" * 60)
        
        # Read script
        try:
            with open(self.script_path, 'r') as f:
                script = f.read()
            print(f"Script loaded: {len(script)} characters")
        except Exception as e:
            print(f"ERROR: Failed to read script: {e}")
            return 1
        
        # Connect to device
        if not self.connect():
            print("ERROR: Failed to connect to device")
            return 1
        
        try:
            # Run measurement
            if self.run_script(script):
                # Save data
                if self.save_data and self.data_points:
                    self.save_data()
                print(f"Total data points: {len(self.data_points)}")
                return 0
            else:
                return 1
                
        finally:
            self.disconnect()


def main():
    if len(sys.argv) < 2:
        print("Usage: python run_measurement_serial.py <script_path>")
        sys.exit(1)
    
    script_path = sys.argv[1]
    if not os.path.exists(script_path):
        print(f"ERROR: Script file not found: {script_path}")
        sys.exit(1)
    
    runner = SerialMeasurementRunner(script_path)
    exit_code = runner.run()
    sys.exit(exit_code)

if __name__ == "__main__":
    main()