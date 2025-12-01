Pyrec Logger

Pyrec Logger is a robust, Python-based desktop application designed for real-time sensor data acquisition, visualization, and logging. It features a responsive 24-hour view, logarithmic time scaling, and support for both serial devices and simulation modes.

Features

Real-Time Visualization: Live plotting of up to 8 sensor channels using matplotlib.

Data Persistence: Automatic daily CSV logging (log_YYYY-MM-DD.csv) with background threading.

Interactive Navigation: * Pan & Zoom: Inspect historical data without interrupting the logger.

Logarithmic Slider: Adjust the view window from 1 minute up to 24 hours instantly.

Auto-Scroll: Locks to the latest data stream.

Channel Control: Individually toggle channels, set custom colors, and apply real-time math (Factor/Offset) to raw readings.

Data Export: Built-in tool to export specific time ranges to Excel (.xlsx), preserving raw data and applying calibration formulas automatically.

Hardware Support:

Simulation Mode: Generates dummy data for testing without hardware.

Standard Serial: Compatible with any device sending comma-separated values (CSV) over USB/Serial.

BalkonLogger: Specialized driver support for batch-protocol devices.


Installation

Prerequisites

Python 3.8+

Install Dependencies

The application relies on matplotlib for plotting. For serial connection and Excel export features, you will need additional libraries.

# Core requirements
pip install matplotlib

# For Serial/USB connection (Required for hardware)
pip install pyserial

# For Excel Export functionality (Highly Recommended)
pip install pandas openpyxl


Usage

Start the Application:

python main.py


Connecting to Data:

Simulation: Selected by default. Click Connect to see generated sine waves and noise.

Hardware: Select your COM port from the dropdown, choose your Baud Rate (e.g., 9600), and select the "Standard" driver for generic Arduino/microcontroller data.

Navigation:

Slider: Drag the bottom slider to change the time window (Left = 1 min, Right = 24 hours).

Toolbar Buttons:

🏠 Home: Return to live view and reset axes.

✋ Pan: Drag the graph to look back in time (pauses auto-scroll).

🔍 Zoom: Draw a rectangle to zoom into a specific anomaly.

💾 Save: Save the current plot as an image.

📊 Export: Open the Excel export menu.

Channel Settings:

Use the sidebar to toggle channels on/off.

Factor: Multiplies the raw value (e.g., convert Volts to degrees).

Offset: Adds to the raw value (calibration).


Serial Data Format

To use the Standard driver with your own hardware (Arduino, ESP32, Raspberry Pi), send data as a simple comma-separated line ending with a newline character.

Example Arduino Code:

void loop() {
  float sensor1 = analogRead(A0);
  float sensor2 = analogRead(A1);
  
  Serial.print(sensor1);
  Serial.print(",");
  Serial.println(sensor2);
  
  delay(1000);
}


Data Storage

Logs: Data is saved automatically to log_YYYY-MM-DD.csv in the application directory.

Settings: UI preferences (window size, channel factors/offsets) are saved to sensor_settings.ini.


Configuration

You can edit sensor_settings.ini generated after the first run to persist your graph limits and channel configurations.


Contributing

Feel free to fork this repository and submit pull requests. For major changes, please open an issue first to discuss what you would like to change.
