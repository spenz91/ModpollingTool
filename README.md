# ModPolling Tool

A graphical user interface (GUI) tool for interacting with Modpoll, a Modbus master simulator and test utility. This tool allows users to configure Modbus communication settings, select equipment presets, start and stop polling, and view logs in real-time.

## Features

- **Equipment Presets**: Quickly select predefined settings for various equipment models.
- **Custom Configuration**: Manually set COM port, baud rate, parity, data bits, stop bits, and more.
- **Modbus/TCP Support**: Option to use Modbus over TCP/IP.
- **Real-time Logging**: View Modpoll output and errors within the GUI.
- **Status Indicator**: Visual feedback for polling status with color-coded indicators.
- **Advanced Settings**: Access advanced Modbus settings through an intuitive interface.

<p align="center">
    <img width="300" src="https://i.ibb.co/SyFVh35/Skjermbilde-2024-11-18-142520.png" alt="Material Bread logo">
</p>

## Requirements

- **Python 3.x**
- **Tkinter** (usually included with Python)
- **PySerial**: Install via `pip install pyserial`

## Installation

1. **Clone or Download the Repository**

   - Clone the repository using Git:

     ```bash
     git clone https://github.com/yourusername/modpolling-tool.git
     ```

   - Or download the ZIP file and extract it.

2. **Install Dependencies**

   - Navigate to the project directory:

     ```bash
     cd modpolling-tool
     ```

   - Install required Python packages:

     ```bash
     pip install pyserial
     ```

3. **Download Modpoll**

   - Download `modpoll.exe` from the [Modpoll Official Website](https://www.modbusdriver.com/modpoll.html).
   - Place `modpoll.exe` in the `C:\iwmac\bin\` directory. If the directory doesn't exist, create it.

## Usage

1. **Run the Tool**

   - Execute the Python script:

     ```bash
     python modpolling_tool.py
     ```

2. **Select Equipment**

   - From the **Equipment** panel, double-click an equipment model to auto-populate communication settings.

3. **Configure Settings**

   - **Basic Tab**:
     - **COM Port**: Select or refresh available COM ports.
     - **Baudrate**: Choose the baud rate.
     - **Parity**: Set the parity (none, even, odd).
     - **Address**: Enter the slave address.
   - **Advanced Tab**:
     - **Data Bits**: Set data bits (7 or 8).
     - **Stop Bits**: Set stop bits (1 or 2).
     - **Start Reference**: Enter the starting reference register.
     - **Number of Registers**: Specify how many registers to read.
     - **Register Data Type**: Define the register data type.
     - **Modbus/TCP**: Enter the IP address and port if using Modbus over TCP.

4. **Start Polling**

   - Click **Start Polling** to begin communication.
   - The **Log** panel will display real-time data and messages.
   - The status indicator will provide visual feedback:
     - **Green**: Successful communication.
     - **Yellow**: Warnings like checksum errors.
     - **Red**: Errors such as time-outs or port issues.

5. **Stop Polling**

   - Click **Stop Polling** to end the session safely.

## Notes

- Ensure that `modpoll.exe` is correctly placed in `C:\iwmac\bin\`.
- The tool is designed for **Windows OS** due to the use of Windows-specific features like registry access.
- If you're using a COM port with a number higher than 9 (e.g., COM10), the tool formats it correctly for Modpoll.
- Always run the tool with appropriate permissions if you encounter access issues.

## Troubleshooting

- **COM Port Issues**:
  - If no COM ports are listed, click **Refresh** next to the COM port selector.
  - The tool scans both pySerial and Windows Registry for available COM ports.

- **Modpoll Not Found**:
  - If you receive an error stating `modpoll.exe not found`, verify that it's placed in `C:\iwmac\bin\`.

- **Access Denied Errors**:
  - Ensure that no other applications are using the COM port.
  - Close any instances of Modpoll or similar tools before starting polling.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Modpoll](https://www.modbusdriver.com/modpoll.html) by ModbusDriver.
- The tool uses the [PySerial](https://github.com/pyserial/pyserial) library for serial communication.
