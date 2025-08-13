# ModPolling Tool

A professional GUI application for Modbus communication testing and device configuration using the modpoll utility.

![ModPolling Tool](https://img.shields.io/badge/Version-2.0-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Download](#download)
- [Installation](#installation)
- [How It Works](#how-it-works)
- [Equipment Presets](#equipment-presets)
- [Usage Guide](#usage-guide)
- [Requirements](#requirements)
- [Building from Source](#building-from-source)
- [Contributing](#contributing)
- [License](#license)

## 🎯 Overview

ModPolling Tool is a comprehensive Windows GUI application designed to simplify Modbus communication testing and device configuration. It provides an intuitive interface for the modpoll utility, making it easy to communicate with Modbus RTU and TCP devices without memorizing complex command-line parameters.

## ✨ Features

- **🖥️ User-Friendly GUI**: Clean, modern interface built with tkinter
- **🔧 Equipment Presets**: Pre-configured settings for popular Carlo Gavazzi and SLV devices
- **⚡ Single-Click Selection**: Instant equipment selection with immediate command generation
- **🌐 Dual Protocol Support**: Modbus RTU (serial) and Modbus TCP support
- **📊 Real-Time Logging**: Live command execution and response monitoring
- **🔍 Auto-Detection**: Automatic COM port detection and validation
- **📝 Custom Commands**: Support for custom modpoll command arguments
- **🎨 Professional Interface**: Tabbed interface with Basic and Advanced modes

## 📥 Download

### Latest Release
Download the latest executable directly from our releases:

**[Download ModPollingTool.exe](https://github.com/spenz91/ModpollingTool/releases/download/modpollv2/ModPollingTool.exe)**

- **File Size**: ~10.6 MB
- **Platform**: Windows 10/11 (64-bit)
- **Dependencies**: None (standalone executable)

## 🚀 Installation

1. **Download** the `ModPollingTool.exe` file from the link above
2. **Extract** or place the file in your desired directory
3. **Run** the executable by double-clicking `ModPollingTool.exe`
4. **No installation required** - it's a portable application

## 🔧 How It Works

### Core Architecture

The ModPolling Tool is built around the modpoll utility and provides a graphical wrapper that:

1. **Generates Commands**: Automatically builds modpoll command-line arguments based on user selections
2. **Manages Communication**: Handles both serial (RTU) and TCP Modbus protocols
3. **Provides Feedback**: Real-time logging of commands and responses
4. **Simplifies Configuration**: Pre-configured equipment presets for common devices

### Key Components

- **Equipment Selection**: Choose from predefined device configurations
- **Communication Settings**: Configure COM ports, baud rates, parity, etc.
- **Command Generation**: Automatic modpoll command building
- **Execution Engine**: Subprocess management for modpoll execution
- **Logging System**: Real-time command and response monitoring

## 📱 Equipment Presets

The tool includes pre-configured settings for popular industrial devices

## 📖 Usage Guide

### Quick Start

1. **Launch** the ModPollingTool.exe
2. **Select Equipment**: Click on your device from the equipment list
3. **Configure Settings**: Adjust COM port, slave address, and register settings
4. **Start Polling**: Click "Start Polling" to begin communication

### Basic Mode

1. **Equipment Tab**: Select your device from the predefined list
2. **Basic Tab**: Configure essential parameters:
   - COM Port selection
   - Slave Address
   - Start Reference
   - Number of Values
   - Register Data Type
3. **Execute**: Click "Start Polling" to run the command

### Advanced Mode

1. **Custom Commands**: Enter custom modpoll arguments
2. **Manual Configuration**: Set all parameters manually
3. **Command Preview**: See the generated command before execution

### Features

- **Single-Click Equipment Selection**: Click once on any equipment to apply its preset
- **Immediate Command Display**: See the full modpoll command instantly
- **Real-Time Logging**: Monitor command execution and responses
- **Error Handling**: Comprehensive error reporting and validation

## 📋 Requirements

### System Requirements
- **OS**: Windows 10 or Windows 11 (64-bit)
- **RAM**: 4 GB minimum (8 GB recommended)
- **Storage**: 50 MB free space
- **Network**: Required for Modbus TCP communication

### Software Dependencies
- **modpoll.exe**: Embedded in the application bundle
- **No Python Required**: Standalone executable
- **No Additional Libraries**: Self-contained application

## 🔨 Building from Source

If you want to build the application from source:

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Build Steps

1. **Clone the repository**:
   ```bash
   git clone https://github.com/spenz91/ModpollingTool.git
   cd ModpollingTool
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Build executable**:
   ```bash
   python build_exe.py
   ```

4. **Find the executable** in the `dist/` folder

### Build Files
- `modpoll.py`: Main application source code
- `requirements.txt`: Python dependencies
- `build_exe.py`: PyInstaller build script
- `build.bat`: Windows batch build script

## 🤝 Contributing

We welcome contributions! Here's how you can help:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/AmazingFeature`)
3. **Commit** your changes (`git commit -m 'Add some AmazingFeature'`)
4. **Push** to the branch (`git push origin feature/AmazingFeature`)
5. **Open** a Pull Request

### Development Guidelines
- Follow PEP 8 style guidelines
- Add comments for complex logic
- Test thoroughly before submitting
- Update documentation as needed

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **modpoll utility**: The underlying Modbus communication tool
- **Carlo Gavazzi**: Equipment specifications and testing
- **SLV**: Device configuration support
- **tkinter**: GUI framework
- **PyInstaller**: Executable packaging

## 🔄 Version History

- **v2.0**: Professional GUI with equipment presets
- **v1.0**: Initial command-line version

---

**Made with ❤️ for the industrial automation community**
