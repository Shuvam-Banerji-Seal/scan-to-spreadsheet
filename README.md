# QRScannerApp

QRScannerApp is a simple desktop application built using Python's Tkinter library. It allows users to scan QR codes using their webcam and save the scanned data to a linked CSV spreadsheet.

## Features

- **QR Code Scanning**: Utilizes the webcam to scan QR codes in real-time.
- **Spreadsheet Linking**: Allows users to link a CSV file where the scanned QR code data will be saved.
- **Real-Time Feedback**: Provides immediate feedback on the success or failure of each operation.

## Installation

### Using the Source Code

1. **Clone the repository**:

    ```bash
    git clone https://github.com/yourusername/QRScannerApp.git
    cd QRScannerApp
    ```

2. **Install the required libraries**:

    Ensure you have Python 3.x installed. Then, install the necessary dependencies:

    ```bash
    pip install opencv-python pyzbar
    ```

3. **Run the application**:

    ```bash
    python qr_scanner_app.py
    ```

### Using the Executable File

For convenience, an executable file named `scan.exe` is also provided. You can use the application without needing to install Python or any dependencies.

1. **Download the executable file**:

    Download `scan.exe` from the [releases page](https://github.com/yourusername/QRScannerApp/releases).

2. **Run the application**:

    Double-click on `scan.exe` to start the application.

## Usage

1. **Start the Application**:
   - **Using Python script**: Run the application by executing `qr_scanner_app.py` script.

    ```bash
    python qr_scanner_app.py
    ```

   - **Using Executable**: Double-click on `scan.exe`.

2. **Link a Spreadsheet**:
   Click the "Link Spreadsheet" button to choose a CSV file where the scanned QR code data will be saved.

3. **Scan QR Codes**:
   Click the "Scan" button to start scanning QR codes using your webcam. The scanned data will be added to the linked CSV file.

## Code Overview

The application consists of a single class `QRScannerApp`:

- **`__init__(self, root)`**: Initializes the main application window and buttons.
- **`scan_qr(self)`**: Captures video from the webcam, scans for QR codes, and saves the data to the linked spreadsheet.
- **`link_spreadsheet(self)`**: Opens a file dialog to select and link a CSV spreadsheet.
- **`add_to_spreadsheet(self, data)`**: Adds the scanned QR code data to the linked CSV file.

## Dependencies

- `tkinter`: Standard Python interface to the Tk GUI toolkit.
- `opencv-python`: Library for real-time computer vision.
- `pyzbar`: Library for decoding barcodes and QR codes.
- `csv`: Module for handling CSV files.

## Future Improvements

- **Error Handling**: Improve error handling for various edge cases.
- **Multiple QR Codes**: Add support for handling multiple QR codes in a single frame.
- **GUI Enhancements**: Improve the user interface and experience.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your improvements.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

---

Feel free to customize this README further based on your specific needs and preferences.
