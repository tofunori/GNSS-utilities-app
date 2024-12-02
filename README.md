# GNSS Processing Tools

A comprehensive toolkit for GNSS data processing and conversion, built with Python and Tkinter.

## Features

### 1. GNSS Data Viewer
- Import and parse .sum files
- Automatic handling of GNSS model-specific reference points (ARP/APC)
- Elevation adjustments for different GNSS models:
  - EMLID INREACH RS2 (L1 = 0.135m, L2 = 0.137m)
  - FOIF A30 (Auto APC reference)
- DMS to decimal degrees conversion
- Excel export functionality
- Data table view with sorting capabilities

### 2. POS to Excel Converter
- Batch processing of .pos files
- Preview data before conversion
- Export to Excel format
- Supports multiple data columns including:
  - Date and Time
  - Coordinates (Latitude, Longitude, Height)
  - Quality indicators
  - Standard deviations

### 3. DMS Converter
- Convert Degrees Minutes Seconds (DMS) to Decimal Degrees
- Format: XX° YY' ZZ.ZZZZZ"
- High precision output (9 decimal places)
- User-friendly interface

### 4. F16 to R27 Converter
- Batch conversion of .F16 files to .R27 format
- Source and destination folder selection
- Progress tracking
- Success confirmation

## Installation

### Prerequisites
```bash
pip install pandas
```

### Required Python Packages
- tkinter (usually comes with Python)
- pandas
- datetime
- matplotlib

### Running the Application
1. Clone the repository
2. Install required packages
3. Run the main script:
```bash
python main.py
```

## Usage

### GNSS Data Viewer
1. Launch the application
2. Click "GNSS Data Viewer"
3. Use "Import .sum File" to load data
4. Select appropriate GNSS model
5. Reference point will automatically adjust based on model:
   - EMLID INREACH RS2: Choose between ARP/APC
   - FOIF A30: Automatically sets to APC

### POS to Excel Converter
1. Select "POS to Excel Converter"
2. Choose input directory containing .pos files
3. Preview data
4. Select output location and save to Excel

### DMS Converter
1. Enter DMS value in format: 73° 9' 18.99435"
2. Click Convert
3. Get decimal degrees result

### F16 to R27 Converter
1. Select source folder containing .F16 files
2. Choose destination folder
3. Click Convert
4. Check success message

## File Formats

### .sum File Structure
The application expects .sum files with the following data:
- MKR: Date marker
- RNX: File information
- BEG: Start time
- END: End time
- INT: Interval
- POS LAT/LON: Position data
- PRJ TYPE: UTM projection data

### .pos File Structure
Expected format for position files:
- Column-based data
- Headers for coordinate information
- Quality indicators
- Standard deviation values

## Contributing
Feel free to submit issues and enhancement requests.

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments
- Built using Python's Tkinter library
- Utilizes pandas for data handling
- Implements GNSS industry standards for coordinate processing
