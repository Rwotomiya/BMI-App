# BMI Calculator

## Overview

This BMI Calculator is a simple desktop application built using Python, Tkinter for the GUI, Pillow for image handling, and OpenPyXL for saving data to an Excel file. The application allows users to input their height, weight, and age to calculate their Body Mass Index (BMI) and classify it into categories such as underweight, normal weight, overweight, and obese. The results are displayed in the application and saved to an Excel file.

## Features

- Calculate BMI based on user input.
- Classify BMI into categories.
- Save BMI results to an Excel file.
- Clear input fields and previous results.
- User-friendly interface with hover effects on buttons.

## Requirements

- Python 3.x
- Tkinter (included with Python)
- Pillow
- OpenPyXL

## Installation

1. **Clone the repository**:
    ```bash
    git clone https://github.com/yourusername/bmi-calculator.git
    cd bmi-calculator
    ```

2. **Install the required packages**:
    ```bash
    pip install pillow openpyxl
    ```

## Usage

1. **Run the application**:
    ```bash
    python bmi_calculator.py
    ```

2. **Input your data**:
    - Enter your first name, last name, height (in meters), weight (in kilograms), and age.

3. **Calculate BMI**:
    - Click the "Calculate" button to calculate your BMI and display the result.

4. **Clear entries**:
    - Click the "Clear" button to clear all input fields and previous results.

5. **Check the results**:
    - The BMI result and category will be displayed in the application.
    - The results will also be saved to an Excel file named `bmi_results.xlsx` in the same directory.

## File Structure

- `bmi_calculator.py`: The main Python script for the BMI Calculator application.
- `logo.png`: The icon image for the application window.
- `bmi_results.xlsx`: The Excel file where BMI results are saved.

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

## License

This project is licensed under the MIT License. Please take a look at the [LICENSE](LICENSE) file for details.

## Contact

If you have any questions or suggestions, please contact [daudivincent20@gmail.com](mailto:daudivincent20@gmail.com).

## Acknowledgments

- Thanks to Tkinter, Pillow, and OpenPyXL developers for providing the tools to build this application.

---

