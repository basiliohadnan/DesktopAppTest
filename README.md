# Desktop App Testing with Selenium and Appium

## Overview
This project aims to automate testing for a desktop application using Selenium and Appium. The application under test is a simple note-taking app.

## Setup Instructions
To set up the testing environment for this project, follow these steps:

1. **Install WinAppDriver**: 
   - Download WinAppDriver from [here](https://github.com/Microsoft/WinAppDriver/releases).
   - Install WinAppDriver on your machine.

2. **Clone the Repository**: 
   - Clone this repository to your local machine.

3. **Install Dependencies**:
   - Make sure you have Visual Studio installed.
   - Restore the NuGet packages for the solution.

4. **Configure Paths**: 
   - Update the `WinAppDriverPath` constant in the `NoteAppTest.cs` file to point to the location where WinAppDriver is installed.
   - Update the `appId` constant with the correct application ID.

## Running Tests
To run the tests, follow these steps:

1. **Start WinAppDriver**:
   - Before running the tests, make sure WinAppDriver is running. If not, start WinAppDriver.

2. **Build Solution**:
   - Build the solution in Visual Studio.

3. **Run Tests**:
   - Once the solution is built successfully, you can run the tests using Visual Studio's test runner.

## Test Cases
### Validates Text Inserted
- **Test Description**: Verifies that text entered into the note-taking app is correctly displayed.
- **Steps**:
  1. Insert the text into the app.
  2. Capture a screenshot.
  3. Extract text from the screenshot using OCR.
  4. Compare the extracted text with the expected result.
- **Expected Result**: The extracted text should match the expected text.

## Note
- Make sure to update the `ScreenshotsDirectory` constant in the `NoteAppTest.cs` file with the desired directory path where screenshots will be saved.

