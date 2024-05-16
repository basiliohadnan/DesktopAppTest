# Desktop App Testing with Selenium and Appium

## Overview
This project aims to automate testing for a desktop application using Selenium and Appium.

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
   - Update the `appPath` constant with the correct application path.

## Running Tests
To run the tests, follow these steps:

1. **Build Solution**:
   - Build the solution in Visual Studio.

2. **Run Tests**:
   - Once the solution is built successfully, you can run the tests using Visual Studio's test runner.

## Note
- Make sure to update the `ScreenshotsDirectory` constant with the desired directory path where screenshots will be saved.

