# ExcelPlannerGenerator

Simple C# script that generates Excel planners.

## Features
- Generates an Excel planner with customizable dates, time slots, and note sections.
- Automatically formats headers, workdays, and weekends.
- Highlights past dates in red and today's date in blue.
- Fully customizable colors and styles.

## Requirements
- [.NET 6 or later](https://dotnet.microsoft.com/en-us/download)
- [EPPlus](https://www.nuget.org/packages/EPPlus) (for Excel file generation)

## Installation
1. Clone this repository:
   ```sh
   git clone https://github.com/radovicdanilo/ExcelPlannerGenerator.git
   cd ExcelPlannerGenerator
   ```

2. Install dependencies:
   ```sh
   dotnet add package EPPlus
   ```
3. Build project:
   ```sh
   dotnet build
   ```

## Usage
Run the script with the following command:

   ```sh
   dotnet run YYYY-MM-DD num_days
   ```

- `YYYY-MM-DD`: Start date of the planner.
- `num_days`: Number of days to generate.

## License
This project is licensed under the MIT License.
