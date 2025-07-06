# OpenAPI to Sheet (Testcase Format)

A Node.js tool to convert OpenAPI YAML specifications into Excel test case sheets, designed to help QA engineers generate API test cases efficiently.

## Features

- Parses OpenAPI 3.0.x YAML files.
- Extracts API endpoints, methods, parameters, request bodies, and responses.
- Generates detailed test cases for each API operation and response status.
- Outputs test cases in an Excel (.xlsx) file with multiple sheets for API summary, test cases, request details, and response details.

![image](https://github.com/user-attachments/assets/983b02d4-853f-4b0b-b340-5677e458fcbd)

## Installation

Make sure you have Node.js installed.

```bash
npm install
```

## Usage

Run the tool with the OpenAPI YAML file as input:

```bash
node main.js <openapi-file.yaml> [output-file.xlsx]
```

- `<openapi-file.yaml>`: Path to your OpenAPI specification file.
- `[output-file.xlsx]`: Optional output Excel file name (default: `api-test-cases.xlsx`).

Example:

```bash
node main.js openapi-sample.yaml testcases-api.xlsx
```

## Project Structure

- `main.js`: Main script that parses the OpenAPI file, generates test cases, and creates the Excel file.
- `openapi.yaml`, `openapi-sample.yaml`, `openapi-deprecated.yaml`: Sample OpenAPI specification files.
- `testcases-api.xlsx`: Example output Excel file with generated test cases.
- `package.json` and `package-lock.json`: Project dependencies and metadata.

## How It Works

1. Parses the OpenAPI YAML file using `js-yaml`.
2. Iterates through each API path and HTTP method.
3. Extracts parameters (header, path, query), request body schema, and response schemas.
4. Generates test case details including test name, scenario, steps, test data, and expected results.
5. Formats test data with valid and invalid examples for different response codes.
6. Creates an Excel file with organized sheets for easy QA review.

## Dependencies

- `js-yaml`: For parsing YAML files.
- `xlsx`: For creating Excel files.
- `fs` and `path`: Node.js core modules for file handling.

## Contributing

Contributions are welcome! Please open issues or submit pull requests for improvements.

## License

This project is licensed under the ISC License.

---

This tool is ideal for QA teams looking to automate the generation of API test cases from OpenAPI specifications, improving testing coverage and efficiency.
        
