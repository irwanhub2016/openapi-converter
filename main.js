// OpenAPI to Excel Test Case Generator
// This script converts an OpenAPI YAML specification to an Excel file with test cases
// for each endpoint, including different response scenarios

const fs = require('fs');
const yaml = require('js-yaml');
const XLSX = require('xlsx');
const path = require('path');

// Function to parse OpenAPI YAML file
function parseOpenApiYaml(filePath) {
  try {
    const fileContent = fs.readFileSync(filePath, 'utf8');
    return yaml.load(fileContent);
  } catch (error) {
    console.error(`Error reading or parsing YAML file: ${error.message}`);
    process.exit(1);
  }
}

// Function to generate test cases from OpenAPI spec
function generateTestCases(openApiSpec) {
  const testCases = [];
  const paths = openApiSpec.paths || {};
  
  // Process each path and method
  Object.keys(paths).forEach(path => {
    const pathItem = paths[path];
    
    // Process each HTTP method (GET, POST, PUT, DELETE, etc.)
    Object.keys(pathItem).forEach(method => {
      if (['get', 'post', 'put', 'patch', 'delete'].includes(method)) {
        const operation = pathItem[method];
        const operationId = operation.operationId || `${method.toUpperCase()} ${path}`;
        const summary = operation.summary || '';
        const description = operation.description || '';
        
        // Get request parameters
        const parameters = operation.parameters || [];
        const requestBody = operation.requestBody || {};
        
        // Build request information
        const requestInfo = {
          headers: [],
          pathParams: [],
          queryParams: [],
          requestBody: null
        };
        
        // Process parameters
        parameters.forEach(param => {
          const paramInfo = {
            name: param.name,
            required: param.required || false,
            type: param.schema?.type || 'string',
            example: param.schema?.example || ''
          };
          
          if (param.in === 'header') {
            requestInfo.headers.push(paramInfo);
          } else if (param.in === 'path') {
            requestInfo.pathParams.push(paramInfo);
          } else if (param.in === 'query') {
            requestInfo.queryParams.push(paramInfo);
          }
        });
        
        // Process request body
        if (requestBody.content && requestBody.content['application/json']?.schema) {
          const schema = requestBody.content['application/json'].schema;
          const required = schema.required || [];
          
          requestInfo.requestBody = {
            properties: {},
            required: required
          };
          
          if (schema.properties) {
            Object.keys(schema.properties).forEach(propName => {
              const prop = schema.properties[propName];
              requestInfo.requestBody.properties[propName] = {
                type: prop.type || 'string',
                example: prop.example || '',
                required: required.includes(propName)
              };
            });
          } else if (schema.$ref) {
            // Handle reference to a schema
            const refName = schema.$ref.split('/').pop();
            const refSchema = openApiSpec.components.schemas[refName];
            if (refSchema) {
              const refRequired = refSchema.required || [];
              
              requestInfo.requestBody.properties = {};
              requestInfo.requestBody.required = refRequired;
              
              if (refSchema.properties) {
                Object.keys(refSchema.properties).forEach(propName => {
                  const prop = refSchema.properties[propName];
                  requestInfo.requestBody.properties[propName] = {
                    type: prop.type || 'string',
                    example: prop.example || '',
                    required: refRequired.includes(propName)
                  };
                });
              }
            }
          }
        }
        
        // Get responses
        const responses = operation.responses || {};
        
        // Generate test cases for each response status code
        Object.keys(responses).forEach(statusCode => {
          const response = responses[statusCode];
          
          // Create test case
          const testCase = {
            endpoint: path,
            method: method.toUpperCase(),
            operationId,
            summary,
            description,
            statusCode,
            responseDescription: response.description || '',
            requestHeaders: JSON.stringify(requestInfo.headers.length > 0 ? requestInfo.headers : {}),
            pathParams: JSON.stringify(requestInfo.pathParams.length > 0 ? requestInfo.pathParams : {}),
            queryParams: JSON.stringify(requestInfo.queryParams.length > 0 ? requestInfo.queryParams : {}),
            requestBody: JSON.stringify(requestInfo.requestBody || {}),
            expectedResponse: extractResponseSchema(response, openApiSpec),
            testName: `${operationId} - ${statusCode} - ${response.description}`,
            testScenario: generateTestScenario(method, statusCode, operation.summary, response.description),
            testSteps: generateTestSteps(method, path, statusCode),
            testData: formatTestData(requestInfo, statusCode),
            expectedResult: `Should return ${statusCode} with ${response.description}`
          };
          
          testCases.push(testCase);
        });
      }
    });
  });
  
  return testCases;
}

// Function to extract response schema
function extractResponseSchema(response, openApiSpec) {
  if (!response.content || !response.content['application/json']?.schema) {
    return '{}';
  }
  
  const schema = response.content['application/json'].schema;
  let schemaContent = {};
  
  // Direct schema
  if (schema.properties) {
    schemaContent = { properties: schema.properties };
  }
  // Reference schema
  else if (schema.$ref) {
    const refName = schema.$ref.split('/').pop();
    const refSchema = openApiSpec.components.schemas[refName];
    if (refSchema) {
      schemaContent = refSchema;
    }
  }
  
  // Try to include example if available
  const example = response.content['application/json'].example;
  if (example) {
    schemaContent.example = example;
  }
  
  return JSON.stringify(schemaContent);
}

// Function to generate test scenario
function generateTestScenario(method, statusCode, summary, responseDescription) {
  const statusNum = parseInt(statusCode);
  
  if (statusNum >= 200 && statusNum < 300) {
    return `Verify successful ${method.toUpperCase()} request - ${summary}`;
  } else if (statusNum >= 400 && statusNum < 500) {
    if (statusNum === 401) {
      return `Verify authentication failure case`;
    } else if (statusNum === 403) {
      return `Verify authorization failure case`;
    } else if (statusNum === 404) {
      return `Verify not found case`;
    } else {
      return `Verify client error case - ${responseDescription}`;
    }
  } else if (statusNum >= 500) {
    return `Verify server error handling - ${responseDescription}`;
  }
  
  return `Verify ${method.toUpperCase()} request with ${statusCode} response`;
}

// Function to generate test steps
function generateTestSteps(method, path, statusCode) {
  const statusNum = parseInt(statusCode);
  const steps = [];
  
  steps.push(`1. Prepare the ${method.toUpperCase()} request to ${path}`);
  
  if (method.toLowerCase() === 'get') {
    steps.push(`2. Set required headers and parameters`);
  } else {
    steps.push(`2. Set required headers and body parameters`);
  }
  
  if (statusNum >= 400 && statusNum < 500) {
    if (statusNum === 401) {
      steps.push(`3. Send request with invalid or missing authentication`);
    } else if (statusNum === 403) {
      steps.push(`3. Send request with insufficient permissions`);
    } else if (statusNum === 404) {
      steps.push(`3. Send request with non-existent resource ID`);
    } else {
      steps.push(`3. Send request with invalid data`);
    }
  } else {
    steps.push(`3. Send the request`);
  }
  
  steps.push(`4. Verify the response status code is ${statusCode}`);
  steps.push(`5. Validate the response structure matches the schema`);
  
  if (statusNum >= 200 && statusNum < 300) {
    steps.push(`6. Validate the business logic of the response`);
  }
  
  return steps.join('\n');
}

// Function to format test data from request info
function formatTestData(requestInfo, statusCode) {
  const statusNum = parseInt(statusCode);
  let testData = {};
  
  // Handle headers
  if (requestInfo.headers.length > 0) {
    testData.headers = {};
    requestInfo.headers.forEach(header => {
      // For 401 error test cases, deliberately make auth headers invalid
      if ((header.name.toLowerCase().includes('auth') || header.name.toLowerCase().includes('api-key')) && statusNum === 401) {
        testData.headers[header.name] = 'INVALID_AUTH_TOKEN';
      } else {
        testData.headers[header.name] = header.example || `{${header.name}}`;
      }
    });
  }
  
  // Handle path parameters
  if (requestInfo.pathParams.length > 0) {
    testData.pathParams = {};
    requestInfo.pathParams.forEach(param => {
      // For 404 error test cases, use non-existent ID
      if (param.name.toLowerCase().includes('id') && statusNum === 404) {
        testData.pathParams[param.name] = 'NON_EXISTENT_ID';
      } else {
        testData.pathParams[param.name] = param.example || `{${param.name}}`;
      }
    });
  }
  
  // Handle query parameters
  if (requestInfo.queryParams.length > 0) {
    testData.queryParams = {};
    requestInfo.queryParams.forEach(param => {
      testData.queryParams[param.name] = param.example || `{${param.name}}`;
    });
  }
  
  // Handle request body
  if (requestInfo.requestBody && requestInfo.requestBody.properties) {
    testData.body = {};
    Object.keys(requestInfo.requestBody.properties).forEach(propName => {
      const prop = requestInfo.requestBody.properties[propName];
      
      // For 400 error test cases, make required fields missing or invalid
      if (prop.required && statusNum === 400) {
        if (Math.random() > 0.5) {
          // Missing field
          // Don't add this field to testData.body
        } else {
          // Invalid field value
          testData.body[propName] = getInvalidValueForType(prop.type);
        }
      } else {
        testData.body[propName] = prop.example || getDefaultValueForType(prop.type);
      }
    });
  }
  
  return JSON.stringify(testData);
}

// Function to get invalid value for a type (for error test cases)
function getInvalidValueForType(type) {
  switch (type) {
    case 'string':
      return 123; // Number instead of string
    case 'integer':
    case 'number':
      return 'not-a-number'; // String instead of number
    case 'boolean':
      return 'not-a-boolean'; // String instead of boolean
    case 'array':
      return {}; // Object instead of array
    case 'object':
      return []; // Array instead of object
    default:
      return null;
  }
}

// Function to get default value for a type
function getDefaultValueForType(type) {
  switch (type) {
    case 'string':
      return 'sample_string';
    case 'integer':
      return 1;
    case 'number':
      return 1.0;
    case 'boolean':
      return true;
    case 'array':
      return [];
    case 'object':
      return {};
    default:
      return null;
  }
}

// Function to create Excel file from test cases
function createExcelFile(testCases, outputFile) {
  try {
    // Create workbook and worksheet
    const workbook = XLSX.utils.book_new();
    
    // API Summary sheet
    const apiSummary = testCases.map(tc => ({
      'Endpoint': tc.endpoint,
      'Method': tc.method,
      'Operation ID': tc.operationId,
      'Summary': tc.summary,
      'Status Codes': tc.statusCode
    }));
    
    // Remove duplicates from API summary (keep one entry per endpoint+method)
    const uniqueApis = {};
    apiSummary.forEach(api => {
      const key = `${api.Method}:${api.Endpoint}`;
      if (!uniqueApis[key]) {
        uniqueApis[key] = api;
      } else {
        uniqueApis[key]['Status Codes'] += `, ${api['Status Codes']}`;
      }
    });
    
    const apiSummarySheet = XLSX.utils.json_to_sheet(Object.values(uniqueApis));
    XLSX.utils.book_append_sheet(workbook, apiSummarySheet, 'API Summary');
    
    // Test Cases sheet
    const testCasesSheet = XLSX.utils.json_to_sheet(testCases.map(tc => ({
      'Test ID': `TC_${tc.operationId}_${tc.statusCode}`.replace(/[^a-zA-Z0-9_]/g, '_'),
      'Test Name': tc.testName,
      'Endpoint': tc.endpoint,
      'Method': tc.method,
      'Test Scenario': tc.testScenario,
      'Test Steps': tc.testSteps,
      'Status Code': tc.statusCode,
      'Expected Result': tc.expectedResult
    })));
    XLSX.utils.book_append_sheet(workbook, testCasesSheet, 'Test Cases');
    
    // Request Details sheet
    const requestDetailsSheet = XLSX.utils.json_to_sheet(testCases.map(tc => ({
      'Test ID': `TC_${tc.operationId}_${tc.statusCode}`.replace(/[^a-zA-Z0-9_]/g, '_'),
      'Endpoint': tc.endpoint,
      'Method': tc.method,
      'Headers': tc.requestHeaders,
      'Path Parameters': tc.pathParams,
      'Query Parameters': tc.queryParams,
      'Request Body': tc.requestBody,
      'Test Data': tc.testData
    })));
    XLSX.utils.book_append_sheet(workbook, requestDetailsSheet, 'Request Details');
    
    // Response Details sheet
    const responseDetailsSheet = XLSX.utils.json_to_sheet(testCases.map(tc => ({
      'Test ID': `TC_${tc.operationId}_${tc.statusCode}`.replace(/[^a-zA-Z0-9_]/g, '_'),
      'Endpoint': tc.endpoint,
      'Method': tc.method,
      'Status Code': tc.statusCode,
      'Response Description': tc.responseDescription,
      'Expected Response Schema': tc.expectedResponse
    })));
    XLSX.utils.book_append_sheet(workbook, responseDetailsSheet, 'Response Details');
    
    // Write file
    XLSX.writeFile(workbook, outputFile);
    console.log(`Excel file created successfully: ${outputFile}`);
  } catch (error) {
    console.error(`Error creating Excel file: ${error.message}`);
    process.exit(1);
  }
}

// Main function
function main() {
  // Check command line arguments
  if (process.argv.length < 3) {
    console.error('Usage: node main.js <openapi-file.yaml> [output-file.xlsx]');
    process.exit(1);
  }
  
  const inputFile = process.argv[2];
  const outputFile = process.argv[3] || 'api-test-cases.xlsx';
  
  console.log(`Converting OpenAPI file: ${inputFile}`);
  console.log(`Output will be saved to: ${outputFile}`);
  
  // Parse OpenAPI YAML
  const openApiSpec = parseOpenApiYaml(inputFile);
  
  // Generate test cases
  const testCases = generateTestCases(openApiSpec);
  
  // Create Excel file
  createExcelFile(testCases, outputFile);
}

// Run the script
main();