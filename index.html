<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Google Apps Script Executor</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            line-height: 1.6;
        }

        h1 {
            color: #4285f4;
            margin-bottom: 20px;
        }

        .container {
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }

        textarea {
            width: 100%;
            height: 150px;
            padding: 10px;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-family: monospace;
            margin-bottom: 10px;
        }

        input[type="text"] {
            width: 100%;
            padding: 8px;
            margin-bottom: 10px;
            border: 1px solid #ccc;
            border-radius: 4px;
        }

        button {
            background-color: #4285f4;
            color: white;
            border: none;
            padding: 10px 15px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }

        button:hover {
            background-color: #3367d6;
        }

        button:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }

        #result {
            background-color: #f9f9f9;
            padding: 15px;
            border-radius: 4px;
            white-space: pre-wrap;
            overflow-x: auto;
            min-height: 100px;
        }

        .auth-container {
            margin-bottom: 20px;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 20px;
        }

        .hidden {
            display: none;
        }

        #auth-status {
            margin-bottom: 15px;
            font-weight: bold;
        }

        .mode-selector {
            margin-bottom: 20px;
        }

        .required-field::after {
            content: " *";
            color: red;
        }

        .error-message {
            color: red;
            font-size: 14px;
            margin-top: -8px;
            margin-bottom: 10px;
            display: none;
        }
    </style>
</head>

<body>
    <h1>Google Apps Script Executor</h1>

    <div class="auth-container">
        <div id="auth-status">Not authenticated</div>
        <h2 class="required-field">OAuth Client ID</h2>
        <input type="text" id="client-id" placeholder="Enter your Google OAuth Client ID">
        <div id="client-id-error" class="error-message">Este campo es obligatorio</div>
        <button id="auth-button">Authenticate</button>
        <button id="signout-button" class="hidden">Sign Out</button>
    </div>

    <div class="container hidden" id="executor-container">
        <div class="mode-selector">
            <input type="radio" id="use-existing" name="mode" value="existing" checked>
            <label for="use-existing">Use Existing Script</label>
            
            <input type="radio" id="update-existing" name="mode" value="update" disabled>
            <label for="update-existing">Update Existing Script</label>
        </div>

        <div id="existing-script-container">
            <h2>Script ID (Prefilled)</h2>
            <input type="text" id="script-id" value="12RBJMvqQePFAxr498N3i8u7U2DsixdFijUwrKX6zf_QU8tp4yq4o6Ftf" readonly>
        </div>

        <div id="new-script-container" class="hidden">
            <h2>Script Code</h2>
            <textarea id="script-code" placeholder="Enter your Google Apps Script code here">function runThis() {
                Logger.log('Hello from dynamic script!');
                return 'Script executed successfully!';
                }
            </textarea>
        </div>

        <div id="update-script-container" class="hidden">
            <h2>Script ID (Prefilled)</h2>
            <input type="text" id="update-script-id" value="12RBJMvqQePFAxr498N3i8u7U2DsixdFijUwrKX6zf_QU8tp4yq4o6Ftf" readonly>
            <h2>Script Code</h2>
            <textarea id="update-script-code" placeholder="Enter your Google Apps Script code here">function runThis() {
                Logger.log('Hello from updated script!');
                return 'Script updated and executed successfully!';
                }
            </textarea>
        </div>

        <h2 class="required-field">Function to Execute</h2>
        <input type="text" id="function-name" placeholder="Function name to execute" value="runThis">
        <div id="function-name-error" class="error-message">Este campo es obligatorio</div>

        <h2 class="required-field">Function Parameters</h2>
        <textarea id="function-params" placeholder="Enter parameters as JSON array (e.g., ['param1', 'param2'])"></textarea>
        <div id="function-params-error" class="error-message">Los parámetros son obligatorios. Ingresa un array JSON válido.</div>

        <button id="execute-button">Execute Script</button>

        <h2>Result</h2>
        <div id="result">Results will appear here</div>
    </div>

    <!-- Load the Google API client library for REST API calls -->
    <script src="https://apis.google.com/js/api.js"></script>
    <!-- Load the newer Google Identity Services library -->
    <script src="https://accounts.google.com/gsi/client"></script>

    <script>
        let tokenClient;
        let accessToken = '';
        const SCRIPT_ID = '12RBJMvqQePFAxr498N3i8u7U2DsixdFijUwrKX6zf_QU8tp4yq4o6Ftf';
        const SCOPES = 'https://www.googleapis.com/auth/script.projects https://www.googleapis.com/auth/script.deployments https://www.googleapis.com/auth/script.external_request https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/forms';
        const API_KEY = ''; // Optional for this use case

        // Initialize the page
        document.addEventListener('DOMContentLoaded', initializePage);

        function initializePage() {
            // Initialize the Google API client library
            gapi.load('client', initGapiClient);

            // Set up button click handlers
            document.getElementById('auth-button').addEventListener('click', requestAccessToken);
            document.getElementById('signout-button').addEventListener('click', handleSignOut);
            document.getElementById('execute-button').addEventListener('click', executeScript);

            // Set up mode selector
            document.querySelectorAll('input[name="mode"]').forEach(radio => {
                radio.addEventListener('change', updateMode);
            });

            // Add validation listeners
            document.getElementById('client-id').addEventListener('input', validateClientId);
            document.getElementById('function-name').addEventListener('input', validateFunctionName);
            document.getElementById('function-params').addEventListener('input', validateFunctionParams);
        }

        // Validate client ID
        function validateClientId() {
            const clientId = document.getElementById('client-id').value.trim();
            const errorElement = document.getElementById('client-id-error');
            const authButton = document.getElementById('auth-button');
            
            if (!clientId) {
                errorElement.style.display = 'block';
                authButton.disabled = true;
                return false;
            } else {
                errorElement.style.display = 'none';
                authButton.disabled = false;
                return true;
            }
        }

        // Validate function name
        function validateFunctionName() {
            const functionName = document.getElementById('function-name').value.trim();
            const errorElement = document.getElementById('function-name-error');
            
            if (!functionName) {
                errorElement.style.display = 'block';
                return false;
            } else {
                errorElement.style.display = 'none';
                return true;
            }
        }

        // Validate function parameters
        function validateFunctionParams() {
            const paramsInput = document.getElementById('function-params').value.trim();
            const errorElement = document.getElementById('function-params-error');
            
            if (!paramsInput) {
                errorElement.style.display = 'block';
                return false;
            }
            
            try {
                const params = JSON.parse(paramsInput);
                if (!Array.isArray(params)) {
                    errorElement.style.display = 'block';
                    errorElement.textContent = 'Los parámetros deben ser un array JSON.';
                    return false;
                }
                errorElement.style.display = 'none';
                return true;
            } catch (e) {
                errorElement.style.display = 'block';
                errorElement.textContent = 'Error de formato JSON. Por favor ingrese un array JSON válido.';
                return false;
            }
        }

        // Validate all fields before execution
        function validateAllFields() {
            const functionNameValid = validateFunctionName();
            const functionParamsValid = validateFunctionParams();
            
            return functionNameValid && functionParamsValid;
        }

        // Update UI based on selected mode
        function updateMode() {
            const mode = document.querySelector('input[name="mode"]:checked').value;
            const existingContainer = document.getElementById('existing-script-container');
            const newContainer = document.getElementById('new-script-container');
            const updateContainer = document.getElementById('update-script-container');
            
            if (mode === 'existing') {
                existingContainer.classList.remove('hidden');
                newContainer.classList.add('hidden');
                updateContainer.classList.add('hidden');
            } else if (mode === 'update') {
                existingContainer.classList.add('hidden');
                newContainer.classList.add('hidden');
                updateContainer.classList.remove('hidden');
            }
        }

        // Initialize the GAPI client
        async function initGapiClient() {
            await gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: ['https://script.googleapis.com/$discovery/rest?version=v1']
            });

            // Check if we have a stored token and validate it
            const storedToken = localStorage.getItem('gis_access_token');
            if (storedToken) {
                try {
                    accessToken = storedToken;
                    gapi.client.setToken({ access_token: accessToken });
                    // Try a simple API call to verify token is still valid
                    await gapi.client.script.projects.list({
                        pageSize: 1
                    });
                    updateSigninStatus(true);
                } catch (error) {
                    // Token is invalid, clear it
                    accessToken = '';
                    localStorage.removeItem('gis_access_token');
                    updateSigninStatus(false);
                }
            } else {
                updateSigninStatus(false);
            }
        }

        // Handle the token response
        function handleTokenResponse(tokenResponse) {
            if (tokenResponse && tokenResponse.access_token) {
                accessToken = tokenResponse.access_token;
                // Store token for future use
                localStorage.setItem('gis_access_token', accessToken);
                // Set the token for API calls
                gapi.client.setToken({ access_token: accessToken });
                updateSigninStatus(true);
            } else {
                updateSigninStatus(false);
            }
        }

        // Update UI based on auth status
        function updateSigninStatus(isSignedIn) {
            const authStatus = document.getElementById('auth-status');
            const authButton = document.getElementById('auth-button');
            const signoutButton = document.getElementById('signout-button');
            const executorContainer = document.getElementById('executor-container');

            if (isSignedIn) {
                authStatus.textContent = 'Authenticated';
                authStatus.style.color = 'green';
                authButton.classList.add('hidden');
                signoutButton.classList.remove('hidden');
                executorContainer.classList.remove('hidden');
            } else {
                authStatus.textContent = 'Not authenticated';
                authButton.classList.remove('hidden');
                signoutButton.classList.add('hidden');
                executorContainer.classList.add('hidden');
            }
        }

        // Handle sign out
        function handleSignOut() {
            // Clear the token
            accessToken = '';
            localStorage.removeItem('gis_access_token');
            gapi.client.setToken(null);
            updateSigninStatus(false);
        }

        // Request access token
        function requestAccessToken() {
            if (!validateClientId()) {
                return;
            }

            const clientId = document.getElementById('client-id').value.trim();

            // Initialize token client with provided client ID
            tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: clientId,
                scope: SCOPES,
                callback: handleTokenResponse
            });
            
            tokenClient.requestAccessToken();
        }

        // Execute the script
        function executeScript() {
            // Check if authenticated
            if (!accessToken) {
                requestAccessToken();
                return;
            }

            // Validate all fields before proceeding
            if (!validateAllFields()) {
                return;
            }

            const resultDiv = document.getElementById('result');
            resultDiv.textContent = 'Executing script...';

            const mode = document.querySelector('input[name="mode"]:checked').value;
            const functionName = document.getElementById('function-name').value.trim();
            let functionParams = [];

            try {
                const paramsInput = document.getElementById('function-params').value.trim();
                functionParams = JSON.parse(paramsInput);
            } catch (e) {
                resultDiv.textContent = 'Error parsing parameters. Ensure they are in valid JSON array format.';
                return;
            }

            if (mode === 'existing') {
                executeExistingScript(SCRIPT_ID, functionName, functionParams, resultDiv);
            } else if (mode === 'update') {
                const scriptCode = document.getElementById('update-script-code').value;
                updateExistingScript(SCRIPT_ID, scriptCode, functionName, functionParams, resultDiv);
            }
        }

        // Execute an existing script
        function executeExistingScript(scriptId, functionName, functionParams, resultDiv) {
            resultDiv.textContent = `Executing function "${functionName}" from script ID "${scriptId}"...`;
            
            gapi.client.script.scripts.run({
                scriptId: scriptId,
                resource: {
                    function: functionName,
                    parameters: functionParams,
                    devMode: true
                }
            }).then(response => {
                if (response.result.error) {
                    resultDiv.textContent = `Error: ${JSON.stringify(response.result.error, null, 2)}`;
                } else {
                    let outputText = 'Execution successful!';
                    if (response.result.response && response.result.response.result) {
                        outputText += `\n\nResult: ${JSON.stringify(response.result.response.result, null, 2)}`;
                    }
                    resultDiv.textContent = outputText;
                }
            }).catch(error => {
                resultDiv.textContent = `Error: ${error.result ? JSON.stringify(error.result, null, 2) : error.message}`;
                // If the error is an auth error, attempt to refresh the token
                if (error.status === 401) {
                    requestAccessToken();
                }
            });
        }

        // Update an existing script and execute it
        function updateExistingScript(scriptId, scriptCode, functionName, functionParams, resultDiv) {
            resultDiv.textContent = `Updating and executing function "${functionName}" from script ID "${scriptId}"...`;
            
            // Step 1: Update script content
            gapi.client.script.projects.updateContent({
                scriptId: scriptId,
                resource: {
                    files: [
                        {
                            name: 'main',
                            type: 'SERVER_JS',
                            source: scriptCode
                        },
                        {
                            name: 'appsscript',
                            type: 'JSON',
                            source: JSON.stringify({
                                timeZone: 'America/New_York',
                                dependencies: {},
                                exceptionLogging: 'CLOUD'
                            })
                        }
                    ]
                }
            }).then(() => {
                resultDiv.textContent += "\nScript content updated successfully.";

                // Step 2: Create a new script version
                return gapi.client.script.projects.versions.create({
                    scriptId: scriptId
                });
            }).then(versionResponse => {
                const versionNumber = versionResponse.result.versionNumber;
                resultDiv.textContent += `\nNew version created: ${versionNumber}`;

                // Step 3: Deploy the new version
                return gapi.client.script.projects.deployments.create({
                    scriptId: scriptId,
                    resource: {
                        versionNumber: versionNumber,
                        manifestFileName: 'appsscript',
                        description: "Automated deployment"
                    }
                });
            }).then(deploymentResponse => {
                const deploymentId = deploymentResponse.result.deploymentId;
                resultDiv.textContent += `\nDeployed with ID: ${deploymentId}\nExecuting...`;

                // Step 4: Wait a moment for deployment to complete
                return new Promise(resolve => setTimeout(resolve, 5000))
                    .then(() => gapi.client.script.scripts.run({
                        scriptId: scriptId,
                        resource: {
                            function: functionName,
                            parameters: functionParams,
                            devMode: true
                        }
                    }));
            }).then(executionResponse => {
                if (executionResponse.result.error) {
                    resultDiv.textContent += `\nExecution Error: ${JSON.stringify(executionResponse.result.error, null, 2)}`;
                } else {
                    let outputText = '\nExecution successful!';
                    if (executionResponse.result.response && executionResponse.result.response.result) {
                        outputText += `\n\nResult: ${JSON.stringify(executionResponse.result.response.result, null, 2)}`;
                    }
                    resultDiv.textContent += outputText;
                }
            }).catch(error => {
                resultDiv.textContent += `\nError: ${error.result ? JSON.stringify(error.result, null, 2) : error.message}`;
            });
        }
    </script>
</body>

</html>