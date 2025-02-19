/* global CustomFunctions */
// Create a shared key for localStorage
const WARNING_KEY = 'your_warning_key';
const CHECK_INTERVAL = 1000; // 1 second

// Global variables for batch handling
let _batch = [];
let _isBatchedRequestScheduled = false;

/**
 * Gets a indicator value for a dataset from the database
 * @customfunction
 * @param {string} name Name parameter
 * @param {string} indicator Indicator parameter
 * @param {string} modul Module parameter
 * @returns {string|number} Result value or error message
 */
function get(name, indicator, modul) {
  // Push the operation to batch and return promise
  return _pushOperation("get", [name, indicator, modul]);
}

/**
 * Helper function to add operations to the batch
 * @param {string} op Operation name
 * @param {any[]} args Operation arguments
 * @returns {Promise} Promise that resolves with the operation result
 */
function _pushOperation(op, args) {
  // Create an entry for the custom function
  const invocationEntry = {
    operation: op,
    args: args,
    resolve: undefined,
    reject: undefined
  };

  // Create a unique promise for this invocation
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch
  _batch.push(invocationEntry);

  // Schedule remote request if not already scheduled
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  return promise;
}

/**
 * Makes the remote request with batched operations
 */
async function _makeRemoteRequest() {
  try {
    // Copy and clear the batch
    const batchCopy = _batch.splice(0, _batch.length);
    _isBatchedRequestScheduled = false;

    // Build request batch with only necessary data
    const requestBatch = batchCopy.map(item => ({
      operation: item.operation,
      args: item.args
    }));

    // Make the remote request
    const response = await fetch('YOUR_API_ENDPOINT/batch', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ batch: requestBatch })
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const responseBatch = await response.json();

    // Process each response
    responseBatch.results.forEach((result, index) => {
      if (result.error) {
        batchCopy[index].reject(new Error(result.error));
      } else {
        // Handle warning status
        if (result.warning) {
          localStorage.setItem(WARNING_KEY, JSON.stringify({
            warning: true,
            timestamp: Date.now(),
            function: 'get'
          }));
        }

        // Resolve with appropriate value based on status
        if (result.status === 'limit_reached') {
          batchCopy[index].resolve('limit');
        } else if (result.status === 'free_token') {
          batchCopy[index].resolve('kostenlose version');
        } else {
          batchCopy[index].resolve(result.value);
        }
      }
    });
  } catch (error) {
    // Handle network errors
    if (error.message === "Failed to fetch" || !window.navigator.onLine) {
      batchCopy.forEach(entry => entry.resolve("keine verbindung zum server"));
    } else {
      batchCopy.forEach(entry => entry.resolve("kein wert vorhanden"));
    }
  }
  finally{
    await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
  }
}

/**
 * Gets general information about a dataset from the database
 * @customfunction 
 * @param {string} name Name parameter
 * @param {string} info Info parameter
 * @returns {string|number} Result value or error message
 */
async function info(name, info) {
  if (!name || !info) {
    return 'kein wert vorhanden';
  }
  try {
    const response = await fetch(`YOUR_API_ENDPOINT/info?name=${encodeURIComponent(name)}&info=${encodeURIComponent(info)}`);
    
    if (!response.ok) {
      return 'kein wert vorhanden';
    }
    
    const data = await response.json();
    
    if (data.warning) {
      localStorage.setItem(WARNING_KEY, JSON.stringify({
        warning: true,
        timestamp: Date.now(),
        function: 'info'
      }));
    }
    
    if (data.status === 'limit_reached') {
      return 'limit';
    } else if (data.status === 'free_token') {
      return 'kostenlose version';
    }
    
    return data.value || 'kein wert vorhanden';
    
  } catch (error) {
    console.error('Error in info function:', error);
    return 'kein wert vorhanden';
  }
  finally{
    await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
  }
}

/**
 * Gets the lifespan value for a dataset from the database
 * @customfunction
 * @param {string} name Name of the dataset
 * @returns {string|number} Lifespan value or error message
 */
async function life(name) {
  if (!name) {
    return 'kein datensatz gefunden';
  }
  try {
    const response = await fetch(`YOUR_API_ENDPOINT/lifespan/${encodeURIComponent(name)}`);
    
    if (!response.ok) {
      return 'kein datensatz gefunden';
    }
    
    const data = await response.json();
    
    if (data.warning) {
      localStorage.setItem(WARNING_KEY, JSON.stringify({
        warning: true,
        timestamp: Date.now(),
        function: 'life'
      }));
    }
    
    if (data.status === 'limit_reached') {
      return 'limit';
    } else if (data.status === 'free_token') {
      return 'kostenlose version';
    }
    
    return data.value || 'kein datensatz gefunden';
    
  } catch (error) {
    console.error('Error in life function:', error);
    return 'kein datensatz gefunden';
  }
  finally{
    await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
  } 
}

/**
 * Finds a dataset value for a material/ component
 * @customfunction
 * @param {string} name Name of the material/ component
 * @returns {string|number} dataset value or error message
 */
async function data(name) {
  if (!name) {
    return 'kein datensatz gefunden';
  }
  try {
    const response = await fetch(`YOUR_API_ENDPOINT/material-match/${encodeURIComponent(name)}`);
    
    if (!response.ok) {
      return 'kein datensatz gefunden';
    }
    
    const data = await response.json();
    
    if (data.warning) {
      localStorage.setItem(WARNING_KEY, JSON.stringify({
        warning: true,
        timestamp: Date.now(),
        function: 'data'
      }));
    }
    
    if (data.status === 'limit_reached') {
      return 'limit';
    } else if (data.status === 'free_token') {
      return 'kostenlose version';
    }
    
    if (data.match && data.match.material) {
      return data.match.material;
    }
    
    return 'kein datensatz gefunden';
    
  } catch (error) {
    console.error('Error in data function:', error);
    return 'kein datensatz gefunden';
  }
  finally{
    await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
  }
}

// Initialize Office
Office.onReady((context) => {
  if (context.host === Office.HostType.Excel) {
    try { 
      if (!CustomFunctions._namespace) {
        console.warn("Namespace missing! Setting manually...");
        CustomFunctions._namespace = "YOUR_NAMESPACE";
      }

      CustomFunctions.associate('get', get);
      CustomFunctions.associate('info', info);
      CustomFunctions.associate('life', life);
      CustomFunctions.associate('data', data);
      console.log('Excel functions registered');

      const registeredFunctions = Object.keys(CustomFunctions._association);
      console.log("Registered functions:", registeredFunctions);
      console.log("Current namespace:", CustomFunctions._namespace);

    } catch (error) {
      console.error('Error registering functions:', error);
    }
  } 
});