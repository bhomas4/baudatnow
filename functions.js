/* global CustomFunctions, Office, OfficeRuntime */

/**
 * Configuration constants
 * @type {Object}
 */
const CONFIG = {
    WARNING_KEY: 'your_warning_key',
    CHECK_INTERVAL: 1000,
    API_BASE_URL: 'YOUR_API_ENDPOINT',
    ERROR_MESSAGES: {
      NO_VALUE: 'kein wert vorhanden',
      NO_DATASET: 'kein datensatz gefunden',
      NO_CONNECTION: 'keine verbindung zum server',
      FREE_VERSION: 'kostenlose version',
      LIMIT_REACHED: 'limit'
    }
  };
  
  // Batch processing state
  let _batch = [];
  let _isBatchedRequestScheduled = false;
  
  /**
   * Gets an indicator value for a dataset from the database
   * @customfunction
   * @param {string} name - Dataset identifier
   * @param {string} indicator - Type of indicator to retrieve
   * @param {string} modul - Module specification
   * @returns {Promise<string|number>} Result value or error message
   */
  function get(name, indicator, modul) {
    return _pushOperation("get", [name, indicator, modul]);
  }
  
  /**
   * Adds an operation to the batch queue
   * @private
   * @param {string} op - Operation name
   * @param {Array<any>} args - Operation arguments
   * @returns {Promise<any>} Promise that resolves with operation result
   */
  function _pushOperation(op, args) {
    const invocationEntry = {
      operation: op,
      args: args,
      resolve: undefined,
      reject: undefined
    };
  
    const promise = new Promise((resolve, reject) => {
      invocationEntry.resolve = resolve;
      invocationEntry.reject = reject;
    });
  
    _batch.push(invocationEntry);
  
    if (!_isBatchedRequestScheduled) {
      _isBatchedRequestScheduled = true;
      setTimeout(_makeRemoteRequest, 100);
    }
  
    return promise;
  }
  
  /**
   * Stores a warning status in localStorage
   * @private
   * @param {string} functionName - Name of the function generating the warning
   */
  function _setWarningStatus(functionName) {
    localStorage.setItem(CONFIG.WARNING_KEY, JSON.stringify({
      warning: true,
      timestamp: Date.now(),
      function: functionName
    }));
  }
  
  /**
   * Makes the remote request with batched operations
   * @private
   */
  async function _makeRemoteRequest() {
    try {
      const batchCopy = _batch.splice(0, _batch.length);
      _isBatchedRequestScheduled = false;
  
      const requestBatch = batchCopy.map(item => ({
        operation: item.operation,
        args: item.args
      }));
  
      const response = await fetch(`${CONFIG.API_BASE_URL}/batch`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ batch: requestBatch })
      });
  
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
  
      const responseBatch = await response.json();
  
      responseBatch.results.forEach((result, index) => {
        if (result.error) {
          batchCopy[index].reject(new Error(result.error));
        } else {
          if (result.warning) {
            _setWarningStatus('get');
          }
  
          let resolveValue = result.value;
          if (result.status === 'limit_reached') {
            resolveValue = CONFIG.ERROR_MESSAGES.LIMIT_REACHED;
          } else if (result.status === 'free_token') {
            resolveValue = CONFIG.ERROR_MESSAGES.FREE_VERSION;
          }
  
          batchCopy[index].resolve(resolveValue);
        }
      });
    } catch (error) {
      const errorMessage = (!window.navigator.onLine || error.message === "Failed to fetch")
        ? CONFIG.ERROR_MESSAGES.NO_CONNECTION
        : CONFIG.ERROR_MESSAGES.NO_VALUE;
      
      batchCopy.forEach(entry => entry.resolve(errorMessage));
    } finally {
      await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
    }
  }
  
  /**
   * Gets information about a dataset property
   * @customfunction
   * @param {string} name - Dataset identifier
   * @param {string} info - Information type to retrieve
   * @returns {Promise<string|number>} Requested information or error message
   */
  async function info(name, info) {
    if (!name || !info) {
      return CONFIG.ERROR_MESSAGES.NO_VALUE;
    }
  
    try {
      const response = await fetch(
        `${CONFIG.API_BASE_URL}/info?name=${encodeURIComponent(name)}&info=${encodeURIComponent(info)}`
      );
      
      if (!response.ok) {
        return CONFIG.ERROR_MESSAGES.NO_VALUE;
      }
      
      const data = await response.json();
      
      if (data.warning) {
        _setWarningStatus('info');
      }
      
      if (data.status === 'limit_reached') {
        return CONFIG.ERROR_MESSAGES.LIMIT_REACHED;
      } else if (data.status === 'free_token') {
        return CONFIG.ERROR_MESSAGES.FREE_VERSION;
      }
      
      return data.value || CONFIG.ERROR_MESSAGES.NO_VALUE;
    } catch (error) {
      console.error('Error in info function:', error);
      return CONFIG.ERROR_MESSAGES.NO_VALUE;
    } finally {
      await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
    }
  }
  
  /**
   * Gets the service life (Nutzungsdauer) value for a dataset
   * @customfunction
   * @param {string} name - Dataset identifier
   * @returns {Promise<string|number>} Service life value or error message
   */
  async function life(name) {
    if (!name) {
      return CONFIG.ERROR_MESSAGES.NO_DATASET;
    }
  
    try {
      const response = await fetch(`${CONFIG.API_BASE_URL}/lifespan/${encodeURIComponent(name)}`);
      
      if (!response.ok) {
        return CONFIG.ERROR_MESSAGES.NO_DATASET;
      }
      
      const data = await response.json();
      
      if (data.warning) {
        _setWarningStatus('life');
      }
      
      if (data.status === 'limit_reached') {
        return CONFIG.ERROR_MESSAGES.LIMIT_REACHED;
      } else if (data.status === 'free_token') {
        return CONFIG.ERROR_MESSAGES.FREE_VERSION;
      }
      
      return data.value || CONFIG.ERROR_MESSAGES.NO_DATASET;
    } catch (error) {
      console.error('Error in life function:', error);
      return CONFIG.ERROR_MESSAGES.NO_DATASET;
    } finally {
      await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
    }
  }
  
  /**
   * Finds a matching dataset for a material/component
   * @customfunction
   * @param {string} name - Material/component name
   * @returns {Promise<string>} Matching dataset identifier or error message
   */
  async function data(name) {
    if (!name) {
      return CONFIG.ERROR_MESSAGES.NO_DATASET;
    }
  
    try {
      const response = await fetch(`${CONFIG.API_BASE_URL}/material-match/${encodeURIComponent(name)}`);
      
      if (!response.ok) {
        return CONFIG.ERROR_MESSAGES.NO_DATASET;
      }
      
      const data = await response.json();
      
      if (data.warning) {
        _setWarningStatus('data');
      }
      
      if (data.status === 'limit_reached') {
        return CONFIG.ERROR_MESSAGES.LIMIT_REACHED;
      } else if (data.status === 'free_token') {
        return CONFIG.ERROR_MESSAGES.FREE_VERSION;
      }
      
      return (data.match && data.match.material) 
        ? data.match.material 
        : CONFIG.ERROR_MESSAGES.NO_DATASET;
    } catch (error) {
      console.error('Error in data function:', error);
      return CONFIG.ERROR_MESSAGES.NO_DATASET;
    } finally {
      await OfficeRuntime.storage.setItem('customFunctionDone', Date.now());
    }
  }
  
  // Initialize Office Add-in
  Office.onReady((context) => {
    if (context.host === Office.HostType.Excel) {
      try {
        if (!CustomFunctions._namespace) {
          console.warn("Namespace missing! Setting manually...");
          CustomFunctions._namespace = "YOUR_NAMESPACE";
        }
  
        // Register all custom functions
        CustomFunctions.associate('get', get);
        CustomFunctions.associate('info', info);
        CustomFunctions.associate('life', life);
        CustomFunctions.associate('data', data);
        
        console.log('Excel functions registered');
        console.log("Registered functions:", Object.keys(CustomFunctions._association));
        console.log("Current namespace:", CustomFunctions._namespace);
      } catch (error) {
        console.error('Error registering functions:', error);
      }
    }
  });
