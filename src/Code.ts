// CONFIGURACIÓN INICIAL COMPLETA
var CONFIG = {
  azureDevOpsUrl: "https://dev.azure.com/GrupoNutresa",  // Reemplaza {organization}
  projectName: "Novaventa",                      // Reemplaza con tu proyecto
  apiVersion: "7.1-preview.1",                             // Versión más reciente de API
  maxPipelines: 500,                                       // Límite aumentado de pipelines
  maxRuns: 100,                                            // Máximo de ejecuciones por pipeline
  cacheDuration: 15                                        // Minutos de caché
};

// SISTEMA DE CACHÉ COMPLETO
var CACHE = {
  pipelines: null,         // Cache para listado de pipelines
  runs: {},                // Cache para ejecuciones por pipeline
  runDetails: {},          // Cache para detalles de ejecuciones fallidas
  statistics: null,        // Cache para estadísticas
  projects: null,          // Cache para proyectos de Azure DevOps
  lastUpdated: null,       // Fecha última actualización
  currentParams: null,     // Parámetros actuales de filtrado
  baseParams: null         // Parámetros base (proyecto y rango de días)
};

/**
 * Obtiene el token PAT almacenado en las propiedades del script.
 * Se debe configurar previamente mediante setPatToken().
 * @return {string} PAT para autenticación
 */
function getPatToken() {
  return PropertiesService.getScriptProperties().getProperty('PAT_TOKEN') || '';
}

/**
 * Guarda el token PAT de forma segura en las propiedades del script.
 * @param {string} token - PAT de Azure DevOps
 */
function setPatToken(token) {
  PropertiesService.getScriptProperties().setProperty('PAT_TOKEN', token);
}

/**
 * FUNCIÓN PRINCIPAL PARA WEB APPS - REQUERIDA
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Azure DevOps Pipeline Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * OBTENER DATOS PARA EL DASHBOARD (FUNCIÓN PRINCIPAL)
 * @param {Object} params - Parámetros de filtrado
 * @return {Object} Datos completos para el dashboard
 */
function getDashboardData(params) {
  var startTime = new Date();
  console.log("Iniciando carga de datos - Hora de inicio: " + startTime.toISOString());

  try {
    // 1. PROCESAR PARÁMETROS DE FILTRADO
    if (params.projectName) {
      CONFIG.projectName = params.projectName;
    }
    var days = params.days || 30;
    var pipelineFilter = params.pipelineFilter ? new RegExp(params.pipelineFilter, 'i') : null;
    // Los filtros por stage ya no se utilizan
    var cacheKey = JSON.stringify(params);
    var baseKey = JSON.stringify({ projectName: CONFIG.projectName, days: days });
    var useCachedRuns = CACHE.baseParams === baseKey &&
      CACHE.lastUpdated &&
      (new Date() - CACHE.lastUpdated) < (CONFIG.cacheDuration * 60 * 1000);
    // 2. VERIFICAR CACHÉ (COMPROBACIÓN COMPLETA)
    if (CACHE.statistics && 
        CACHE.lastUpdated && 
        (new Date() - CACHE.lastUpdated) < (CONFIG.cacheDuration * 60 * 1000) && 
        CACHE.currentParams === cacheKey) {
      console.log("Devolviendo datos desde caché - Válido hasta: " + 
        new Date(CACHE.lastUpdated.getTime() + (CONFIG.cacheDuration * 60 * 1000)).toISOString());
      return {
        success: true,
        data: CACHE.statistics,
        lastUpdated: CACHE.lastUpdated.toISOString(),
        fromCache: true
      };
    }
    
    // 3. OBTENER PIPELINES (CON CACHÉ)
    var pipelines = CACHE.pipelines || getPipelines();
    CACHE.pipelines = pipelines;

    // Obtener repositorios activos para filtrar pipelines
    var activeRepoIds = new Set();
    var repos = getRepositories();
    if (repos && repos.value) {
      repos.value.forEach(function(r) {
        if (!r.isDisabled) activeRepoIds.add(r.id);
      });
    }
    
    // 4. INICIALIZAR ESTRUCTURA DE DATOS
    var analysis = {
      totals: {
        pipelines: 0,      // Contador total de pipelines
        runs: 0,           // Total de ejecuciones
        success: 0,        // Ejecuciones exitosas
        failed: 0,         // Ejecuciones fallidas
        other: 0           // Otros estados (cancelados, etc.)
      },
      failedPipelines: [],  // Pipelines fallidos con detalles
      allPipelines: [],     // Todos los pipelines
      executionTimes: [],   // Tiempos de ejecución
      progress: {          // Para seguimiento de progreso
        totalPipelines: pipelines.value.length,
        processed: 0
      }
    };
    analysis.pipelineStats = {};
    
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    // 5. PREPARAR SOLICITUDES DE EJECUCIONES
    var runRequests = [];
    var pipelinesToProcess = [];
    var runResponses = [];
    var pipelinesOrdered = [];
    var pipelineObjects = [];
    pipelines.value.forEach(function(pipeline, index) {
      // Excluir pipelines deshabilitados o de repositorios inactivos
      if (pipeline.queueStatus && pipeline.queueStatus.toLowerCase() !== 'enabled') {
        analysis.progress.processed++;
        return;
      }
      if (pipeline.repository && !activeRepoIds.has(pipeline.repository.id)) {
        analysis.progress.processed++;
        return;
      }
      // Aplicar filtro por nombre de pipeline
      if (pipelineFilter && !pipelineFilter.test(pipeline.name)) {
        analysis.progress.processed++;
        return;
      }

      analysis.totals.pipelines++;
      var pipelineObj = {
        id: pipeline.id,
        name: pipeline.name,
        url: `${CONFIG.azureDevOpsUrl}/${CONFIG.projectName}/_build?definitionId=${pipeline.id}`,
        lastRun: null,
        lastStatus: null,
        lastDuration: null
      };
      analysis.allPipelines.push(pipelineObj);
      pipelineObjects.push(pipelineObj);

      console.log(`Procesando pipeline ${index + 1}/${pipelines.value.length}: ${pipeline.name}`);

      var cachedRuns = useCachedRuns && CACHE.runs[pipeline.id];
      if (cachedRuns) {
        runResponses.push(cachedRuns);
        pipelinesOrdered.push(pipeline);
      } else {
        pipelinesToProcess.push(pipeline);
        runRequests.push(`pipelines/${pipeline.id}/runs?$top=${CONFIG.maxRuns}&api-version=7.1-preview.1`);
      }
    });
    var fetchedRuns = callAzureApiBatch(runRequests);
    fetchedRuns.forEach(function(res, idx) {
      var pipeline = pipelinesToProcess[idx];
      CACHE.runs[pipeline.id] = res;
      runResponses.push(res);
      pipelinesOrdered.push(pipeline);
    });


    runResponses.forEach(function(runs, idx) {
      var pipeline = pipelinesOrdered[idx];
      var pipelineObj = pipelineObjects[idx];
      if (!runs || !runs.value) {
        analysis.progress.processed++;
        return;
      }

      var firstValid = true;
      runs.value.slice(0, CONFIG.maxRuns).forEach(function(run) {
        var duration = null;
        if (!run.finishedDate || new Date(run.finishedDate) < cutoffDate) return;

        analysis.totals.runs++;
        var result = run.result ? run.result.toLowerCase() : 'other';

        if (result === 'succeeded') analysis.totals.success++;
        else if (result === 'failed') analysis.totals.failed++;
        else analysis.totals.other++;

        if (!analysis.pipelineStats[pipeline.name]) {
          analysis.pipelineStats[pipeline.name] = { success: 0, failed: 0, other: 0, durations: [] };
        }
        if (result === 'succeeded') analysis.pipelineStats[pipeline.name].success++;
        else if (result === 'failed') analysis.pipelineStats[pipeline.name].failed++;
        else analysis.pipelineStats[pipeline.name].other++;

        if (run.createdDate && run.finishedDate) {
          duration = (new Date(run.finishedDate) - new Date(run.createdDate)) / 60000;
          analysis.executionTimes.push({
            pipeline: pipeline.name,
            runId: run.id,
            duration: duration
          });
          analysis.pipelineStats[pipeline.name].durations.push(duration);
        }

        if (firstValid) {
          pipelineObj.lastRun = run.finishedDate;
          pipelineObj.lastStatus = run.result || 'unknown';
          pipelineObj.lastDuration = duration ? duration.toFixed(2) : null;
          firstValid = false;
        }

        if (result === 'failed') {
          analysis.failedPipelines.push({
            id: pipeline.id,
            name: pipeline.name,
            runId: run.id,
            date: run.finishedDate,
            url: `${CONFIG.azureDevOpsUrl}/${CONFIG.projectName}/_build/results?buildId=${run.id}`
          });
        }
      });

      analysis.progress.processed++;
    });

    // No se procesan detalles de stages
    // 16. CALCULAR PORCENTAJES POR PIPELINE
    for (var name in analysis.pipelineStats) {
      var stats = analysis.pipelineStats[name];
      var total = stats.success + stats.failed + stats.other;
      stats.successRate = total ? ((stats.success / total) * 100).toFixed(2) : '0.00';
      stats.failureRate = total ? ((stats.failed / total) * 100).toFixed(2) : '0.00';
      if (stats.durations && stats.durations.length) {
        var sum = stats.durations.reduce(function(a, b) { return a + b; }, 0);
        stats.avgDuration = (sum / stats.durations.length).toFixed(2);
      } else {
        stats.avgDuration = '0.00';
      }
    }
    // 16. CALCULAR ESTADÍSTICAS DE TIEMPO
    analysis.executionStats = calculateStats(analysis.executionTimes.map(x => x.duration));
    
    // 17. ACTUALIZAR CACHÉ
    CACHE.statistics = analysis;
    CACHE.lastUpdated = new Date();
    CACHE.currentParams = cacheKey;
    CACHE.baseParams = baseKey;
    
    var endTime = new Date();
    var processingTime = (endTime - startTime) / 1000;
    console.log(`Procesamiento completado en ${processingTime} segundos`);
    
    return {
      success: true,
      data: analysis,
      lastUpdated: endTime.toISOString(),
      processingTime: processingTime,
      fromCache: false
    };
    
  } catch (e) {
    console.error("Error en getDashboardData:", e);
    return {
      success: false,
      error: e.message,
      lastUpdated: new Date().toISOString()
    };
  }
}

/**
 * EXPORTAR DATOS A CSV (FUNCIÓN COMPLETA)
 * @param {Object} params - Parámetros de filtrado
 * @return {String} CSV con los datos
 */
function exportToCSV(params) {
  var data = getDashboardData(params);
  if (!data.success) throw new Error(data.error);
  
  var csvContent = [];
  
  // 1. ENCABEZADOS CSV
  csvContent.push("Pipeline,Run ID,Status,Date,Duration (min),URL");
  
  // 2. PREPARAR SOLICITUDES DE EJECUCIONES
  var runRequests = [];
  var pipelinesToProcess = [];
  data.data.allPipelines.forEach(function(pipeline) {
    pipelinesToProcess.push(pipeline);
    runRequests.push(`pipelines/${pipeline.id}/runs?$top=${CONFIG.maxRuns}&api-version=7.1-preview.1`);
  });

  var runResponses = callAzureApiBatch(runRequests);

  var rows = [];

  runResponses.forEach(function(runs, idx) {
    var pipeline = pipelinesToProcess[idx];
    if (!runs || !runs.value) return;

    runs.value.forEach(function(run) {
      var row = [
        `"${pipeline.name.replace(/"/g, '""')}"`,
        run.id,
        run.result || 'unknown',
        run.finishedDate || '',
        run.createdDate && run.finishedDate ?
          ((new Date(run.finishedDate) - new Date(run.createdDate)) / 60000).toFixed(2) : '',
        run.url || ''
      ];
      rows.push(row);
    });
  });


  rows.forEach(function(r) {
    csvContent.push(r.join(','));
  });
  
  return csvContent.join('\n');
}

// FUNCIONES AUXILIARES COMPLETAS:

/**
 * OBTENER LISTADO DE PIPELINES
 * @return {Object} Respuesta de API
 */
function getPipelines() {
  return callAzureApi("pipelines?$top=" + CONFIG.maxPipelines);
}

/**
 * OBTENER EJECUCIONES DE UN PIPELINE
 * @param {Number} pipelineId - ID del pipeline
 * @return {Object} Respuesta de API
 */
function getPipelineRuns(pipelineId) {
  return callAzureApi(`pipelines/${pipelineId}/runs?$top=${CONFIG.maxRuns}&api-version=7.1-preview.1`);
}

/**
 * OBTENER DETALLES DE EJECUCIÓN
 * @param {Number} pipelineId - ID del pipeline
 * @param {Number} runId - ID de la ejecución
 * @return {Object} Respuesta de API
 */
function getRunDetails(pipelineId, runId) {
  return callAzureApi(`pipelines/${pipelineId}/runs/${runId}?api-version=7.1-preview.1`);
}

/**
 * LLAMADA GENÉRICA A API AZURE DEVOPS
 * @param {String} endpoint - Endpoint de API
 * @return {Object} Respuesta parseada
 */
function callAzureApi(endpoint) {
  var url = endpoint.startsWith('http') ? endpoint :
    `${CONFIG.azureDevOpsUrl}/${CONFIG.projectName}/_apis/${endpoint}`;
  
  var options = {
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(':' + getPatToken()),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    console.error("API Error:", e, "Endpoint:", url);
    return null;
  }
}

/**
 * LLAMADAS EN BATCH A API AZURE DEVOPS
 * @param {Array<string>} endpoints - Lista de endpoints
 * @return {Array<Object>} Respuestas parseadas
 */
function callAzureApiBatch(endpoints) {
  if (!endpoints || endpoints.length === 0) return [];

  var options = {
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(':' + getPatToken()),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  var requests = endpoints.map(function(ep) {
    var url = ep.startsWith('http') ? ep :
      `${CONFIG.azureDevOpsUrl}/${CONFIG.projectName}/_apis/${ep}`;
    return { url: url, headers: options.headers, muteHttpExceptions: true };
  });

  try {
    var responses = UrlFetchApp.fetchAll(requests);
    return responses.map(function(res) {
      try {
        return JSON.parse(res.getContentText());
      } catch (e) {
        console.error('Batch parse error', e);
        return null;
      }
    });
  } catch (e) {
    console.error('Batch API Error', e);
    return endpoints.map(function() { return null; });
  }
}

/**
 * CALCULAR ESTADÍSTICAS DE TIEMPOS
 * @param {Array} values - Array de valores
 * @return {Object} Estadísticas
 */
function calculateStats(values) {
  if (!values || values.length === 0) return null;
  
  values.sort((a, b) => a - b);
  
  return {
    average: (values.reduce((a, b) => a + b, 0) / values.length).toFixed(2),
    median: values[Math.floor(values.length / 2)].toFixed(2),
    min: Math.min(...values).toFixed(2),
    max: Math.max(...values).toFixed(2),
    percentile95: values[Math.floor(values.length * 0.95)].toFixed(2)
  };
}

/**
 * LIMPIAR CACHÉ MANUALMENTE
 * @return {Object} Resultado de operación
 */
function clearCache() {
  CACHE = {
    pipelines: null,
    runs: {},
    runDetails: {},
    statistics: null,
    projects: null,
    lastUpdated: null,
    currentParams: null,
    baseParams: null
  };
  console.log("Cache cleared manually");
  return {
    success: true,
    message: "Cache cleared successfully",
    timestamp: new Date().toISOString()
  };
}

/**
 * OBTENER LISTADO DE REPOSITORIOS
 * @return {Object} Respuesta de API
 */
function getRepositories() {
  return callAzureApi("git/repositories?api-version=" + CONFIG.apiVersion);
}

/**
 * Obtener lista de proyectos disponibles en la organización
 * El resultado se cachea para evitar llamadas repetitivas
 * @return {Object} Respuesta de API
 */
function getProjects() {
  if (CACHE.projects) {
    return CACHE.projects;
  }
  var result = callAzureApi(`${CONFIG.azureDevOpsUrl}/_apis/projects?api-version=${CONFIG.apiVersion}`);
  CACHE.projects = result;
  return result;
}

/**
 * Método expuesto al frontend para obtener la lista de proyectos
 */
function listProjects() {
  return getProjects();
}
