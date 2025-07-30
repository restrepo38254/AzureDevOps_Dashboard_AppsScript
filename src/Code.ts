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
    var stageFilter = params.stageFilter ? new RegExp(params.stageFilter, 'i') : null;
    var stageTypeFilter = params.stageTypeFilter ? params.stageTypeFilter.toLowerCase() : null;
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
    
    // 4. INICIALIZAR ESTRUCTURA DE DATOS
    var analysis = {
      totals: {
        pipelines: 0,      // Contador total de pipelines
        runs: 0,           // Total de ejecuciones
        success: 0,        // Ejecuciones exitosas
        failed: 0,         // Ejecuciones fallidas
        other: 0           // Otros estados (cancelados, etc.)
      },
      stages: {},          // Estadísticas por stage
      failedPipelines: [],  // Pipelines fallidos con detalles
      allPipelines: [],     // Todos los pipelines
      errorDetails: [],     // Detalles de errores
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
    pipelines.value.forEach(function(pipeline, index) {
      // Aplicar filtro por nombre de pipeline
      if (pipelineFilter && !pipelineFilter.test(pipeline.name)) {
        analysis.progress.processed++;
        return;
      }

      analysis.totals.pipelines++;
      analysis.allPipelines.push({
        id: pipeline.id,
        name: pipeline.name,
        url: `${CONFIG.azureDevOpsUrl}/${CONFIG.projectName}/_build?definitionId=${pipeline.id}`
      });

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

    var detailRequests = [];
    var detailInfo = [];
    var detailResponses = [];

    runResponses.forEach(function(runs, idx) {
      var pipeline = pipelinesOrdered[idx];
      if (!runs || !runs.value) {
        analysis.progress.processed++;
        return;
      }

      runs.value.slice(0, CONFIG.maxRuns).forEach(function(run) {
        if (!run.finishedDate || new Date(run.finishedDate) < cutoffDate) return;

        analysis.totals.runs++;
        var result = run.result ? run.result.toLowerCase() : 'other';

        if (result === 'succeeded') analysis.totals.success++;
        else if (result === 'failed') analysis.totals.failed++;
        else analysis.totals.other++;

        if (!analysis.pipelineStats[pipeline.name]) {
          analysis.pipelineStats[pipeline.name] = { success: 0, failed: 0, other: 0 };
        }
        if (result === 'succeeded') analysis.pipelineStats[pipeline.name].success++;
        else if (result === 'failed') analysis.pipelineStats[pipeline.name].failed++;
        else analysis.pipelineStats[pipeline.name].other++;

        if (run.createdDate && run.finishedDate) {
          var duration = (new Date(run.finishedDate) - new Date(run.createdDate)) / 60000;
          analysis.executionTimes.push({
            pipeline: pipeline.name,
            runId: run.id,
            duration: duration
          });
        }

        if (result === 'failed') {
          var cachedDetail = useCachedRuns && CACHE.runDetails[pipeline.id] && CACHE.runDetails[pipeline.id][run.id];
          if (cachedDetail) {
            detailResponses.push(cachedDetail);
            detailInfo.push({ pipeline: pipeline, run: run });
          } else {
            detailRequests.push(`pipelines/${pipeline.id}/runs/${run.id}?api-version=7.1-preview.1`);
            detailInfo.push({ pipeline: pipeline, run: run });
          }
        }
      });

      analysis.progress.processed++;
    });

    var fetchedDetails = callAzureApiBatch(detailRequests);
    fetchedDetails.forEach(function(res, idx) {
      var info = detailInfo[idx];
      CACHE.runDetails[info.pipeline.id] = CACHE.runDetails[info.pipeline.id] || {};
      CACHE.runDetails[info.pipeline.id][info.run.id] = res;
      detailResponses.push(res);
    });

    detailResponses.forEach(function(details, idx) {
      var info = detailInfo[idx];
      var pipeline = info.pipeline;
      var run = info.run;
      if (!details || !details.stages) return;

      var failedStages = [];
      var stageErrors = [];

      details.stages.forEach(function(stage) {
        if (stage.result && stage.result.toLowerCase() === 'failed') {
          const stageName = stage.name.toLowerCase();
          let stageType = 'other';

          if (stageName.includes('test')) stageType = 'tests';
          else if (stageName.includes('build')) stageType = 'build';
          else if (stageName.includes('deploy')) stageType = 'deploy';
          else if (stageName.includes('secure')) stageType = 'security';

          if (stageFilter && !stageFilter.test(stage.name)) return;
          if (stageTypeFilter && stageType !== stageTypeFilter) return;

          failedStages.push({
            name: stage.name,
            type: stageType
          });

          var errorInfo = {
            pipeline: pipeline.name,
            runId: run.id,
            stage: stage.name,
            errors: [],
            logUrl: stage.logs?.url ?
              `${CONFIG.azureDevOpsUrl}/${CONFIG.projectName}/_build/results?buildId=${run.id}&view=logs` : null
          };

          if (stage.error) errorInfo.errors.push(stage.error);
          if (stage.issues) errorInfo.errors = errorInfo.errors.concat(stage.issues);

          if (errorInfo.errors.length > 0) {
            stageErrors.push(errorInfo);
          }

          analysis.stages[stage.name] = (analysis.stages[stage.name] || 0) + 1;
        }
      });

      if (failedStages.length > 0) {
        analysis.failedPipelines.push({
          id: pipeline.id,
          name: pipeline.name,
          runId: run.id,
          date: run.finishedDate,
          url: `${CONFIG.azureDevOpsUrl}/${CONFIG.projectName}/_build/results?buildId=${run.id}`,
          stages: failedStages,
          stageErrors: stageErrors
        });

        analysis.errorDetails = analysis.errorDetails.concat(stageErrors);
      }
    });
    
    // 15. CALCULAR PORCENTAJES DE STAGES
    if (analysis.totals.failed > 0) {
      for (var stage in analysis.stages) {
        analysis.stages[stage] = {
          count: analysis.stages[stage],
          percentage: Math.round((analysis.stages[stage] / analysis.totals.failed) * 100)
        };
      }
    }
    // 16. CALCULAR PORCENTAJES POR PIPELINE
    for (var name in analysis.pipelineStats) {
      var stats = analysis.pipelineStats[name];
      var total = stats.success + stats.failed + stats.other;
      stats.successRate = total ? ((stats.success / total) * 100).toFixed(2) : '0.00';
      stats.failureRate = total ? ((stats.failed / total) * 100).toFixed(2) : '0.00';
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
  csvContent.push("Pipeline,Run ID,Status,Date,Duration (min),Failed Stages,Errors,Log URL");
  
  // 2. PREPARAR SOLICITUDES DE EJECUCIONES
  var runRequests = [];
  var pipelinesToProcess = [];
  data.data.allPipelines.forEach(function(pipeline) {
    pipelinesToProcess.push(pipeline);
    runRequests.push(`pipelines/${pipeline.id}/runs?$top=${CONFIG.maxRuns}&api-version=7.1-preview.1`);
  });

  var runResponses = callAzureApiBatch(runRequests);

  var detailRequests = [];
  var detailInfo = [];
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
        '',
        '',
        run.url || ''
      ];
      var rowIndex = rows.length;
      rows.push(row);

      if (run.result && run.result.toLowerCase() === 'failed') {
        detailRequests.push(`pipelines/${pipeline.id}/runs/${run.id}?api-version=7.1-preview.1`);
        detailInfo.push({ index: rowIndex });
      }
    });
  });

  var detailResponses = callAzureApiBatch(detailRequests);

  detailResponses.forEach(function(details, idx) {
    var info = detailInfo[idx];
    var row = rows[info.index];
    if (!details || !details.stages) return;

    var failedStages = details.stages
      .filter(function(s) { return s.result && s.result.toLowerCase() === 'failed'; })
      .map(function(s) { return s.name; });
    row[5] = `"${failedStages.join('; ').replace(/"/g, '""')}"`;

    var errors = details.stages
      .filter(function(s) { return s.error || s.issues; })
      .map(function(s) { return (s.error || '') + (s.issues ? '; ' + s.issues.join('; ') : ''); });
    row[6] = `"${errors.join(' | ').replace(/"/g, '""')}"`;
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
