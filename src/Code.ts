// CONFIGURACIÓN INICIAL COMPLETA
var CONFIG = {
  azureDevOpsUrl: "https://dev.azure.com/GrupoNutresa",  // Reemplaza {organization}
  projectName: "Novaventa",                      // Reemplaza con tu proyecto
  apiVersion: "7.1-preview.1",                             // Versión más reciente de API
  patToken: "",                      // Tu PAT de Azure DevOps
  maxPipelines: 500,                                       // Límite aumentado de pipelines
  maxRuns: 100,                                            // Máximo de ejecuciones por pipeline
  cacheDuration: 15                                        // Minutos de caché
};

// SISTEMA DE CACHÉ COMPLETO
var CACHE = {
  pipelines: null,         // Cache para listado de pipelines
  runs: {},                // Cache para ejecuciones por pipeline
  statistics: null,        // Cache para estadísticas
  lastUpdated: null,       // Fecha última actualización
  currentParams: null      // Parámetros actuales de filtrado
};

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
    var days = params.days || 30;
    var pipelineFilter = params.pipelineFilter ? new RegExp(params.pipelineFilter, 'i') : null;
    var stageFilter = params.stageFilter ? new RegExp(params.stageFilter, 'i') : null;
    var stageTypeFilter = params.stageTypeFilter ? params.stageTypeFilter.toLowerCase() : null;
    var cacheKey = JSON.stringify(params);
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
    
    // 5. PROCESAR CADA PIPELINE (BUCLE COMPLETO)
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
      
      // 6. OBTENER EJECUCIONES PARA ESTE PIPELINE
      console.log(`Procesando pipeline ${index + 1}/${pipelines.value.length}: ${pipeline.name}`);
      var runs = getPipelineRuns(pipeline.id);
      if (!runs || !runs.value) {
        analysis.progress.processed++;
        return;
      }
      
      // 7. PROCESAR CADA EJECUCIÓN (BUCLE COMPLETO)
      runs.value.slice(0, CONFIG.maxRuns).forEach(function(run) {
        // Filtrar por fecha
        if (!run.finishedDate || new Date(run.finishedDate) < cutoffDate) return;
        
        analysis.totals.runs++;
        var result = run.result ? run.result.toLowerCase() : 'other';
        
        // 8. CONTABILIZAR POR ESTADO
        if (result === 'succeeded') analysis.totals.success++;
        else if (result === 'failed') analysis.totals.failed++;
        else analysis.totals.other++;
        
        // Contadores por pipeline
        if (!analysis.pipelineStats[pipeline.name]) {
          analysis.pipelineStats[pipeline.name] = { success: 0, failed: 0, other: 0 };
        }
        if (result === 'succeeded') analysis.pipelineStats[pipeline.name].success++;
        else if (result === 'failed') analysis.pipelineStats[pipeline.name].failed++;
        else analysis.pipelineStats[pipeline.name].other++;
        
        // 9. REGISTRAR TIEMPO DE EJECUCIÓN
        if (run.createdDate && run.finishedDate) {
          var duration = (new Date(run.finishedDate) - new Date(run.createdDate)) / 60000; // en minutos
          analysis.executionTimes.push({
            pipeline: pipeline.name,
            runId: run.id,
            duration: duration
          });
        }
        
        // 10. ANALIZAR FALLOS (DETALLE COMPLETO)
        if (result === 'failed') {
          var details = getRunDetails(pipeline.id, run.id);
          if (!details || !details.stages) return;
          
          var failedStages = [];
          var stageErrors = [];
          
          // 11. PROCESAR CADA STAGE FALLIDO
          details.stages.forEach(function(stage) {
            if (stage.result && stage.result.toLowerCase() === 'failed') {
              // Clasificar el stage por tipo
              const stageName = stage.name.toLowerCase();
              let stageType = 'other';

              if (stageName.includes('test')) stageType = 'tests';
              else if (stageName.includes('build')) stageType = 'build';
              else if (stageName.includes('deploy')) stageType = 'deploy';
              else if (stageName.includes('secure')) stageType = 'security';

              // Aplicar filtros de stage
              if (stageFilter && !stageFilter.test(stage.name)) return;
              if (stageTypeFilter && stageType !== stageTypeFilter) return;

              failedStages.push({
                name: stage.name,
                type: stageType
              });
              
              // 12. CAPTURAR DETALLES DEL ERROR
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
              
              // 13. CONTABILIZAR FALLOS POR STAGE
              analysis.stages[stage.name] = (analysis.stages[stage.name] || 0) + 1;
            }
          });
          
          // 14. REGISTRAR PIPELINE FALLIDO
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
        }
      });
      
      analysis.progress.processed++;
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
  
  // 2. PROCESAR TODOS LOS PIPELINES
  data.data.allPipelines.forEach(function(pipeline) {
    var runs = getPipelineRuns(pipeline.id);
    if (!runs.value) return;
    
    // 3. PROCESAR CADA EJECUCIÓN
    runs.value.forEach(function(run) {
      var row = [
        `"${pipeline.name.replace(/"/g, '""')}"`,  // Escape de comillas
        run.id,
        run.result || 'unknown',
        run.finishedDate || '',
        run.createdDate && run.finishedDate ? 
          ((new Date(run.finishedDate) - new Date(run.createdDate)) / 60000).toFixed(2) : '',
        '',
        '',
        run.url || ''
      ];
      
      // 4. AGREGAR DETALLES DE FALLOS
      if (run.result && run.result.toLowerCase() === 'failed') {
        var details = getRunDetails(pipeline.id, run.id);
        if (details && details.stages) {
          // 5. STAGES FALLIDOS
          var failedStages = details.stages
            .filter(s => s.result && s.result.toLowerCase() === 'failed')
            .map(s => s.name);
          row[5] = `"${failedStages.join('; ').replace(/"/g, '""')}"`;
          
          // 6. DETALLES DE ERROR
          var errors = details.stages
            .filter(s => s.error || s.issues)
            .map(s => (s.error || '') + (s.issues ? '; ' + s.issues.join('; ') : ''));
          row[6] = `"${errors.join(' | ').replace(/"/g, '""')}"`;
        }
      }
      
      csvContent.push(row.join(','));
    });
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
      'Authorization': 'Basic ' + Utilities.base64Encode(':' + CONFIG.patToken),
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
    statistics: null,
    lastUpdated: null,
    currentParams: null
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
