(function () {
  'use strict';

  var EVENT_TYPE = 'reports:run';
  var pendingJob = null;
  var lastJobId = null;
  var runTimer = null;

  function log(message) {
    try {
      if (window.console && console.log) {
        console.log('[reports-engine]', message);
      }
    } catch (e) {
      // ignore
    }
  }

  function scheduleRun(delay) {
    if (runTimer) {
      clearTimeout(runTimer);
    }
    runTimer = setTimeout(runJob, delay || 0);
  }

  function normalizeAddress(addr) {
    return String(addr || '').trim();
  }

  function normalizeSheetName(name) {
    return String(name || '').trim();
  }

  function runJob() {
    if (!pendingJob) return;

    var job = pendingJob;
    pendingJob = null;

    try {
      Asc.scope.reportsJob = job;
      Asc.plugin.callCommand(function () {
        var api = (window.Asc && window.Asc.editor) ? window.Asc.editor : window.Asc;
        var job = Asc.scope.reportsJob || {};
        var actions = job.actions || [];
        var debug = !!job.debug;
        var skip = window.__reportsPluginSkip;
        if (skip && job.id && skip.jobId === job.id) {
          return;
        }

        function normalizeAddressLocal(addr) {
          return String(addr || '').trim();
        }

        function normalizeSheetNameLocal(name) {
          return String(name || '').trim();
        }

        function buildRangeAddressLocal(sheet, addr) {
          var a = normalizeAddressLocal(addr);
          if (!a) return '';
          var s = normalizeSheetNameLocal(sheet);
          if (!s) return a;
          return s + '!' + a;
        }

        function resolveTargetAddressLocal(action) {
          if (!action) return '';
          if (action.targetMode === 'key') {
            return normalizeAddressLocal(action.target);
          }
          return normalizeAddressLocal(action.target);
        }

        function setRange(sheet, addr) {
          var full = buildRangeAddressLocal(sheet, addr);
          if (full) {
            api.asc_setWorksheetRange(full);
          }
        }

        function insertText(value) {
          var text = String(value || '');
          if (api.asc_insertInCell) {
            api.asc_insertInCell(text);
            if (api.asc_closeCellEditor) {
              api.asc_closeCellEditor();
            }
            return;
          }
          if (api.asc_enterText) {
            var v = text;
            try {
              v = (v && v.codePointsArray) ? v.codePointsArray() : v;
            } catch (e) {}
            api.asc_enterText(v);
            if (api.asc_closeCellEditor) {
              api.asc_closeCellEditor();
            }
          }
        }

        if (debug) {
          try {
            setRange('', 'Z1');
            insertText('REPORTS DEBUG ' + (job.id || ''));
          } catch (e) {
            // ignore
          }
        }

        function runAction(action) {
          if (!action || !action.type) return;

          if (action.type === 'setText') {
            var targetAddr = resolveTargetAddressLocal(action);
            if (!targetAddr) return;
            setRange(action.sheet, targetAddr);
            insertText(action.value || '');
            if (action.merge) {
              try {
                api.asc_mergeCells();
              } catch (e) {
                // ignore
              }
            }
          } else if (action.type === 'groupCols') {
            if (!action.range) return;
            setRange(action.sheet, action.range);
            api.asc_group(false);
            if (typeof action.expanded === 'boolean') {
              try {
                api.asc_changeGroupDetails(!!action.expanded);
              } catch (e) {
                // ignore
              }
            }
          } else if (action.type === 'deleteRow') {
            if (!action.row) return;
            var rowAddr = action.row + ':' + action.row;
            setRange(action.sheet, rowAddr);
            api.asc_deleteCells(Asc.c_oAscDeleteOptions.DeleteRows);
          }
        }

        for (var i = 0; i < actions.length; i += 1) {
          runAction(actions[i]);
        }
      });
    } catch (e) {
      log('run failed: ' + e);
    }
  }

  window.Asc.plugin.onExternalPluginMessage = function (data) {
    if (!data || data.type !== EVENT_TYPE) return;
    if (!data.job) return;
    if (data.job.id && data.job.id === lastJobId) return;

    lastJobId = data.job.id || lastJobId;
    pendingJob = data.job;
    scheduleRun(200);
  };

  window.Asc.plugin.init = function () {
    // system plugin, no UI
  };
})();
