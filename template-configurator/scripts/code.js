(function(){
  "use strict";

  var STORAGE_KEY = "uv_template_config";

  function uid() {
    return "id_" + Math.random().toString(36).slice(2, 10);
  }

  function loadData() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) {}
    return { templates: [] };
  }

  function saveData() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.data, null, 2));
  }

  function byId(list, id) {
    return list.find(function(x){ return x.id === id; });
  }

  var state = {
    data: loadData(),
    activeTemplateId: null
  };

  function ensureTemplate() {
    if (!state.data.templates.length) {
      var t = { id: uid(), name: "Новый шаблон", fields: [], rules: [] };
      state.data.templates.push(t);
      state.activeTemplateId = t.id;
      saveData();
    }
    if (!state.activeTemplateId) {
      state.activeTemplateId = state.data.templates[0].id;
    }
  }

  function renderTemplates() {
    var list = document.getElementById("templateList");
    list.innerHTML = "";
    state.data.templates.forEach(function(t){
      var item = document.createElement("div");
      item.className = "card" + (t.id === state.activeTemplateId ? " active" : "");
      item.innerHTML = "<div><div class='card-title'>" + escapeHtml(t.name) + "</div>" +
        "<div class='card-sub'>" + t.fields.length + " полей, " + t.rules.length + " правил</div></div>" +
        "<div class='card-actions'><button class='btn link' data-id='"+t.id+"'>Открыть</button></div>";
      item.querySelector("button").onclick = function(){
        state.activeTemplateId = t.id;
        renderAll();
      };
      list.appendChild(item);
    });
  }

  function renderFields() {
    var tpl = getActiveTemplate();
    var list = document.getElementById("fieldsList");
    list.innerHTML = "";
    if (!tpl) return;
    tpl.fields.forEach(function(f){
      var item = document.createElement("div");
      item.className = "card";
      item.innerHTML = "<div><div class='card-title'>" + escapeHtml(f.name) + "</div>" +
        "<div class='card-sub'>" + escapeHtml(f.type) + " • " + escapeHtml(f.targetType) + " • " + escapeHtml(f.address || "") + "</div></div>" +
        "<div class='card-actions'>" +
        "<button class='btn link' data-act='edit'>Изменить</button>" +
        "<button class='btn link danger' data-act='del'>Удалить</button>" +
        "</div>";
      item.querySelector("[data-act='edit']").onclick = function(){ editField(f.id); };
      item.querySelector("[data-act='del']").onclick = function(){ deleteField(f.id); };
      list.appendChild(item);
    });
  }

  function renderRules() {
    var tpl = getActiveTemplate();
    var list = document.getElementById("rulesList");
    list.innerHTML = "";
    if (!tpl) return;
    tpl.rules.forEach(function(r){
      var item = document.createElement("div");
      item.className = "card";
      item.innerHTML = "<div><div class='card-title'>" + escapeHtml(r.name) + "</div>" +
        "<div class='card-sub'>" + escapeHtml(r.conditionType) + " → " + escapeHtml(r.actionType) + "</div></div>" +
        "<div class='card-actions'>" +
        "<button class='btn link' data-act='edit'>Изменить</button>" +
        "<button class='btn link danger' data-act='del'>Удалить</button>" +
        "</div>";
      item.querySelector("[data-act='edit']").onclick = function(){ editRule(r.id); };
      item.querySelector("[data-act='del']").onclick = function(){ deleteRule(r.id); };
      list.appendChild(item);
    });
  }

  function renderAll() {
    renderTemplates();
    var tpl = getActiveTemplate();
    document.getElementById("templateName").textContent = tpl ? tpl.name : "—";
    renderFields();
    renderRules();
  }

  function getActiveTemplate() {
    return byId(state.data.templates, state.activeTemplateId);
  }

  function escapeHtml(s){
    return String(s || "").replace(/[&<>\"']/g, function(m){
      return ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"})[m];
    });
  }

  function openModal(title, html, onOk) {
    var modal = document.getElementById("modal");
    document.getElementById("modalTitle").textContent = title;
    var content = document.getElementById("modalContent");
    content.innerHTML = html;
    modal.classList.remove("hidden");

    document.getElementById("modalCancel").onclick = function(){
      modal.classList.add("hidden");
    };
    document.getElementById("modalOk").onclick = function(){
      if (onOk) onOk(content);
      modal.classList.add("hidden");
    };
  }

  function addTemplate() {
    openModal("Новый шаблон", "<div class='row'><div><label>Название</label><input class='input' id='tplName'></div></div>", function(c){
      var name = c.querySelector("#tplName").value || "Новый шаблон";
      var t = { id: uid(), name: name, fields: [], rules: [] };
      state.data.templates.push(t);
      state.activeTemplateId = t.id;
      saveData();
      renderAll();
    });
  }

  function renameTemplate() {
    var tpl = getActiveTemplate();
    if (!tpl) return;
    openModal("Переименовать шаблон", "<div class='row'><div><label>Название</label><input class='input' id='tplName' value='"+escapeHtml(tpl.name)+"'></div></div>", function(c){
      tpl.name = c.querySelector("#tplName").value || tpl.name;
      saveData();
      renderAll();
    });
  }

  function deleteTemplate() {
    var tpl = getActiveTemplate();
    if (!tpl) return;
    state.data.templates = state.data.templates.filter(function(t){ return t.id !== tpl.id; });
    state.activeTemplateId = state.data.templates.length ? state.data.templates[0].id : null;
    saveData();
    ensureTemplate();
    renderAll();
  }

  function fieldForm(f) {
    f = f || { name: "", type: "text", targetType: "cell", sheet: "", address: "", value: "" };
    return "" +
      "<div class='row'>" +
      "<div><label>Название поля</label><input class='input' id='fName' value='"+escapeHtml(f.name)+"'></div>" +
      "<div><label>Тип</label><select id='fType' class='input'>" +
        optionList(["text","date","number","list"], f.type) +
      "</select></div>" +
      "</div>" +
      "<div class='row'>" +
      "<div><label>Цель (тип)</label><select id='fTargetType' class='input'>" +
        optionList(["cell","range","namedRange"], f.targetType) +
      "</select></div>" +
      "<div><label>Лист (опционально)</label><input class='input' id='fSheet' value='"+escapeHtml(f.sheet)+"'></div>" +
      "</div>" +
      "<div class='row'>" +
      "<div><label>Адрес (A1, A1:B2, имя диапазона)</label><input class='input' id='fAddress' value='"+escapeHtml(f.address)+"'></div>" +
      "<div><label>Значение</label><input class='input' id='fValue' value='"+escapeHtml(f.value)+"'></div>" +
      "</div>";
  }

  function addField() {
    var tpl = getActiveTemplate();
    if (!tpl) return;
    openModal("Новое поле", fieldForm(), function(c){
      var f = {
        id: uid(),
        name: c.querySelector("#fName").value || "Поле",
        type: c.querySelector("#fType").value,
        targetType: c.querySelector("#fTargetType").value,
        sheet: c.querySelector("#fSheet").value,
        address: c.querySelector("#fAddress").value,
        value: c.querySelector("#fValue").value
      };
      tpl.fields.push(f);
      saveData();
      renderFields();
    });
  }

  function editField(id) {
    var tpl = getActiveTemplate();
    var f = byId(tpl.fields, id);
    if (!f) return;
    openModal("Изменить поле", fieldForm(f), function(c){
      f.name = c.querySelector("#fName").value || f.name;
      f.type = c.querySelector("#fType").value;
      f.targetType = c.querySelector("#fTargetType").value;
      f.sheet = c.querySelector("#fSheet").value;
      f.address = c.querySelector("#fAddress").value;
      f.value = c.querySelector("#fValue").value;
      saveData();
      renderFields();
    });
  }

  function deleteField(id) {
    var tpl = getActiveTemplate();
    tpl.fields = tpl.fields.filter(function(x){ return x.id !== id; });
    saveData();
    renderFields();
  }

  function ruleForm(r) {
    r = r || { name: "", targetType: "range", sheet: "", address: "", conditionType: "always", conditionValue: "", actionType: "hideRows", actionValue: "" };
    return "" +
      "<div class='row'>" +
      "<div><label>Название правила</label><input class='input' id='rName' value='"+escapeHtml(r.name)+"'></div>" +
      "<div><label>Цель (тип)</label><select id='rTargetType' class='input'>" +
        optionList(["cell","range","rows","columns"], r.targetType) +
      "</select></div>" +
      "</div>" +
      "<div class='row'>" +
      "<div><label>Лист (опционально)</label><input class='input' id='rSheet' value='"+escapeHtml(r.sheet)+"'></div>" +
      "<div><label>Адрес (A1, A1:B2, 5:10, B:D)</label><input class='input' id='rAddress' value='"+escapeHtml(r.address)+"'></div>" +
      "</div>" +
      "<div class='row'>" +
      "<div><label>Условие</label><select id='rConditionType' class='input'>" +
        optionList(["always","cellEmpty","fieldNotEmpty","fieldEquals","fieldCountGE"], r.conditionType) +
      "</select></div>" +
      "<div><label>Значение условия</label><input class='input' id='rConditionValue' value='"+escapeHtml(r.conditionValue)+"'></div>" +
      "</div>" +
      "<div class='row'>" +
      "<div><label>Действие</label><select id='rActionType' class='input'>" +
        optionList(["hideRows","showRows","hideColumns","showColumns","insertRows","deleteRows","setValue","setFormula","format"], r.actionType) +
      "</select></div>" +
      "<div><label>Значение действия</label><input class='input' id='rActionValue' value='"+escapeHtml(r.actionValue)+"' placeholder='пример: 3 или {" + "" + "\"bold\":true," + "\"align\":\"center\"}" + "'></div>" +
      "</div>";
  }

  function addRule() {
    var tpl = getActiveTemplate();
    if (!tpl) return;
    openModal("Новое правило", ruleForm(), function(c){
      var r = {
        id: uid(),
        name: c.querySelector("#rName").value || "Правило",
        targetType: c.querySelector("#rTargetType").value,
        sheet: c.querySelector("#rSheet").value,
        address: c.querySelector("#rAddress").value,
        conditionType: c.querySelector("#rConditionType").value,
        conditionValue: c.querySelector("#rConditionValue").value,
        actionType: c.querySelector("#rActionType").value,
        actionValue: c.querySelector("#rActionValue").value
      };
      tpl.rules.push(r);
      saveData();
      renderRules();
    });
  }

  function editRule(id) {
    var tpl = getActiveTemplate();
    var r = byId(tpl.rules, id);
    if (!r) return;
    openModal("Изменить правило", ruleForm(r), function(c){
      r.name = c.querySelector("#rName").value || r.name;
      r.targetType = c.querySelector("#rTargetType").value;
      r.sheet = c.querySelector("#rSheet").value;
      r.address = c.querySelector("#rAddress").value;
      r.conditionType = c.querySelector("#rConditionType").value;
      r.conditionValue = c.querySelector("#rConditionValue").value;
      r.actionType = c.querySelector("#rActionType").value;
      r.actionValue = c.querySelector("#rActionValue").value;
      saveData();
      renderRules();
    });
  }

  function deleteRule(id) {
    var tpl = getActiveTemplate();
    tpl.rules = tpl.rules.filter(function(x){ return x.id !== id; });
    saveData();
    renderRules();
  }

  function optionList(list, selected) {
    return list.map(function(v){
      var sel = v === selected ? " selected" : "";
      return "<option value='"+v+"'"+sel+">"+v+"</option>";
    }).join("");
  }

  function applyTemplate() {
    var tpl = getActiveTemplate();
    if (!tpl) return;
    var payload = JSON.parse(JSON.stringify(tpl));
    Asc.scope.templatePayload = payload;
    Asc.plugin.callCommand(function(){
      var api = (window.Asc && Asc.editor) ? Asc.editor : Asc;
      var tpl = Asc.scope.templatePayload;

      function setRange(target) {
        var addr = target.address || "";
        if (target.sheet) addr = target.sheet + "!" + addr;
        api.asc_setWorksheetRange(addr);
      }

      function applyField(f) {
        if (!f.address) return;
        setRange(f);
        api.asc_insertInCell(String(f.value || ""));
      }

      function getCellValueAt(target) {
        setRange(target);
        var info = api.asc_getCellInfo();
        return info && info.text ? info.text : "";
      }

      function conditionOk(rule) {
        if (rule.conditionType === "always") return true;
        if (rule.conditionType === "cellEmpty") {
          var v = getCellValueAt(rule);
          return v === "" || v === null || v === undefined;
        }
        if (rule.conditionType === "fieldNotEmpty") {
          var f = tpl.fields.find(function(x){ return x.name === rule.conditionValue; });
          return f && String(f.value || "") !== "";
        }
        if (rule.conditionType === "fieldEquals") {
          var f2 = tpl.fields.find(function(x){ return x.name === rule.conditionValue.split("=")[0]; });
          var exp = rule.conditionValue.split("=")[1] || "";
          return f2 && String(f2.value || "") === exp;
        }
        if (rule.conditionType === "fieldCountGE") {
          var f3 = tpl.fields.find(function(x){ return x.name === rule.conditionValue.split(":")[0]; });
          var min = parseInt((rule.conditionValue.split(":")[1] || "0"), 10);
          if (!f3) return false;
          var count = String(f3.value || "").split(/\n+/).filter(Boolean).length;
          return count >= min;
        }
        return false;
      }

      function applyRule(r) {
        if (!r.address) return;
        if (!conditionOk(r)) return;
        setRange(r);

        if (r.actionType === "hideRows") api.asc_hideRows();
        else if (r.actionType === "showRows") api.asc_showRows();
        else if (r.actionType === "hideColumns") api.asc_hideColumns();
        else if (r.actionType === "showColumns") api.asc_showColumns();
        else if (r.actionType === "insertRows") api.asc_insertCells(Asc.c_oAscInsertOptions.InsertRows);
        else if (r.actionType === "deleteRows") api.asc_deleteCells(Asc.c_oAscDeleteOptions.DeleteRows);
        else if (r.actionType === "setValue") api.asc_insertInCell(String(r.actionValue || ""));
        else if (r.actionType === "setFormula") api.asc_insertInCell(String(r.actionValue || ""));
        else if (r.actionType === "format") {
          try {
            var fmt = JSON.parse(r.actionValue || "{}");
            if (fmt.bold !== undefined) api.asc_setCellBold(!!fmt.bold);
            if (fmt.align) api.asc_setCellAlign(fmt.align);
            if (fmt.fill) api.asc_setCellFill(fmt.fill);
          } catch(e) {}
        }
      }

      tpl.fields.forEach(applyField);
      tpl.rules.forEach(applyRule);
    });
  }

  function exportData() {
    openModal("Экспорт", "<textarea id='exp' class='input' style='height:220px'></textarea>", function(){});
    var txt = document.getElementById("exp");
    txt.value = JSON.stringify(state.data, null, 2);
  }

  function importData() {
    openModal("Импорт", "<textarea id='imp' class='input' style='height:220px'></textarea>", function(c){
      try {
        var data = JSON.parse(c.querySelector("#imp").value);
        if (data && data.templates) {
          state.data = data;
          state.activeTemplateId = data.templates.length ? data.templates[0].id : null;
          saveData();
          ensureTemplate();
          renderAll();
        }
      } catch(e) {}
    });
  }

  function bindTabs() {
    var tabs = document.querySelectorAll(".tab");
    tabs.forEach(function(t){
      t.onclick = function(){
        tabs.forEach(function(x){ x.classList.remove("active"); });
        t.classList.add("active");
        var name = t.getAttribute("data-tab");
        document.querySelectorAll(".tab-panel").forEach(function(p){
          p.classList.remove("active");
        });
        document.getElementById(name === "form" ? "tabForm" : "tabRules").classList.add("active");
      };
    });
  }

  window.Asc.plugin.init = function() {
    ensureTemplate();
    bindTabs();
    renderAll();
  };

  window.Asc.plugin.button = function(id) {
    if (id === 0) window.Asc.plugin.executeCommand("close", "");
  };

  document.addEventListener("DOMContentLoaded", function(){
    document.getElementById("btnAddTemplate").onclick = addTemplate;
    document.getElementById("btnRenameTemplate").onclick = renameTemplate;
    document.getElementById("btnDeleteTemplate").onclick = deleteTemplate;
    document.getElementById("btnAddField").onclick = addField;
    document.getElementById("btnAddRule").onclick = addRule;
    document.getElementById("btnSave").onclick = function(){ saveData(); };
    document.getElementById("btnApply").onclick = applyTemplate;
    document.getElementById("btnExport").onclick = exportData;
    document.getElementById("btnImport").onclick = importData;
  });
})();
