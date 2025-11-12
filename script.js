// script.js - Fresadora PRO 2.1 Web Edition
// Traducción directa de la lógica Python -> JS

// ----- Parámetros (mismos que en Python) -----
const PARAMS_OPERACION = {
  "Taladrado": {herramienta:"Broca HSS", fz:null, Vc:null, z:2, ae_rule:"full"},
  "Avellanado": {herramienta:"Avellanador HSS", fz:null, Vc:null, z:3, ae_rule:"full"},
  "Abocardado": {herramienta:"Escariador HSS", fz:null, Vc:null, z:3, ae_rule:"full"},
  "Ranurado": {herramienta:"Fresa Ranuradora", fz:null, Vc:null, z:4, ae_rule:"partial_60"}
};

const MATERIALES = {
  "Acero medio carbono (AISI 1045)": {Vc:100.0, fz:0.10, Kc:1800.0, eta:0.80},
  "Aluminio (serie 6xxx)": {Vc:250.0, fz:0.08, Kc:600.0, eta:0.85},
  "Acero inoxidable (AISI 304)": {Vc:60.0, fz:0.06, Kc:2400.0, eta:0.78},
  "Fundición gris": {Vc:90.0, fz:0.09, Kc:1400.0, eta:0.75},
  "Latón": {Vc:180.0, fz:0.12, Kc:900.0, eta:0.85}
};

// ----- Helpers: mismas fórmulas -----
function calcularRPM(Vc, D){
  Vc = Number(Vc); D = Number(D);
  if(!D || !Vc) return 0;
  return (1000.0 * Vc) / (Math.PI * D);
}
function calcularVf(fz, z, rpm){
  fz = Number(fz); z = Number(z); rpm = Number(rpm);
  if(!fz || !z || !rpm) return 0;
  return fz * z * rpm;
}
function calcularMRR(ap, ae, vf){
  ap = Number(ap); ae = Number(ae); vf = Number(vf);
  if(!ap || !ae || !vf) return 0;
  return ap * ae * vf;
}
function calcularP_kW(mrr, Kc, eta){
  mrr = Number(mrr); Kc = Number(Kc); eta = Number(eta);
  if(!mrr || !Kc || !eta) return 0;
  return (mrr * Kc) / (60.0 * 1e3 * eta);
}
function calcularTorque(P_kW, rpm){
  if(!P_kW || !rpm) return 0;
  return (9550.0 * P_kW) / rpm;
}
function calcularFuerza(P_kW, Vc){
  if(!P_kW || !Vc) return 0;
  return (P_kW * 60000.0) / Vc;
}
function calcularTiempo(L, vf){
  if(!L || !vf) return 0;
  return L / vf;
}

// ----- DOM refs -----
const el = id => document.getElementById(id);
const fase = el("fase"), subfase = el("subfase"), operacion = el("operacion"),
      material = el("material"), herramienta = el("herramienta"),
      D = el("D"), z = el("z"), fz = el("fz"), Vc = el("Vc"),
      ap = el("ap"), ae = el("ae"), rpm = el("rpm"),
      vf = el("vf"), mrr = el("mrr"), pkw = el("pkw"), torque = el("torque"),
      fc = el("fc"), L = el("L"), tme = el("tme"), pasadas = el("pasadas"),
      utiles = el("utiles"), refrigeracion = el("refrigeracion"), notas = el("notas");

const tblBody = document.querySelector("#tbl tbody");
const btnAgregar = el("btnAgregar"), btnEliminar = el("btnEliminar"),
      btnLimpiar = el("btnLimpiar"), btnGuardar = el("btnGuardar"),
      btnCargar = el("btnCargar"), inputCargar = el("inputCargar");

// ----- Inicialización de selects -----
function initMateriales(){
  const keys = Object.keys(MATERIALES);
  material.innerHTML = "";
  keys.forEach(k=>{
    const opt = document.createElement("option");
    opt.value = k; opt.textContent = k;
    material.appendChild(opt);
  });
}
function initOperacion(){
  // Set herramienta y z por defecto
  herramienta.value = PARAMS_OPERACION[operacion.value].herramienta;
  if(!z.value) z.value = PARAMS_OPERACION[operacion.value].z || "";
  applyAeRule();
}

// ----- Reglas de a_e según operación -----
function applyAeRule(){
  const op = operacion.value;
  const rule = PARAMS_OPERACION[op].ae_rule || "full";
  const Dval = parseFloat(D.value) || 0;
  if(rule === "full"){
    if(Dval>0 && !ae.value) ae.value = Dval.toFixed(3);
  } else if(rule === "partial_60"){
    if(Dval>0 && !ae.value) ae.value = (0.6*Dval).toFixed(3);
  }
}

// ----- Recalcular preview (en vivo) -----
function recalcularPreview(){
  applyAeRule();
  const Dv = Number(D.value) || 0;
  const Vcv = Number(Vc.value) || 0;
  const zv = Number(z.value) || 0;
  const fzv = Number(fz.value) || 0;
  const apv = Number(ap.value) || 0;
  const aev = Number(ae.value) || 0;
  const Lv = Number(L.value) || 0;

  const rpmv = (Dv && Vcv)? calcularRPM(Vcv, Dv) : 0;
  const vfv = calcularVf(fzv, zv, rpmv);
  const mrrv = calcularMRR(apv, aev, vfv);
  const mat = material.value;
  const Kc = (MATERIALES[mat] && MATERIALES[mat].Kc) ? MATERIALES[mat].Kc : 1800.0;
  const eta = (MATERIALES[mat] && MATERIALES[mat].eta) ? MATERIALES[mat].eta : 0.8;
  const pkwv = calcularP_kW(mrrv, Kc, eta);
  const torqv = calcularTorque(pkwv, rpmv);
  const fcv = calcularFuerza(pkwv, Vcv);
  const timev = calcularTiempo(Lv, vfv);

  rpm.value = rpmv? rpmv.toFixed(2) : "0.00";
  vf.value = vfv? vfv.toFixed(2) : "0.00";
  mrr.value = mrrv? mrrv.toFixed(2) : "0.00";
  pkw.value = pkwv? pkwv.toFixed(4) : "0.0000";
  torque.value = torqv? torqv.toFixed(2) : "0.00";
  fc.value = fcv? fcv.toFixed(2) : "0.00";
  tme.value = timev? timev.toFixed(4) : "0.0000";
}

// ----- Agregar fila -----
function agregarFila(){
  // Validaciones mínimas
  const op = operacion.value;
  const Dv = Number(D.value) || 0;
  const Vcv = Number(Vc.value) || 0;
  if(!op){ alert("Selecciona una operación."); return; }
  if(!Dv){ alert("Ingresa Ø herramienta válido (>0)."); return; }
  if(!Vcv){ alert("Ingresa Vc (m/min) válido (>0)."); return; }

  // recalcular para asegurar valores
  recalcularPreview();

  // construir fila
  const row = document.createElement("tr");
  const cols = [
    fase.value || "", subfase.value || "", operacion.value || "", material.value || "", herramienta.value || "",
    (Number(D.value)||0).toFixed(3), (Number(ap.value)||0).toFixed(3), (Number(ae.value)||0).toFixed(3),
    (Number(z.value)||0).toString(), (Number(fz.value)||0).toFixed(4), (Number(Vc.value)||0).toFixed(2),
    rpm.value, vf.value, mrr.value, pkw.value, torque.value, fc.value,
    (Number(L.value)||0).toFixed(2), tme.value, (Number(pasadas.value)||1).toString(),
    utiles.value || "", refrigeracion.value || "", notas.value || ""
  ];
  cols.forEach(c=>{
    const td = document.createElement("td");
    td.textContent = c;
    row.appendChild(td);
  });
  // Make row selectable (single select)
  row.addEventListener("click", ()=> {
    // unselect others
    document.querySelectorAll("#tbl tbody tr").forEach(r=> r.classList.remove("selected"));
    row.classList.add("selected");
  });

  tblBody.appendChild(row);
  alert("Subfase agregada correctamente.");
}

// ----- Eliminar fila seleccionada -----
function eliminarFila(){
  const selected = document.querySelector("#tbl tbody tr.selected");
  if(!selected){ alert("Selecciona una fila para eliminar."); return; }
  selected.remove();
}

// ----- Limpiar tabla -----
function limpiarTabla(){
  if(confirm("¿Vaciar toda la tabla?")) {
    tblBody.innerHTML = "";
  }
}
async function exportExcel() {
  // Asegurar que ExcelJS está disponible
  if (typeof ExcelJS === "undefined") {
    alert("Error: no se encontró la librería ExcelJS");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Hoja de Ruta");

  // Obtener encabezados de la tabla
  const headers = Array.from(document.querySelectorAll("#tbl thead th")).map(th => th.textContent.trim());
  worksheet.addRow(headers);

  // Agregar filas de datos
  document.querySelectorAll("#tbl tbody tr").forEach(tr => {
    const row = Array.from(tr.children).map(td => td.textContent.trim());
    worksheet.addRow(row);
  });

  // === ESTILOS ===
  const headerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F4A460" } // Naranja similar al de tu imagen
  };
  const headerBorder = {
    top: { style: "thin", color: { argb: "000000" } },
    bottom: { style: "thin", color: { argb: "000000" } },
    left: { style: "thin", color: { argb: "000000" } },
    right: { style: "thin", color: { argb: "000000" } }
  };

  // Aplicar estilo al encabezado
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = { bold: true, color: { argb: "000000" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = headerBorder;
  });

  // Bordes y centrado a las demás filas
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // saltar encabezado
    row.eachCell(cell => {
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin", color: { argb: "808080" } },
        bottom: { style: "thin", color: { argb: "808080" } },
        left: { style: "thin", color: { argb: "808080" } },
        right: { style: "thin", color: { argb: "808080" } }
      };
    });
  });

  // Ajustar ancho de columnas automáticamente
  worksheet.columns.forEach(column => {
    let maxLength = 10;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const len = cell.value ? cell.value.toString().length : 0;
      if (len > maxLength) maxLength = len;
    });
    column.width = maxLength + 2;
  });

  // Descargar el archivo Excel
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "Hoja_de_Ruta_Fresadora.xlsx");
}


// ----- Cargar XLSX (SheetJS) -----
function cargarExcel(file){
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:"array"});
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, {header:1});
    if(json.length < 2){ alert("Archivo vacío o sin filas."); return; }
    // first row assumed header
    tblBody.innerHTML = "";
    for(let i=1;i<json.length;i++){
      const row = json[i];
      const tr = document.createElement("tr");
      row.forEach(cell => {
        const td = document.createElement("td");
        td.textContent = (cell===undefined) ? "" : cell.toString();
        tr.appendChild(td);
      });
      // ensure missing cells are filled to match header count
      while(tr.children.length < json[0].length){
        const td = document.createElement("td"); td.textContent="";
        tr.appendChild(td);
      }
      tr.addEventListener("click", ()=> {
        document.querySelectorAll("#tbl tbody tr").forEach(r=> r.classList.remove("selected"));
        tr.classList.add("selected");
      });
      tblBody.appendChild(tr);
    }
    alert("Archivo cargado correctamente.");
  };
  reader.readAsArrayBuffer(file);
}

// ----- Events -----
document.addEventListener("DOMContentLoaded", ()=>{
  initMateriales();
  initOperacion();

  // Fill default material values into Vc and fz if empty (first load)
  const mat0 = material.value;
  if(mat0 && MATERIALES[mat0]){
    if(!Vc.value) Vc.value = MATERIALES[mat0].Vc;
    if(!fz.value) fz.value = MATERIALES[mat0].fz;
  }

  // Bind input change events to recalc
  [D, z, fz, Vc, ap, ae, L, material].forEach(elm => {
    elm.addEventListener("input", ()=> recalcularPreview());
  });
  operacion.addEventListener("change", ()=>{
    herramienta.value = PARAMS_OPERACION[operacion.value].herramienta;
    if(!z.value) z.value = PARAMS_OPERACION[operacion.value].z || "";
    applyAeRule();
    recalcularPreview();
  });
  material.addEventListener("change", ()=>{
    const m = MATERIALES[material.value];
    if(m){
      if(!Vc.value) Vc.value = m.Vc;
      if(!fz.value) fz.value = m.fz;
    }
    recalcularPreview();
  });

  // Buttons
  btnAgregar.addEventListener("click", agregarFila);
  btnEliminar.addEventListener("click", eliminarFila);
  btnLimpiar.addEventListener("click", limpiarTabla);
  btnGuardar.addEventListener("click", exportExcel);

  btnCargar.addEventListener("click", ()=> inputCargar.click());
  inputCargar.addEventListener("change", (ev)=>{
    const f = ev.target.files[0];
    if(f) cargarExcel(f);
    inputCargar.value = "";
  });

  // initial recalc
  recalcularPreview();
});


