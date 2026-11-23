let baseOperativos = [];
let basePersonas = [];
let baseVehiculos = [];

function cambiarTab(idContenido, elementoClickeado) {
    const tabs = document.querySelectorAll('#tab-headers > div');
    tabs.forEach(tab => {
        tab.classList.remove('tab-active');
        tab.classList.add('tab-inactive');
    });
    elementoClickeado.classList.remove('tab-inactive');
    elementoClickeado.classList.add('tab-active');

    const contenidos = document.querySelectorAll('.tab-content');
    contenidos.forEach(cont => {
        cont.classList.add('hidden');
        cont.classList.remove('block');
    });

    const contenidoMostrar = document.getElementById(idContenido);
    if (contenidoMostrar) {
        contenidoMostrar.classList.remove('hidden');
        contenidoMostrar.classList.add('block');
    }
}

function normalizarDatos(arreglo) {
    return arreglo.map(fila => {
        let filaNormalizada = {};
        for (let clave in fila) {
            let nuevaClave = clave.trim().toUpperCase(); 
            filaNormalizada[nuevaClave] = fila[clave];
        }
        return filaNormalizada;
    });
}

// NUEVO: Funciones robustas para extraer Novedad y Acción de Registro
function obtenerNovedad(fila) {
    return fila['NOVEDAD'] || fila['NOVEDADES'] || fila['OBSERVACION'] || fila['OBSERVACIONES'] || fila['NOVEDAD PERSONA'] || '-';
}

function obtenerAccionRegistro(fila) {
    return fila['ACCION DE REGISTRO'] || fila['ACCIÓN DE REGISTRO'] || fila['TIPO ACCION'] || fila['TIPO ACCIÓN'] || fila['TIPO DE ACCION'] || fila['ACCION REGISTRO'] || fila['ACCION'] || fila['TIPO REGISTRO'] || '-';
}

function extraerResponsable(fila) {
    if (fila['USUARIO SIIPNE']) return String(fila['USUARIO SIIPNE']);
    if (fila['SIIPNE']) return String(fila['SIIPNE']);
    if (fila['QUIEN HACE OPERATIVO']) return String(fila['QUIEN HACE OPERATIVO']);
    if (fila['RESPONSABLE']) return String(fila['RESPONSABLE']);
    if (fila['SERVIDOR POLICIAL']) return String(fila['SERVIDOR POLICIAL']);
    if (fila['USUARIO']) return String(fila['USUARIO']);
    
    const keys = Object.keys(fila);
    const key = keys.find(k => k.includes('SIIPNE') || k.includes('RESPONSABLE') || k.includes('QUIEN') || k.includes('SERVIDOR') || k.includes('USUARIO'));
    
    return key && fila[key] ? String(fila[key]) : 'NO ESPECIFICADO';
}

function parseExcelDate(valor) {
    if (!valor) return null;
    let d;
    if (valor instanceof Date) {
        d = new Date(valor.getTime());
    } else if (typeof valor === 'number') {
        d = new Date(Math.round((valor - 25569) * 86400 * 1000));
        d = new Date(d.getTime() + d.getTimezoneOffset() * 60000);
    } else if (typeof valor === 'string') {
        const txt = valor.trim();
        if (txt.match(/^\d{4}-\d{2}-\d{2}$/)) { 
            const [y, m, day] = txt.split('-');
            d = new Date(y, m - 1, day);
        } else if (txt.match(/^\d{2}\/\d{2}\/\d{4}/)) { 
            const partes = txt.split(/[\/\-:\s]/); 
            d = new Date(partes[2], partes[1] - 1, partes[0]);
        } else {
            d = new Date(txt);
        }
    }
    if (!d || isNaN(d.getTime())) return null;
    return d;
}

function obtenerHora(fila, dateObj) {
    if (dateObj && (dateObj.getHours() > 0 || dateObj.getMinutes() > 0 || dateObj.getSeconds() > 0)) {
        return String(dateObj.getHours()).padStart(2, '0') + ':' + String(dateObj.getMinutes()).padStart(2, '0');
    }
    
    const horaVal = fila['HORA'] || fila['HORA OPERATIVO'];
    if (horaVal !== undefined && horaVal !== '') {
        if (typeof horaVal === 'number') {
            let totalMinutos = Math.round(horaVal * 24 * 60);
            let h = Math.floor(totalMinutos / 60);
            let m = totalMinutos % 60;
            return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
        } else if (typeof horaVal === 'string') {
            let partes = horaVal.trim().split(':');
            if (partes.length >= 2) return `${partes[0].padStart(2, '0')}:${partes[1].padStart(2, '0')}`;
            return horaVal;
        }
    }
    return '00:00';
}

document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellDates: true});
            const nombresHojas = workbook.SheetNames;
            
            const nombreBase = nombresHojas.find(n => n.toUpperCase().includes('BASE') || n.toUpperCase().includes('SIIPNE')) || nombresHojas[0];
            const nombrePersonas = nombresHojas.find(n => n.toUpperCase().includes('PERSONAS'));
            const nombreVehiculos = nombresHojas.find(n => n.toUpperCase().includes('VEHICULOS') || n.toUpperCase().includes('VEHÍCULOS'));

            if (nombreBase) {
                baseOperativos = normalizarDatos(XLSX.utils.sheet_to_json(workbook.Sheets[nombreBase], { defval: "" }));
            }
            if (nombrePersonas) {
                basePersonas = normalizarDatos(XLSX.utils.sheet_to_json(workbook.Sheets[nombrePersonas], { defval: "" }));
            }
            if (nombreVehiculos) {
                baseVehiculos = normalizarDatos(XLSX.utils.sheet_to_json(workbook.Sheets[nombreVehiculos], { defval: "" }));
            }
            
            document.getElementById('fileName').innerText = file.name;
            const now = new Date();
            document.getElementById('fileDate').innerText = now.toISOString().split('T')[0] + ' ' + String(now.getHours()).padStart(2,'0') + ':' + String(now.getMinutes()).padStart(2,'0');
            document.getElementById('fileCount').innerText = baseOperativos.length;
            document.getElementById('fileInfo').classList.remove('hidden');
            document.getElementById('noFile').classList.add('hidden');

            inicializarFiltros();
            aplicarFiltros(); 
            
        } catch (error) {
            console.error("Error procesando Excel:", error);
            alert("No se pudo extraer la información del Excel. Revisa el archivo.");
        }
    };
    reader.readAsArrayBuffer(file);
});

function borrarDatos() {
    baseOperativos = [];
    basePersonas = [];
    baseVehiculos = [];
    document.getElementById('fileInput').value = "";
    document.getElementById('fileInfo').classList.add('hidden');
    document.getElementById('noFile').classList.remove('hidden');
    document.getElementById('tabla-responsables').innerHTML = '';
    document.getElementById('tabla-vehiculos').innerHTML = '';
    document.getElementById('tabla-personas').innerHTML = '';
    
    if (document.getElementById('buscadorResponsable')) document.getElementById('buscadorResponsable').value = '';
    if (document.getElementById('buscadorSiipne')) document.getElementById('buscadorSiipne').value = '';
    
    document.getElementById('tabla-consulta-siipne').innerHTML = '<div class="p-4 text-slate-400 text-center text-sm">Empieza a escribir para buscar...</div>';
    quitarFiltros();
    calcularMetricas([]);
}

function inicializarFiltros() {
    const subzonas = new Set(['Todas']);
    const tipos = new Set(['Todos']);

    baseOperativos.forEach(fila => {
        if (fila['SUBZONA']) subzonas.add(String(fila['SUBZONA']).trim());
        if (fila['TIPO DE OPERATIVO']) tipos.add(String(fila['TIPO DE OPERATIVO']).trim());
    });

    llenarSelect('filtroSubzona', Array.from(subzonas));
    llenarSelect('filtroTipo', Array.from(tipos));
    actualizarDistritos();
}

function cambioSubzona() {
    actualizarDistritos();
    aplicarFiltros();
}

function actualizarDistritos() {
    const subzonaSeleccionada = document.getElementById('filtroSubzona').value;
    const distritos = new Set(['Todos']);
    
    baseOperativos.forEach(fila => {
        const valSubzona = fila['SUBZONA'] ? String(fila['SUBZONA']).trim() : "";
        const valDistrito = fila['DISTRITO'] ? String(fila['DISTRITO']).trim() : "";
        if (subzonaSeleccionada === 'Todas' || valSubzona === subzonaSeleccionada) {
            if (valDistrito && valDistrito !== "") distritos.add(valDistrito);
        }
    });

    const currentDistrito = document.getElementById('filtroDistrito').value;
    llenarSelect('filtroDistrito', Array.from(distritos));
    document.getElementById('filtroDistrito').value = distritos.has(currentDistrito) ? currentDistrito : 'Todos';
}

function llenarSelect(id, opciones) {
    const select = document.getElementById(id);
    if (!select) return;
    select.innerHTML = '';
    opciones.forEach(opc => {
        if(opc && opc !== "undefined" && opc !== "") select.innerHTML += `<option value="${opc}">${opc}</option>`;
    });
}

function quitarFiltros() {
    if(document.getElementById('filtroSubzona')) document.getElementById('filtroSubzona').value = 'Todas';
    actualizarDistritos(); 
    if(document.getElementById('filtroDistrito')) document.getElementById('filtroDistrito').value = 'Todos';
    if(document.getElementById('filtroTipo')) document.getElementById('filtroTipo').value = 'Todos';
    if(document.getElementById('filtroDesde')) document.getElementById('filtroDesde').value = '';
    if(document.getElementById('filtroHasta')) document.getElementById('filtroHasta').value = '';
    aplicarFiltros();
}

function aplicarFiltros() {
    const subzonaEl = document.getElementById('filtroSubzona');
    if(!subzonaEl) return;
    
    const subzona = subzonaEl.value;
    const distrito = document.getElementById('filtroDistrito').value;
    const tipo = document.getElementById('filtroTipo').value;
    
    let strDesde = document.getElementById('filtroDesde').value;
    let strHasta = document.getElementById('filtroHasta').value;
    let dateDesde = strDesde ? parseExcelDate(strDesde) : null;
    let dateHasta = strHasta ? parseExcelDate(strHasta) : null;
    
    if (dateDesde) dateDesde.setHours(0, 0, 0, 0);
    if (dateHasta) dateHasta.setHours(23, 59, 59, 999);

    const operativosFiltrados = baseOperativos.filter(fila => {
        const valSubzona = fila['SUBZONA'] ? String(fila['SUBZONA']).trim() : "";
        const valDistrito = fila['DISTRITO'] ? String(fila['DISTRITO']).trim() : "";
        const valTipo = fila['TIPO DE OPERATIVO'] ? String(fila['TIPO DE OPERATIVO']).trim() : "";

        const cumpleSub = (subzona === 'Todas') || (valSubzona === subzona);
        const cumpleDist = (distrito === 'Todos') || (valDistrito === distrito);
        const cumpleTipo = (tipo === 'Todos') || (valTipo === tipo);
        
        let cumpleFecha = true;
        if (dateDesde || dateHasta) {
            let dOp = parseExcelDate(fila['FECHA HORA'] || fila['FECHA OPERATIVO'] || fila['FECHA']);
            if (dOp) {
                if (dateDesde && dOp.getTime() < dateDesde.getTime()) cumpleFecha = false;
                if (dateHasta && dOp.getTime() > dateHasta.getTime()) cumpleFecha = false;
            } else {
                cumpleFecha = false; 
            }
        }
        return cumpleSub && cumpleDist && cumpleTipo && cumpleFecha;
    });

    calcularMetricas(operativosFiltrados);
}

function formatearFechaEspecial(valor) {
    if (!valor) return '-';
    if (typeof valor === 'number') {
        const d = parseExcelDate(valor);
        return d ? d.toISOString().split('T')[0] : valor;
    }
    if (valor instanceof Date) return valor.toISOString().split('T')[0];
    return String(valor).substring(0, 10);
}

function calcularMetricas(operativosFiltrados) {
    const idsValidos = new Set(operativosFiltrados.map(op => String(op['FROT'] || op['ID OPERATIVO']).trim()));
    const subzonasUnicas = new Set();
    const distritosUnicos = new Set();
    const servidoresUnicos = new Set();
    
    const contadorPersonasFrot = {};
    const contadorVehiculosFrot = {};
    const personasUnicas = new Set();
    const vehiculosUnicos = new Set();
    
    const registrosPersonas = {}; 
    const registrosVehiculos = {}; 
    
    let totalPersonas = 0;
    let totalVehiculos = 0;

    basePersonas.forEach(persona => {
        const frot = String(persona['ID OPERATIVO'] || persona['FROT']).trim();
        if (idsValidos.has(frot)) {
            contadorPersonasFrot[frot] = (contadorPersonasFrot[frot] || 0) + 1;
            if (persona['CEDULA']) {
                totalPersonas++;
                const cedula = String(persona['CEDULA']).trim();
                personasUnicas.add(cedula);
                
                if (!registrosPersonas[cedula]) registrosPersonas[cedula] = [];
                registrosPersonas[cedula].push(persona);
            }
        }
    });

    baseVehiculos.forEach(vehiculo => {
        const frot = String(vehiculo['ID OPERATIVO'] || vehiculo['FROT']).trim();
        if (idsValidos.has(frot)) {
            contadorVehiculosFrot[frot] = (contadorVehiculosFrot[frot] || 0) + 1;
            if (vehiculo['PLACA'] || vehiculo['CHASIS']) {
                totalVehiculos++;
                const placa = String(vehiculo['PLACA'] || vehiculo['CHASIS']).trim();
                vehiculosUnicos.add(placa);
                
                if (!registrosVehiculos[placa]) registrosVehiculos[placa] = [];
                registrosVehiculos[placa].push(vehiculo);
            }
        }
    });
    
    const agrupadosResponsables = {};
    const infoOperativos = {}; 
    let minDate = null;
    let maxDate = null;

    operativosFiltrados.forEach(fila => {
        const frot = String(fila['FROT'] || fila['ID OPERATIVO']).trim();
        if (fila['SUBZONA']) subzonasUnicas.add(String(fila['SUBZONA']).trim());
        if (fila['DISTRITO']) distritosUnicos.add(String(fila['DISTRITO']).trim());
        
        let resp = extraerResponsable(fila).trim();
        const fechaCruda = fila['FECHA HORA'] || fila['FECHA OPERATIVO'] || fila['FECHA'];
        const d = parseExcelDate(fechaCruda);
        
        const fechaSolo = d ? d.toISOString().split('T')[0] : '-';
        const horaDetectada = obtenerHora(fila, d);
        const fechaStr = `${fechaSolo} ${horaDetectada}`;

        infoOperativos[frot] = {
            fecha: fechaStr,
            responsable: resp && resp !== 'undefined' && resp !== '' ? resp : 'NO ESPECIFICADO'
        };

        if (resp && resp !== 'undefined' && resp !== '' && resp !== 'NO ESPECIFICADO') {
            servidoresUnicos.add(resp);
            
            if (!agrupadosResponsables[resp]) {
                agrupadosResponsables[resp] = {
                    documento: fila['DOCUMENTO'] || fila['CEDULA'] || fila['IDENTIFICACION'] || '-',
                    subzona: fila['SUBZONA'] || '-',
                    distrito: fila['DISTRITO'] || '-',
                    detalles: []
                };
            }
            
            const personas = fila['PERSONAS'] || fila['CANTIDAD PERSONAS'] || contadorPersonasFrot[frot] || '0';
            const vehiculos = fila['VEHICULOS'] || fila['VEHÍCULOS'] || contadorVehiculosFrot[frot] || '0';
            const lat = fila['LATITUD'] || fila['LAT'] || '-';
            const lng = fila['LONGITUD'] || fila['LON'] || fila['LONG'] || '-';

            agrupadosResponsables[resp].detalles.push({
                frot: frot || '-', fechaStr, personas, vehiculos, lat, lng
            });
        }
        
        let dOp = parseExcelDate(fechaCruda);
        if (dOp) {
            if (!minDate || dOp < minDate) minDate = dOp;
            if (!maxDate || dOp > maxDate) maxDate = dOp;
        }
    });

    let htmlResp = `<div class="max-h-[500px] overflow-y-auto pr-2 rounded-lg border border-slate-700 custom-scrollbar">
        <table class="w-full text-left text-sm text-slate-300 relative">
            <thead class="text-xs text-slate-400 uppercase bg-slate-900 sticky top-0 z-10 shadow-md">
                <tr>
                    <th class="px-4 py-3 font-semibold min-w-[200px]">Persona Responsable</th>
                    <th class="px-4 py-3 font-semibold">Documento</th>
                    <th class="px-4 py-3 font-semibold">Subzona</th>
                    <th class="px-4 py-3 font-semibold">Distrito</th>
                    <th class="px-4 py-3 font-semibold text-center">Operativos</th>
                </tr>
            </thead>
            <tbody class="divide-y divide-slate-700/50">`;

    const listaAgrupada = Object.entries(agrupadosResponsables).sort((a, b) => b[1].detalles.length - a[1].detalles.length);

    listaAgrupada.forEach(([resp, data], idx) => {
        htmlResp += `<tr class="hover:bg-slate-700/70 transition-colors cursor-pointer bg-slate-800" onclick="document.getElementById('detalle-resp-${idx}').classList.toggle('hidden')">
            <td class="px-4 py-3 font-medium text-white flex items-center gap-2">
                <span class="text-[0.6rem] bg-blue-500/20 text-blue-400 px-1.5 py-0.5 rounded border border-blue-500/30">▼</span>
                ${resp}
            </td>
            <td class="px-4 py-3">${data.documento}</td>
            <td class="px-4 py-3">${data.subzona}</td>
            <td class="px-4 py-3">${data.distrito}</td>
            <td class="px-4 py-3 text-center text-blue-400 font-bold">${data.detalles.length}</td>
        </tr>`;

        htmlResp += `<tr id="detalle-resp-${idx}" class="hidden bg-slate-800/40">
            <td colspan="5" class="p-4">
                <div class="bg-slate-900 rounded-md border border-slate-700 overflow-hidden shadow-inner">
                    <table class="w-full text-left text-xs text-slate-300">
                        <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800/80 border-b border-slate-700">
                            <tr>
                                <th class="px-3 py-2 font-semibold">FROT</th>
                                <th class="px-3 py-2 font-semibold">Fecha HORA Operativo</th>
                                <th class="px-3 py-2 font-semibold text-center">Personas</th>
                                <th class="px-3 py-2 font-semibold text-center">Vehículos</th>
                                <th class="px-3 py-2 font-semibold">Latitud</th>
                                <th class="px-3 py-2 font-semibold">Longitud</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-700/30">`;
        
        data.detalles.forEach(det => {
            htmlResp += `<tr class="hover:bg-slate-700/50 transition-colors">
                <td class="px-3 py-2 font-mono text-blue-300">${det.frot}</td>
                <td class="px-3 py-2 whitespace-nowrap">${det.fechaStr}</td>
                <td class="px-3 py-2 text-center font-bold text-emerald-400">${det.personas}</td>
                <td class="px-3 py-2 text-center font-bold text-amber-400">${det.vehiculos}</td>
                <td class="px-3 py-2 font-mono text-slate-400">${det.lat}</td>
                <td class="px-3 py-2 font-mono text-slate-400">${det.lng}</td>
            </tr>`;
        });

        htmlResp += `       </tbody>
                    </table>
                </div>
            </td>
        </tr>`;
    });
    
    htmlResp += `</tbody></table></div>`;
    if(document.getElementById('tabla-responsables')) document.getElementById('tabla-responsables').innerHTML = listaAgrupada.length ? htmlResp : '<div class="p-4 text-slate-400 text-center text-sm">No hay datos de responsables</div>';

    let vehiculosDup = Object.entries(registrosVehiculos)
        .filter(([placa, registros]) => registros.length > 1)
        .sort((a, b) => b[1].length - a[1].length);

    let htmlVeh = `<div class="max-h-[500px] overflow-y-auto pr-2 rounded-lg border border-slate-700 custom-scrollbar">
        <table class="w-full text-left text-sm text-slate-300 relative">
            <thead class="text-xs text-slate-400 uppercase bg-slate-900 sticky top-0 z-10 shadow-md">
                <tr>
                    <th class="px-4 py-3 font-semibold">Placa / Chasis</th>
                    <th class="px-4 py-3 font-semibold w-32 text-center">Registrado en</th>
                </tr>
            </thead>
            <tbody class="divide-y divide-slate-700/50">`;

    vehiculosDup.forEach(([placa, registros], idx) => { 
        htmlVeh += `<tr class="hover:bg-slate-700/70 transition-colors cursor-pointer bg-slate-800" onclick="document.getElementById('detalle-veh-${idx}').classList.toggle('hidden')">
            <td class="px-4 py-3 font-medium text-white flex items-center gap-2">
                <span class="text-[0.6rem] bg-orange-500/20 text-orange-400 px-1.5 py-0.5 rounded border border-orange-500/30">▼</span>
                ${placa}
            </td>
            <td class="px-4 py-3 text-center font-bold text-orange-400">${registros.length} operativos</td>
        </tr>`;

        htmlVeh += `<tr id="detalle-veh-${idx}" class="hidden bg-slate-800/40">
            <td colspan="2" class="p-4">
                <div class="bg-slate-900 rounded-md border border-slate-700 overflow-x-auto shadow-inner">
                    <table class="w-full text-left text-xs text-slate-300 whitespace-nowrap">
                        <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800/80 border-b border-slate-700">
                            <tr>
                                <th class="px-3 py-2 font-semibold">Fecha HORA</th>
                                <th class="px-3 py-2 font-semibold">Responsable</th>
                                <th class="px-3 py-2 font-semibold">ID Operativo</th>
                                <th class="px-3 py-2 font-semibold">Motor</th>
                                <th class="px-3 py-2 font-semibold">Chasis</th>
                                <th class="px-3 py-2 font-semibold">Marca</th>
                                <th class="px-3 py-2 font-semibold">Modelo</th>
                                <th class="px-3 py-2 font-semibold">Clase</th>
                                <th class="px-3 py-2 font-semibold">Acción Reg.</th>
                                <th class="px-3 py-2 font-semibold">Novedad</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-700/30">`;
        
        registros.forEach(reg => {
            const idOp = String(reg['ID OPERATIVO'] || reg['FROT']).trim();
            const info = infoOperativos[idOp] || { fecha: '-', responsable: '-' };
            
            htmlVeh += `<tr class="hover:bg-slate-700/50 transition-colors">
                <td class="px-3 py-2 text-slate-300">${info.fecha}</td>
                <td class="px-3 py-2 font-medium text-white">${info.responsable}</td>
                <td class="px-3 py-2 font-mono text-blue-300">${idOp}</td>
                <td class="px-3 py-2">${reg['MOTOR'] || '-'}</td>
                <td class="px-3 py-2">${reg['CHASIS'] || '-'}</td>
                <td class="px-3 py-2">${reg['MARCA'] || '-'}</td>
                <td class="px-3 py-2">${reg['MODELO'] || '-'}</td>
                <td class="px-3 py-2">${reg['CLASE'] || '-'}</td>
                <td class="px-3 py-2">${obtenerAccionRegistro(reg)}</td>
                <td class="px-3 py-2">${obtenerNovedad(reg)}</td>
            </tr>`;
        });

        htmlVeh += `       </tbody>
                    </table>
                </div>
            </td>
        </tr>`;
    });
    
    htmlVeh += `</tbody></table></div>`;
    if(document.getElementById('tabla-vehiculos')) document.getElementById('tabla-vehiculos').innerHTML = vehiculosDup.length ? htmlVeh : '<div class="p-4 text-slate-400 text-center text-sm">No hay vehículos duplicados</div>';

    let personasDup = Object.entries(registrosPersonas)
        .filter(([ced, registros]) => registros.length > 1)
        .sort((a, b) => b[1].length - a[1].length);

    let htmlPer = `<div class="max-h-[500px] overflow-y-auto pr-2 rounded-lg border border-slate-700 custom-scrollbar">
        <table class="w-full text-left text-sm text-slate-300 relative">
            <thead class="text-xs text-slate-400 uppercase bg-slate-900 sticky top-0 z-10 shadow-md">
                <tr>
                    <th class="px-4 py-3 font-semibold">Cédula</th>
                    <th class="px-4 py-3 font-semibold">Nombres</th>
                    <th class="px-4 py-3 font-semibold w-32 text-center">Registrado en</th>
                </tr>
            </thead>
            <tbody class="divide-y divide-slate-700/50">`;

    personasDup.forEach(([ced, registros], idx) => {
        const nombreP = registros[0]['NOMBRES'] || '-';
        
        htmlPer += `<tr class="hover:bg-slate-700/70 transition-colors cursor-pointer bg-slate-800" onclick="document.getElementById('detalle-per-${idx}').classList.toggle('hidden')">
            <td class="px-4 py-3 font-medium text-white flex items-center gap-2">
                <span class="text-[0.6rem] bg-emerald-500/20 text-emerald-400 px-1.5 py-0.5 rounded border border-emerald-500/30">▼</span>
                ${ced}
            </td>
            <td class="px-4 py-3 text-slate-300 truncate max-w-xs">${nombreP}</td>
            <td class="px-4 py-3 text-center font-bold text-emerald-400">${registros.length} operativos</td>
        </tr>`;

        htmlPer += `<tr id="detalle-per-${idx}" class="hidden bg-slate-800/40">
            <td colspan="3" class="p-4">
                <div class="bg-slate-900 rounded-md border border-slate-700 overflow-x-auto shadow-inner">
                    <table class="w-full text-left text-xs text-slate-300 whitespace-nowrap">
                        <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800/80 border-b border-slate-700">
                            <tr>
                                <th class="px-3 py-2 font-semibold">Fecha HORA (Principal)</th>
                                <th class="px-3 py-2 font-semibold">Responsable</th>
                                <th class="px-3 py-2 font-semibold">ID Operativo</th>
                                <th class="px-3 py-2 font-semibold">Sexo</th>
                                <th class="px-3 py-2 font-semibold">Fecha Nac.</th>
                                <th class="px-3 py-2 font-semibold">Tipo Acción</th>
                                <th class="px-3 py-2 font-semibold">Condición Reg.</th>
                                <th class="px-3 py-2 font-semibold">Novedad</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-700/30">`;
        
        registros.forEach(reg => {
            const idOp = String(reg['ID OPERATIVO'] || reg['FROT']).trim();
            const info = infoOperativos[idOp] || { fecha: '-', responsable: '-' };
            
            const fechaNac = formatearFechaEspecial(reg['FECHA DE NACIMIENTO'] || reg['FECHA NACIMIENTO'] || reg['A DE NACIMIENTO'] || reg['A DE NACIMIE'] || '-');
            const condRegistro = reg['CONDICION DE REGISTRO'] || reg['CONDICIÓN DE REGISTRO'] || reg['ION DE REGISTRO'] || reg['ION DE REGIS'] || '-';
            
            htmlPer += `<tr class="hover:bg-slate-700/50 transition-colors">
                <td class="px-3 py-2 text-slate-300">${info.fecha}</td>
                <td class="px-3 py-2 font-medium text-white">${info.responsable}</td>
                <td class="px-3 py-2 font-mono text-blue-300">${idOp}</td>
                <td class="px-3 py-2">${reg['SEXO'] || '-'}</td>
                <td class="px-3 py-2">${fechaNac}</td>
                <td class="px-3 py-2">${obtenerAccionRegistro(reg)}</td>
                <td class="px-3 py-2">${condRegistro}</td>
                <td class="px-3 py-2">${obtenerNovedad(reg)}</td>
            </tr>`;
        });

        htmlPer += `       </tbody>
                    </table>
                </div>
            </td>
        </tr>`;
    });
    
    htmlPer += `</tbody></table></div>`;
    if(document.getElementById('tabla-personas')) document.getElementById('tabla-personas').innerHTML = personasDup.length ? htmlPer : '<div class="p-4 text-slate-400 text-center text-sm">No hay personas duplicadas</div>';

    const fmt = (num) => new Intl.NumberFormat('es-ES').format(num || 0);
    const fmtDate = (d) => d ? d.toISOString().split('T')[0] : '--';

    if(document.getElementById('valTotal')) document.getElementById('valTotal').innerText = fmt(operativosFiltrados.length);
    if(document.getElementById('valFechas')) document.getElementById('valFechas').innerText = (minDate && maxDate) ? `${fmtDate(minDate)} a ${fmtDate(maxDate)}` : '--';
    if(document.getElementById('valZona')) document.getElementById('valZona').innerText = fmt(operativosFiltrados.length); 
    if(document.getElementById('valSubzonas')) document.getElementById('valSubzonas').innerText = fmt(subzonasUnicas.size);
    if(document.getElementById('valDistritos')) document.getElementById('valDistritos').innerText = fmt(distritosUnicos.size);
    if(document.getElementById('valServidores')) document.getElementById('valServidores').innerText = fmt(servidoresUnicos.size);
    if(document.getElementById('valPersonas')) document.getElementById('valPersonas').innerText = fmt(totalPersonas);
    if(document.getElementById('valPersonasUnicas')) document.getElementById('valPersonasUnicas').innerText = `${fmt(personasUnicas.size)} cédulas únicas`;
    if(document.getElementById('valVehiculos')) document.getElementById('valVehiculos').innerText = fmt(totalVehiculos);
    if(document.getElementById('valVehiculosUnicos')) document.getElementById('valVehiculosUnicos').innerText = `${fmt(vehiculosUnicos.size)} placas únicas`;
    if(document.getElementById('valPerDuplicadas')) document.getElementById('valPerDuplicadas').innerText = fmt(personasDup.length);
    if(document.getElementById('valVehDuplicados')) document.getElementById('valVehDuplicados').innerText = fmt(vehiculosDup.length);
}

// --- LÓGICA DE BÚSQUEDA Y VISTAS SIIPNE ---

function ejecutarBusqueda(origen) {
    const inputSiipne = document.getElementById('buscadorSiipne');
    const inputNombre = document.getElementById('buscadorResponsable');
    const contenedor = document.getElementById('tabla-consulta-siipne');

    const valSiipne = inputSiipne ? inputSiipne.value.trim() : '';
    const valNombre = inputNombre ? inputNombre.value.trim().toLowerCase() : '';

    if (origen === 'siipne' && valSiipne !== '') {
        const op = baseOperativos.find(f => String(f['FROT'] || f['ID OPERATIVO'] || f['N° SIIPNE']).trim() === valSiipne);
        if (op) {
            const responsable = extraerResponsable(op);
            if (inputNombre && responsable !== 'NO ESPECIFICADO') {
                inputNombre.value = responsable; 
            }
            renderizarDetalleOperativo(valSiipne, op);
        } else {
            contenedor.innerHTML = `<div class="p-4 bg-[#0f172a] border border-slate-700 rounded text-slate-400 text-center text-sm">No se encontró el operativo N° ${valSiipne}</div>`;
        }
    } else if (origen === 'nombre' && valNombre !== '') {
        if(inputSiipne) inputSiipne.value = '';
        renderizarResumenPersona(valNombre);
    } else {
        contenedor.innerHTML = '<div class="p-4 text-slate-400 text-center text-sm">Ingresa un dato para buscar...</div>';
    }
}

function irADetalleOperativo(frot) {
    const inputSiipne = document.getElementById('buscadorSiipne');
    if (inputSiipne) {
        inputSiipne.value = frot;
        ejecutarBusqueda('siipne');
    }
}

// Pestañas internas para la vista de Resumen de Persona
window.cambiarTabInterno = function(tabId) {
    ['ops', 'pers', 'veh'].forEach(t => {
        const btn = document.getElementById('btn-tab-' + t);
        const content = document.getElementById('content-tab-' + t);
        if (btn && content) {
            if (t === tabId) {
                btn.className = "bg-blue-500/20 border border-blue-500/50 text-blue-400 px-4 py-1.5 rounded-full text-xs font-semibold transition-colors";
                content.classList.remove('hidden');
            } else {
                btn.className = "border border-slate-700 text-slate-400 px-4 py-1.5 rounded-full text-xs hover:border-slate-500 hover:text-slate-300 transition-colors cursor-pointer";
                content.classList.add('hidden');
            }
        }
    });
};

function renderizarResumenPersona(nombreBusqueda) {
    const contenedor = document.getElementById('tabla-consulta-siipne');
    
    // Buscar todos los operativos de este responsable
    const opsResponsable = baseOperativos.filter(fila => {
        const resp = extraerResponsable(fila).toLowerCase();
        return resp.includes(nombreBusqueda);
    });

    if (opsResponsable.length === 0) {
        contenedor.innerHTML = `<div class="p-4 bg-[#0f172a] border border-slate-700 rounded text-slate-400 text-center text-sm">No se encontraron operativos para el responsable que contenga "${nombreBusqueda}"</div>`;
        return;
    }

    const nombreReal = extraerResponsable(opsResponsable[0]);
    const idsOperativos = new Set(opsResponsable.map(op => String(op['FROT'] || op['ID OPERATIVO'] || op['N° SIIPNE']).trim()));

    // Filtrar personas y vehículos de esos operativos
    const personasAsociadas = basePersonas.filter(p => idsOperativos.has(String(p['ID OPERATIVO'] || p['FROT']).trim()));
    const vehiculosAsociados = baseVehiculos.filter(v => idsOperativos.has(String(v['ID OPERATIVO'] || v['FROT']).trim()));

    let html = `
    <div class="mb-6 bg-slate-900 border border-slate-700 p-4 rounded-lg">
        <h3 class="text-xl font-bold text-white mb-2">Resumen: ${nombreReal}</h3>
        <div class="flex gap-4 mb-4">
            <div class="bg-slate-800 p-3 rounded flex-1 text-center border border-slate-700">
                <span class="block text-xl font-bold text-blue-400">${opsResponsable.length}</span>
                <span class="text-xs text-slate-400 uppercase">Operativos</span>
            </div>
            <div class="bg-slate-800 p-3 rounded flex-1 text-center border border-slate-700">
                <span class="block text-xl font-bold text-emerald-400">${personasAsociadas.length}</span>
                <span class="text-xs text-slate-400 uppercase">Personas Reg.</span>
            </div>
            <div class="bg-slate-800 p-3 rounded flex-1 text-center border border-slate-700">
                <span class="block text-xl font-bold text-amber-400">${vehiculosAsociados.length}</span>
                <span class="text-xs text-slate-400 uppercase">Vehículos Reg.</span>
            </div>
        </div>

        <div class="flex gap-2 border-b border-slate-700 pb-3 mb-4" id="tabs-internos">
            <button id="btn-tab-ops" onclick="cambiarTabInterno('ops')" class="bg-blue-500/20 border border-blue-500/50 text-blue-400 px-4 py-1.5 rounded-full text-xs font-semibold transition-colors">Operativos (${opsResponsable.length})</button>
            <button id="btn-tab-pers" onclick="cambiarTabInterno('pers')" class="border border-slate-700 text-slate-400 px-4 py-1.5 rounded-full text-xs hover:border-slate-500 hover:text-slate-300 transition-colors cursor-pointer">Personas (${personasAsociadas.length})</button>
            <button id="btn-tab-veh" onclick="cambiarTabInterno('veh')" class="border border-slate-700 text-slate-400 px-4 py-1.5 rounded-full text-xs hover:border-slate-500 hover:text-slate-300 transition-colors cursor-pointer">Vehículos (${vehiculosAsociados.length})</button>
        </div>

        <!-- Contenido Pestaña Operativos -->
        <div id="content-tab-ops" class="block">
            <div class="overflow-x-auto bg-slate-900 rounded-md border border-slate-700 shadow-inner max-h-[300px] custom-scrollbar">
                <table class="w-full text-left text-xs text-slate-300 whitespace-nowrap relative min-w-max">
                    <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800 border-b border-slate-700 sticky top-0 z-10">
                        <tr>
                            <th class="px-3 py-2 font-semibold">FROT / SIIPNE</th>
                            <th class="px-3 py-2 font-semibold">Fecha y Hora</th>
                            <th class="px-3 py-2 font-semibold">Subzona</th>
                            <th class="px-3 py-2 font-semibold">Distrito</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-700/30">`;
    opsResponsable.forEach(op => {
        const idOp = String(op['FROT'] || op['ID OPERATIVO'] || op['N° SIIPNE'] || '-');
        const d = parseExcelDate(op['FECHA HORA'] || op['FECHA OPERATIVO'] || op['FECHA']);
        const fechaStr = (d ? d.toISOString().split('T')[0] : '-') + ' ' + obtenerHora(op, d);
        html += `
        <tr class="hover:bg-slate-700/50 transition-colors cursor-pointer" onclick="irADetalleOperativo('${idOp}')" title="Clic para ver detalle de este operativo">
            <td class="px-3 py-2 font-mono text-blue-400 underline decoration-blue-400/30 underline-offset-2">${idOp}</td>
            <td class="px-3 py-2">${fechaStr}</td>
            <td class="px-3 py-2">${op['SUBZONA'] || '-'}</td>
            <td class="px-3 py-2">${op['DISTRITO'] || '-'}</td>
        </tr>`;
    });
    html += `</tbody></table></div></div>

        <!-- Contenido Pestaña Personas -->
        <div id="content-tab-pers" class="hidden">
            <div class="overflow-x-auto bg-slate-900 rounded-md border border-slate-700 shadow-inner max-h-[300px] custom-scrollbar">
                <table class="w-full text-left text-xs text-slate-300 whitespace-nowrap relative min-w-max">
                    <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800 border-b border-slate-700 sticky top-0 z-10">
                        <tr>
                            <th class="px-3 py-2 font-semibold">ID Operativo</th>
                            <th class="px-3 py-2 font-semibold">Cédula</th>
                            <th class="px-3 py-2 font-semibold">Nombres</th>
                            <th class="px-3 py-2 font-semibold">Acción Reg.</th>
                            <th class="px-3 py-2 font-semibold">Novedad</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-700/30">`;
    if(personasAsociadas.length === 0) html += `<tr><td colspan="5" class="px-3 py-4 text-center text-slate-500">Sin personas</td></tr>`;
    personasAsociadas.forEach(p => {
        const idOp = String(p['ID OPERATIVO'] || p['FROT'] || '-');
        html += `
        <tr class="hover:bg-slate-700/50 transition-colors">
            <td class="px-3 py-2 font-mono text-blue-300 cursor-pointer hover:underline" onclick="irADetalleOperativo('${idOp}')">${idOp}</td>
            <td class="px-3 py-2 font-mono text-white">${p['CEDULA'] || p['IDENTIFICACION'] || '-'}</td>
            <td class="px-3 py-2">${p['NOMBRES'] || '-'}</td>
            <td class="px-3 py-2">${obtenerAccionRegistro(p)}</td>
            <td class="px-3 py-2">${obtenerNovedad(p)}</td>
        </tr>`;
    });
    html += `</tbody></table></div></div>

        <!-- Contenido Pestaña Vehículos -->
        <div id="content-tab-veh" class="hidden">
            <div class="overflow-x-auto bg-slate-900 rounded-md border border-slate-700 shadow-inner max-h-[300px] custom-scrollbar">
                <table class="w-full text-left text-xs text-slate-300 whitespace-nowrap relative min-w-max">
                    <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800 border-b border-slate-700 sticky top-0 z-10">
                        <tr>
                            <th class="px-3 py-2 font-semibold">ID Operativo</th>
                            <th class="px-3 py-2 font-semibold">Placa/Chasis</th>
                            <th class="px-3 py-2 font-semibold">Marca/Modelo</th>
                            <th class="px-3 py-2 font-semibold">Acción Reg.</th>
                            <th class="px-3 py-2 font-semibold">Novedad</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-700/30">`;
    if(vehiculosAsociados.length === 0) html += `<tr><td colspan="5" class="px-3 py-4 text-center text-slate-500">Sin vehículos</td></tr>`;
    vehiculosAsociados.forEach(v => {
        const idOp = String(v['ID OPERATIVO'] || v['FROT'] || '-');
        html += `
        <tr class="hover:bg-slate-700/50 transition-colors">
            <td class="px-3 py-2 font-mono text-blue-300 cursor-pointer hover:underline" onclick="irADetalleOperativo('${idOp}')">${idOp}</td>
            <td class="px-3 py-2 font-mono text-white">${v['PLACA'] || v['CHASIS'] || '-'}</td>
            <td class="px-3 py-2">${(v['MARCA'] || '')} ${(v['MODELO'] || '')}</td>
            <td class="px-3 py-2">${obtenerAccionRegistro(v)}</td>
            <td class="px-3 py-2">${obtenerNovedad(v)}</td>
        </tr>`;
    });
    html += `</tbody></table></div></div>
    </div>`;
    
    contenedor.innerHTML = html;
}

function renderizarDetalleOperativo(valSiipne, op) {
    const contenedor = document.getElementById('tabla-consulta-siipne');
    
    // Filtrar personas y vehículos que pertenecen a este operativo
    const personasOp = basePersonas.filter(p => String(p['ID OPERATIVO'] || p['FROT']).trim() === valSiipne);
    const vehiculosOp = baseVehiculos.filter(v => String(v['ID OPERATIVO'] || v['FROT']).trim() === valSiipne);
    
    const responsableInfo = extraerResponsable(op);
    const d = parseExcelDate(op['FECHA HORA'] || op['FECHA OPERATIVO'] || op['FECHA']);
    const fechaStr = (d ? d.toISOString().split('T')[0] : '-') + ' ' + obtenerHora(op, d);

    // 1. Cabecera del Operativo
    let html = `
    <div class="mb-6 bg-slate-900 border border-slate-700 p-4 rounded-lg">
        <h3 class="text-xl font-bold text-white mb-2">Operativo N° ${valSiipne}</h3>
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm mb-4">
            <div><span class="block text-slate-400 text-xs uppercase">Responsable</span><span class="text-white font-medium cursor-pointer hover:text-blue-400 hover:underline" onclick="document.getElementById('buscadorResponsable').value='${responsableInfo}'; ejecutarBusqueda('nombre');">${responsableInfo}</span></div>
            <div><span class="block text-slate-400 text-xs uppercase">Fecha y Hora</span><span class="text-white">${fechaStr}</span></div>
            <div><span class="block text-slate-400 text-xs uppercase">Tipo</span><span class="text-white">${op['TIPO DE OPERATIVO'] || '-'}</span></div>
            <div><span class="block text-slate-400 text-xs uppercase">Ubicación</span><span class="text-white">${op['DISTRITO'] || '-'} / ${op['CIRCUITO'] || '-'}</span></div>
        </div>
        <div class="flex gap-4">
            <div class="bg-slate-800 p-3 rounded flex-1 text-center border border-slate-700">
                <span class="block text-xl font-bold text-emerald-400">${personasOp.length}</span>
                <span class="text-xs text-slate-400 uppercase">Personas Registradas</span>
            </div>
            <div class="bg-slate-800 p-3 rounded flex-1 text-center border border-slate-700">
                <span class="block text-xl font-bold text-amber-400">${vehiculosOp.length}</span>
                <span class="text-xs text-slate-400 uppercase">Vehículos Registrados</span>
            </div>
        </div>
    </div>`;
    
    // 2. Tabla de Personas
    html += `
    <div class="mb-6">
        <h4 class="text-md font-bold text-emerald-400 mb-2 flex items-center gap-2">
            <span class="bg-emerald-500/20 px-2 py-1 rounded border border-emerald-500/30 text-[0.6rem]">▼</span> 
            Detalle de Personas Registradas
        </h4>
        <div class="overflow-x-auto bg-slate-900 rounded-md border border-slate-700 shadow-inner max-h-[300px] custom-scrollbar">
            <table class="w-full text-left text-xs text-slate-300 whitespace-nowrap relative min-w-max">
                <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800 border-b border-slate-700 sticky top-0 z-10">
                    <tr>
                        <th class="px-3 py-2 font-semibold">CÉDULA</th>
                        <th class="px-3 py-2 font-semibold">NOMBRES</th>
                        <th class="px-3 py-2 font-semibold">SEXO</th>
                        <th class="px-3 py-2 font-semibold">FECHA NAC.</th>
                        <th class="px-3 py-2 font-semibold">ACCIÓN DE REGISTRO</th>
                        <th class="px-3 py-2 font-semibold">NOVEDAD</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-700/30">`;
    if (personasOp.length === 0) {
        html += `<tr><td colspan="6" class="px-3 py-4 text-center text-slate-500 italic">No hay personas registradas en este operativo.</td></tr>`;
    } else {
        personasOp.forEach(p => {
            const fNac = formatearFechaEspecial(p['FECHA DE NACIMIENTO'] || p['FECHA NACIMIENTO'] || p['A DE NACIMIENTO'] || '-');
            html += `
            <tr class="hover:bg-slate-700/50 transition-colors">
                <td class="px-3 py-2 font-mono text-white">${p['CEDULA'] || p['IDENTIFICACION'] || '-'}</td>
                <td class="px-3 py-2">${p['NOMBRES'] || '-'}</td>
                <td class="px-3 py-2">${p['SEXO'] || '-'}</td>
                <td class="px-3 py-2">${fNac}</td>
                <td class="px-3 py-2">${obtenerAccionRegistro(p)}</td>
                <td class="px-3 py-2">${obtenerNovedad(p)}</td>
            </tr>`;
        });
    }
    html += `</tbody></table></div></div>`;

    // 3. Tabla de Vehículos 
    html += `
    <div class="mb-4">
        <h4 class="text-md font-bold text-amber-400 mb-2 flex items-center gap-2">
            <span class="bg-amber-500/20 px-2 py-1 rounded border border-amber-500/30 text-[0.6rem]">▼</span> 
            Detalle de Vehículos Registrados
        </h4>
        <div class="overflow-x-auto bg-slate-900 rounded-md border border-slate-700 shadow-inner max-h-[300px] custom-scrollbar">
            <table class="w-full text-left text-xs text-slate-300 whitespace-nowrap relative min-w-max">
                <thead class="text-[0.65rem] text-slate-400 uppercase bg-slate-800 border-b border-slate-700 sticky top-0 z-10">
                    <tr>
                        <th class="px-3 py-2 font-semibold">PLACA</th>
                        <th class="px-3 py-2 font-semibold">MARCA</th>
                        <th class="px-3 py-2 font-semibold">MODELO</th>
                        <th class="px-3 py-2 font-semibold">AÑO</th>
                        <th class="px-3 py-2 font-semibold">CLASE</th>
                        <th class="px-3 py-2 font-semibold">MOTOR</th>
                        <th class="px-3 py-2 font-semibold">CHASIS</th>
                        <th class="px-3 py-2 font-semibold">SERVICIO</th>
                        <th class="px-3 py-2 font-semibold">ACCIÓN DE REGISTRO</th>
                        <th class="px-3 py-2 font-semibold">NOVEDAD</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-700/30">`;
    if (vehiculosOp.length === 0) {
        html += `<tr><td colspan="10" class="px-3 py-4 text-center text-slate-500 italic">No hay vehículos registrados en este operativo.</td></tr>`;
    } else {
        vehiculosOp.forEach(v => {
            html += `
            <tr class="hover:bg-slate-700/50 transition-colors">
                <td class="px-3 py-2 font-mono text-white">${v['PLACA'] || '-'}</td>
                <td class="px-3 py-2">${v['MARCA'] || '-'}</td>
                <td class="px-3 py-2">${v['MODELO'] || '-'}</td>
                <td class="px-3 py-2">${v['AÑO'] || v['ANIO'] || v['AÑO FABRICACION'] || '-'}</td>
                <td class="px-3 py-2">${v['CLASE'] || '-'}</td>
                <td class="px-3 py-2">${v['MOTOR'] || '-'}</td>
                <td class="px-3 py-2">${v['CHASIS'] || '-'}</td>
                <td class="px-3 py-2">${v['SERVICIO'] || '-'}</td>
                <td class="px-3 py-2">${obtenerAccionRegistro(v)}</td>
                <td class="px-3 py-2">${obtenerNovedad(v)}</td>
            </tr>`;
        });
    }
    html += `</tbody></table></div></div>`;
    
    contenedor.innerHTML = html;
}