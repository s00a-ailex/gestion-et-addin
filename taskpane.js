// taskpane.js
// Lógica completa del panel de Gestión de Publicaciones.
// Maneja: inicialización, lectura de metadata, rendering por rol,
// todas las acciones de botón y comunicación con la Azure Functions API.

'use strict';

// ══════════════════════════════════════════════════════════════
// CONFIGURACIÓN
// ══════════════════════════════════════════════════════════════
const CONFIG = {
  API_BASE: 'https://atcazfunc.azurewebsites.net/api',
};

// ══════════════════════════════════════════════════════════════
// ESTADO DE LA APLICACIÓN
// ══════════════════════════════════════════════════════════════
let state = {
  spItemId:         null,   // SharePoint Item ID (fuente de verdad para identificar doc)
  itemId:           null,   // alias de spItemId para compatibilidad
  documentId:       null,   // atc_Documents.Id en SQL
  versionId:        null,   // atc_DocumentVersions.Id en SQL
  driveId:          null,   // SharePoint Drive ID (solo para publicar)
  driveItemId:      null,   // SharePoint DriveItem ID (solo para publicar)
  webUrl:           null,   // URL del documento en SharePoint
  nombreDocumento:  null,
  metadata:         null,
  emailUsuario:     null,
  userId:           null,
  rolActual:        null,
  rolResuelto:      false,
  identidades:      '',
  demoMode:         false,  // Flag para modo demo/test
};

// ══════════════════════════════════════════════════════════════
// DEFINICIÓN DE PASOS PARA LA BARRA DE PROGRESO
// ══════════════════════════════════════════════════════════════
const PASOS_FLUJO = [
  { key: 'Borrador_Usuario',     label: 'Borrador' },
  { key: 'Solicitud_Publicacion', label: 'Compliance' },  // bug1: nodo compliance se activa desde Solicitud
  { key: 'Revision_Aprobadores', label: 'Aprobación' },
  { key: 'Listo_Publicar',       label: 'Listo' },
  { key: 'Publicado',            label: 'Publicado' },
];

const ORDEN_PASOS = [
  'Borrador',
  'Borrador_Usuario',
  'Iteracion_Usuario',
  'Compliance_Observado',
  'Solicitud_Publicacion',    // desde aquí el nodo Compliance se marca activo
  'Revision_Compliance',
  'Compliance_Aprobado',
  'Revision_Aprobadores',
  'Revision_Aprobadores_NuevaRonda',
  'Aprobadores_Comentario',
  'Listo_Publicar',
  'Publicado',
  'Dado_De_Baja'

];

// ══════════════════════════════════════════════════════════════
// INICIALIZACIÓN DE OFFICE
// ══════════════════════════════════════════════════════════════
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) {
    mostrarToast('Este complemento solo funciona en Word.', 'error');
    return;
  }

  try {
    console.log('[Init] Office listo. Mostrando pantalla de login.');
    // Solo mostramos la pantalla de login
    // El usuario elige entre autenticarse o entrar en modo demo
    ocultarLoading();
  } catch (err) {
    console.error('[Init]', err);
    mostrarToast('Error al inicializar el complemento.', 'error');
  }
});

// ══════════════════════════════════════════════════════════════
// OBTENER CONTEXTO DEL DOCUMENTO
// Busca el List Item ID numérico (SharePointListItemId en SQL)
// que corresponde al ID de la lista de SharePoint (ej: 132)
// ══════════════════════════════════════════════════════════════
async function obtenerContextoDocumento() {
  return new Promise((resolve) => {
    Word.run(async (context) => {
      try {
        const docUrl = Office.context.document?.url || '';
        let nombreArchivo = '';
        let spItemId = null;

        if (docUrl) {
          const urlSinParams = docUrl.split('?')[0];
          nombreArchivo = decodeURIComponent(urlSinParams.split('/').pop() || '');
        }

        // Fallback desarrollo: query params
        const addinParams = new URLSearchParams(window.location.search);
        if (addinParams.get('spItemId') || addinParams.get('itemId')) {
          spItemId      = addinParams.get('spItemId') || addinParams.get('itemId');
          nombreArchivo = addinParams.get('nombre') || nombreArchivo || 'Documento (Dev)';
        }

        state.spItemId        = spItemId;
        state.itemId          = spItemId;
        state.nombreDocumento = nombreArchivo;
        document.getElementById('docNombre').textContent = nombreArchivo || 'Cargando...';

        console.log('[Context] Nombre:', nombreArchivo);

        // Si no tenemos spItemId desde query params,
        // buscarlo en SQL por nombre del archivo
        if (!spItemId && nombreArchivo && nombreArchivo.includes('.doc')) {
          try {
            const datos = await apiCall('GET',
              `sp/getMetadataByName?nombre=${encodeURIComponent(nombreArchivo)}`
            );
            if (datos?.id || datos?.spListItemId) {
              spItemId       = String(datos.spListItemId || datos.id);
              state.spItemId = spItemId;
              state.itemId   = spItemId;
              state.driveId     = datos.driveId || null;
              state.driveItemId = datos.driveItemId || null;
              console.log('[Context] spItemId resuelto por nombre:', spItemId);
            }
          } catch(e) {
            console.warn('[Context] No se pudo resolver por nombre:', e);
          }
        }

        console.log('[Context] spItemId final:', state.spItemId);
        resolve();
      } catch (e) {
        console.warn('[Context] Error:', e);
        resolve();
      }
    });
  });
}

// ══════════════════════════════════════════════════════════════
// OBTENER EMAIL DEL USUARIO
// Intenta obtenerlo desde Office, localStorage o query param
// ══════════════════════════════════════════════════════════════
async function obtenerEmailUsuario() {
  // 1. Intentar desde Office SSO
  try {
    const userProfile = Office.context.mailbox?.userProfile;
    if (userProfile?.emailAddress) {
      state.emailUsuario = userProfile.emailAddress;
      console.log('[Email] Desde Office mailbox:', state.emailUsuario);
      return;
    }
  } catch (e) { /* no disponible */ }

  // 2. Intentar desde roamingSettings de Office
  try {
    const saved = Office.context.roamingSettings?.get('userEmail');
    if (saved) {
      state.emailUsuario = saved;
      console.log('[Email] Desde roamingSettings:', state.emailUsuario);
      return;
    }
  } catch (e) { /* no disponible */ }

  // 3. Query param (cuando viene desde link del email)
  const urlParams = new URLSearchParams(window.location.search);
  const emailParam = urlParams.get('email');
  if (emailParam) {
    state.emailUsuario = emailParam;
    console.log('[Email] Desde query param:', state.emailUsuario);
    return;
  }

  // 4. sessionStorage (si ya lo ingresó antes en esta sesión)
  try {
    const cached = sessionStorage.getItem('gp_userEmail');
    if (cached) {
      state.emailUsuario = cached;
      console.log('[Email] Desde sessionStorage:', state.emailUsuario);
      return;
    }
  } catch (e) { /* no disponible */ }

  console.warn('[Email] No se pudo obtener email automáticamente');
}

// ══════════════════════════════════════════════════════════════
// RESOLVER ROL DEL USUARIO PARA ESTE DOCUMENTO
// Llama a /api/sp/resolverRol con email e itemId
// ══════════════════════════════════════════════════════════════
async function resolverRolUsuario() {
  if (!state.emailUsuario || !state.spItemId) return;

  try {
    mostrarLoading('Verificando permisos...');

    // Leer columnas de SharePoint como fallback
    // (la API las usará solo si SQL está vacío, y las sincronizará)
    let colCompliance  = null;
    let colDuenos      = null;
    let colAprobadores = null;

    try {
      const spMeta = await apiCall('GET',
        `sp/getMetadata?itemId=${state.itemId}`
      );
      colCompliance  = spMeta?.colCompliance  || spMeta?.Compliance  || null;
      colDuenos      = spMeta?.colDuenos      || spMeta?.Duenos      || null;
      colAprobadores = spMeta?.colAprobadores || spMeta?.Aprobadores || null;
      console.log('[Rol] Columnas SP:', { colCompliance, colDuenos, colAprobadores });
    } catch(e) {
      console.warn('[Rol] No se pudieron leer columnas de SP:', e.message);
    }

    const resultado = await apiCall('POST', 'doc/resolverRol', {
      email:          state.emailUsuario,
      spItemId:       state.spItemId,
      colCompliance,
      colDuenos,
      colAprobadores,
    });

    state.rolActual   = resultado.rolDocumento;
    state.userId      = resultado.userId;
    state.documentId  = resultado.documentId;
    state.rolResuelto = true;
    state.identidades = resultado.identificadoresUsuario || '';

    console.log(`[Rol] Resuelto: ${state.rolActual} | fuente: ${resultado.fuenteRoles} | identidades: ${resultado.identificadoresUsuario}`);

    try { sessionStorage.setItem('gp_userEmail', state.emailUsuario); } catch(e) {}

    const selectRol = document.getElementById('selectRol');
    if (selectRol && state.rolActual) selectRol.value = state.rolActual;

  } catch (err) {
    console.warn('[Rol] No se pudo resolver rol:', err.message);
    state.rolActual   = null;
    state.rolResuelto = true;
  } finally {
    ocultarLoading();
  }
}

// ══════════════════════════════════════════════════════════════
// MOSTRAR INPUT DE EMAIL (fallback cuando no se puede obtener)
// ══════════════════════════════════════════════════════════════
function mostrarInputEmail() {
  // Crear overlay de login simple
  const overlay = document.createElement('div');
  overlay.id = 'emailOverlay';
  overlay.style.cssText = `
    position:fixed; inset:0; background:rgba(255,255,255,0.97);
    display:flex; flex-direction:column; align-items:center;
    justify-content:center; z-index:200; padding:32px;
  `;
  overlay.innerHTML = `
    <div style="text-align:center; max-width:280px;">
      <div style="font-size:32px; margin-bottom:12px;">📋</div>
      <div style="font-size:15px; font-weight:600; color:#2f3437; margin-bottom:8px;">
        Gestión de Publicaciones
      </div>
      <div style="font-size:12px; color:#6b6f72; margin-bottom:24px; line-height:1.5;">
        Accede con tu cuenta corporativa para continuar
      </div>
      <div style="font-size:12px; color:#6b6f72; margin-bottom:16px;">
      Este complemento utiliza tu cuenta corporativa para identificar tu rol en el flujo de aprobación.
      </div>
      <input
        id="inputEmail"
        type="email"
        placeholder="usuario@empresa.com"
        style="width:100%; padding:10px 12px; border:1px solid #e0e0e0;
               border-radius:6px; font-size:13px; font-family:inherit;
               margin-bottom:12px; outline:none;"
      />
      <button
        onclick="confirmarEmail()"
        style="width:100%; padding:10px; background:#4f8ef7; color:#fff;
               border:none; border-radius:999px; font-size:13px;
               font-weight:600; cursor:pointer; font-family:inherit;"
      >
        Acceder →
      </button>
      <div id="emailError" style="color:#e53935; font-size:11px; margin-top:8px; display:none;">
        Por favor ingresa un email válido
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  // Focus en el input
  setTimeout(() => document.getElementById('inputEmail')?.focus(), 100);
}

async function confirmarEmail() {
  const input = document.getElementById('inputEmail');
  const email = input?.value?.trim();

  if (!email || !email.includes('@')) {
    document.getElementById('emailError').style.display = 'block';
    return;
  }

  state.emailUsuario = email;

  // Remover overlay
  document.getElementById('emailOverlay')?.remove();

  // Continuar con el flujo normal
  mostrarLoading('Cargando...');
  try {
    await resolverRolUsuario();
    await cargarMetadata();
  } finally {
    ocultarLoading();
  }
}

// ══════════════════════════════════════════════════════════════
// LLAMADAS A LA API (MODO HÍBRIDO: DEMO + PRODUCCIÓN)
// ══════════════════════════════════════════════════════════════
async function apiCall(method, endpoint, body = null) {
  if (state.demoMode) {
    console.log(`[Demo-Sim] ${method} a ${endpoint}`, body);
    
    // Simular latencia de red para una experiencia realista del revisor
    await new Promise(r => setTimeout(r, 500));

    // 1. Manejo de consultas GET (Metadata)
    // Siempre devolvemos el objeto en memoria para que la UI no reciba null
    if (method === 'GET' && (endpoint.includes('doc/metadata') || endpoint.includes('sp/getMetadata'))) {
      return state.metadata;
    }

    // ── 2. Si es cambio de estado, actualizamos la memoria local ──
    if (endpoint.includes('doc/cambiarEstado')) {
      state.metadata.pasoActual = body.nuevoEstado;
      state.metadata.currentStateChangedAt = new Date().toISOString();
      
      // Forzamos que la observación viaje a la metadata en memoria para que renderizarUI la lea
      state.metadata.observacion = body.observacion || ""; 
      
      console.log(`[Demo-Sim] Estado actualizado internamente a: ${body.nuevoEstado}. Observación: ${body.observacion}`);

      // Forzamos el renderizado inmediato con los datos frescos en memoria
      setTimeout(() => renderizarUI(), 100);
      return state.metadata;
    }

    // 3. Gestión de Aprobaciones e Incrementos
    if (endpoint.includes('doc/incrementarAprobacion')) {
      const userAprob = state.metadata.aprobaciones.find(a => a.aprobadorEmail === state.emailUsuario);
      
      if (body.conComentarios || body.observaciones) {
        // FLUJO DE CORRECCIÓN: El documento vuelve a una etapa de revisión
        state.metadata.pasoActual = 'Aprobadores_Comentario'; 
        if (userAprob) {
          userAprob.status = 'Observado';
          userAprob.comentario = body.observaciones || "Favor corregir secciones indicadas.";
        }
        mostrarToast('Comentarios registrados. El emisor debe atenderlos.', 'warning');
      } else {
        // APROBACIÓN NORMAL: Solo cuentan los perfiles de aprobador
        if (userAprob) {
          userAprob.status = 'Aprobado';
          // Solo incrementamos si no había aprobado antes
          state.metadata.aprobadoresConformes = state.metadata.aprobaciones.filter(a => a.status === 'Aprobado' || a.status === 'Conforme').length;
        }
        
        // Si se alcanza el total de aprobadores (excluyendo cumplimiento)
        if (state.metadata.aprobadoresConformes >= state.metadata.aprobadoresTotal) {
          state.metadata.pasoActual = 'Listo_Publicar';
        }
      }
      setTimeout(() => renderizarUI(), 100);
      return state.metadata;
    }

    // 4. FLUJO DE ATENCIÓN DE COMENTARIOS (Owner -> Aprobador)
    if (endpoint.includes('comentarios/atender')) {
      const aprob = state.metadata.aprobaciones.find(a => a.approvalId === body.approvalId);
      if (aprob) {
        aprob.status = 'Atendido'; // Estado intermedio: Esperando confirmación
        console.log(`[Demo] Comentarios de ${aprob.aprobadorNombre} marcados como atendidos.`);
        
        // Si ya no quedan más con "Observado", podemos devolver el paso a revisión activa
        const todaviaObservados = state.metadata.aprobaciones.filter(a => a.status === 'Observado');
        if (todaviaObservados.length === 0) {
          state.metadata.pasoActual = 'Revision_Aprobadores';
        }
      }
      setTimeout(() => renderizarUI(), 100);
      return state.metadata;
    }

    // 5. Conformidad (Aprobador acepta que se atendió su duda)
    if (endpoint.includes('comentarios/conformidad')) {
      const aprob = state.metadata.aprobaciones.find(a => a.approvalId === body.approvalId);
      if (aprob) {
        aprob.status = body.conforme ? 'Conforme' : 'Reabierto';
        if (body.conforme) state.metadata.aprobadoresConformes++;
      }
      setTimeout(() => renderizarUI(), 100);
      return state.metadata;
    }

    return state.metadata;
  }

  // ══════════════════════════════════════════════════════════════
  // CÓDIGO ORIGINAL PARA PRODUCCIÓN (FETCH REAL)
  // ══════════════════════════════════════════════════════════════
  const url = `${CONFIG.API_BASE}/${endpoint}`;
  const response = await fetch(url, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : null
  });
  
  const data = await response.json();
  if (!data.success) {
    throw new Error(data.error || 'Error en la API');
  }
  return data.data;
}

// ══════════════════════════════════════════════════════════════
// Botón para volver a la pantalla de inicio
// ══════════════════════════════════════════════════════════════

function volverAlInicio() {
  // 1. Resetear el estado lógico
  state.demoMode = false;
  state.metadata = null;
  state.rolActual = null;
  state.emailUsuario = null;

  // 2. Limpiar elementos visuales de la pantalla de Login
  const loginLoading = document.getElementById('loginLoading');
  const loginInfo = document.getElementById('loginInfo');
  const loginError = document.getElementById('loginError');
  const loginForm = document.getElementById('loginFormElement');

  if (loginLoading) loginLoading.classList.remove('show');
  if (loginInfo) {
    loginInfo.classList.remove('show');
    loginInfo.textContent = '';
  }
  if (loginError) {
    loginError.classList.remove('show');
    loginError.textContent = '';
  }
  
  // 3. Asegurar que el formulario de correo sea visible
  if (loginForm) loginForm.style.display = 'flex';

  // 4. Alternar contenedores principales
  ocultarAppContainer();
  mostrarLoginScreen();

  const emailInput = document.getElementById('emailInput');
  if (emailInput) emailInput.value = '';
  console.log('[App] Regreso a pantalla inicial exitoso y UI reseteada.');
}

// ══════════════════════════════════════════════════════════════
// CARGAR METADATA DESDE SQL (nueva fuente de verdad)
// ══════════════════════════════════════════════════════════════
async function cargarMetadata() {
  mostrarLoading('Cargando estado del documento...');
  try {
    const spItemId = state.spItemId || state.itemId;
    if (!spItemId) throw new Error('No se pudo identificar el documento');

    const metadata = await apiCall('GET', `doc/metadata?spItemId=${encodeURIComponent(spItemId)}`);
    state.metadata = metadata;

    // Guardar referencias para uso posterior
    if (metadata.sharePointDriveId) state.driveId     = metadata.sharePointDriveId;
    if (metadata.sharePointWebUrl)  state.webUrl      = metadata.sharePointWebUrl;
    if (metadata.documentId)        state.documentId  = metadata.documentId;
    if (metadata.versionId)         state.versionId   = metadata.versionId;

    renderizarUI();
    iniciarPolling();  // iniciar detección de cambios en segundo plano
  } catch (err) {
    console.error('[Metadata]', err);
    mostrarToast('Error al cargar el estado del documento: ' + err.message, 'error');
  } finally {
    ocultarLoading();
  }
}

// ══════════════════════════════════════════════════════════════
// POLLING — Detectar cambios de estado en segundo plano
// Consulta cada 30s sin recargar la UI
// Si detecta cambio → muestra banner de notificación
// ══════════════════════════════════════════════════════════════
let _pollingInterval = null;
let _estadoPendiente = null;  // guarda el nuevo estado detectado

function iniciarPolling() {
  if (_pollingInterval) return;
  _pollingInterval = setInterval(verificarCambioEstado, 30000); // 30 segundos
  console.log('[Polling] Iniciado cada 30s');
}

function detenerPolling() {
  if (_pollingInterval) {
    clearInterval(_pollingInterval);
    _pollingInterval = null;
  }
}

async function verificarCambioEstado() {
  const spItemId = state.spItemId || state.itemId;
  if (!spItemId || !state.metadata) return;

  try {
    const nueva = await apiCall('GET', `doc/metadata?spItemId=${encodeURIComponent(spItemId)}`);
    const actual = state.metadata;

    const cambioEstado      = nueva.pasoActual !== actual.pasoActual;
    const cambioAprobaciones = nueva.aprobadoresConformes !== actual.aprobadoresConformes ||
                               nueva.aprobadoresTotal     !== actual.aprobadoresTotal;

    // Detectar si cambió el status de algún aprobador individual
    const aprobActual = (actual.aprobaciones || []);
    const aprobNueva  = (nueva.aprobaciones  || []);
    const cambioIndividual = aprobNueva.some(an => {
      const anterior = aprobActual.find(aa => aa.approvalId === an.approvalId);
      return anterior && anterior.status !== an.status;
    }) || aprobNueva.length !== aprobActual.length;

    if (cambioEstado || cambioAprobaciones || cambioIndividual) {
      console.log(`[Polling] Cambio detectado: estado=${cambioEstado} aprobaciones=${cambioAprobaciones} individual=${cambioIndividual}`);
      _estadoPendiente = nueva;

      // Mensaje según qué cambió
      let mensaje = '';
      if (cambioEstado) {
        const { texto } = mapearEstadoDisplay(nueva.pasoActual);
        mensaje = 'Estado actualizado: ' + texto;
      } else if (cambioIndividual) {
        mensaje = 'Un aprobador actualizó su revisión';
      } else {
        mensaje = 'Las aprobaciones se han actualizado';
      }
      mostrarBannerCambio(mensaje);
    }
  } catch (e) {
    console.warn('[Polling] Error al verificar:', e.message);
  }
}

function mostrarBannerCambio(mensaje) {
  const banner = document.getElementById('cambioBanner');
  const texto  = document.getElementById('cambioBannerTexto');
  if (!banner || !texto) return;
  texto.textContent = mensaje;
  banner.classList.add('show');
}

function aplicarCambioEstado() {
  const banner = document.getElementById('cambioBanner');
  if (banner) banner.classList.remove('show');

  if (_estadoPendiente) {
    const nueva      = _estadoPendiente;
    _estadoPendiente = null;
    state.metadata   = nueva;
    if (nueva.sharePointDriveId) state.driveId    = nueva.sharePointDriveId;
    if (nueva.documentId)        state.documentId = nueva.documentId;
    if (nueva.versionId)         state.versionId  = nueva.versionId;
    renderizarUI();
  } else {
    cargarMetadata();
  }
}


function renderizarUI() {
  const m = state.metadata;
  if (!m) return;

  const paso = m.pasoActual;

  // ── Selector de rol: visible para admin/compliance ──
  const devSelector = document.getElementById('devRolSelector');
  if (devSelector) {
    // Si es demo, forzamos que SIEMPRE se vea
    if (state.demoMode) {
      devSelector.style.display = 'block';
    } else {
      // Lógica original para producción
      const esAdmin = state.identidades && (state.identidades.includes('Admin') || state.identidades.includes('ComplianceOfficer'));
      devSelector.style.display = (!state.rolResuelto || state.rolActual === 'todos' || esAdmin) ? 'block' : 'none';
    }
  }

  // Actualizar textos básicos
  document.getElementById('docNombre').textContent = m.title || 'Manual de Procedimientos v1';
  

  // ── Info del documento ──
  document.getElementById('infoResponsable').textContent    = m.title || m.code || '—';
  document.getElementById('infoTipoAprobacion').textContent = m.aprobadoresTotal
    ? `${m.aprobadoresConformes}/${m.aprobadoresTotal} aprobadores`
    : '—';
  document.getElementById('infoUltimaAccion').textContent   = m.currentStateChangedAt
    ? new Date(m.currentStateChangedAt).toLocaleDateString('es-CL', { day:'2-digit', month:'short', hour:'2-digit', minute:'2-digit' })
    : '—';

  // ── Estado badge ──
  const { texto, clase } = mapearEstadoDisplay(paso);
  actualizarEstadoUI(texto, clase, paso);

  // ── Barra de progreso ──
  renderizarProgreso(paso);

  // ── Banner publicado ──
// ── Banner Final: Publicado o Dado de Baja ──
  const banner = document.getElementById('publicadoBanner');
  if (paso === 'Publicado' || paso === 'Dado_De_Baja') {
    const iconEl  = document.getElementById('publicadoIcon');
    const titleEl = document.getElementById('publicadoTitle');
    const subEl   = document.getElementById('publicadoSub');

     // El estado explícito determina si es baja (robusto, no depende de texto)
    const esBaja = (paso === 'Dado_De_Baja')

    if (esBaja) {
      // 1. Componentes visuales del panel lateral (Baja)
      if (iconEl)  iconEl.textContent  = '🗑️';
      if (titleEl) titleEl.textContent = 'Documento Dado de Baja';
      if (subEl)   subEl.textContent   = 'Este documento ha sido retirado del sistema definitivamente por el equipo de Compliance.';
      if (banner) {
        banner.style.background = '#fef2f2'; 
        banner.style.borderColor = '#fecaca';
        titleEl.style.color = '#7f1d1d';
        subEl.style.color = '#5c1200';
      }
    } else {
      // 2. Componentes visuales del panel lateral (Publicación)
      if (iconEl)  iconEl.textContent  = '✅';
      if (titleEl) titleEl.textContent = 'Documento Publicado';
      if (subEl)   subEl.textContent   = 'Este documento ha sido publicado exitosamente y movido a la carpeta Maestros.';
      if (banner) {
        banner.style.background = 'var(--success-bg)'; 
        banner.style.borderColor = '#9fd89f';
        titleEl.style.color = 'var(--success)';
        subEl.style.color = '#054005';
      }

      // ⚡ CONTROL DE RENDERIZADO REACTIVO ÚNICO
      if (!window.__caratulaInyectada) {
        window.__caratulaInyectada = true; // El cerrojo se activa de inmediato en este hilo síncrono
        escribirPublicacionEnDocumento().catch(err => {
          window.__caratulaInyectada = false; // Si falla, liberamos el cerrojo
          console.error("[Word API Error]", err);
        });
      }
    }
    
    if (banner) banner.classList.add('show');
  } else {
    if (banner) banner.classList.remove('show');
    window.__caratulaInyectada = false; // Resetea el flag si salen del modo publicado
  }

  // ── Secciones por rol ──
  renderizarSeccionUsuario(paso);
  renderizarSeccionCompliance(paso);
  renderizarSeccionAprobadores(paso, m);
    if (state.demoMode && devSelector) {
        devSelector.style.display = 'block';
        // Aseguramos que el z-index sea alto para que no lo tape ningún modal de comentarios
        devSelector.style.zIndex = '9999';
    }
}

// ── Mapear paso a texto legible y color ──
function mapearEstadoDisplay(paso) {
  const mapa = {
    'Borrador_Usuario':                  { texto: 'Borrador del Usuario',         clase: 'azul'  },
    'Solicitud_Publicacion':             { texto: 'Solicitud Enviada',             clase: 'azul'  },
    'Solicitud_Prorroga':             { texto: 'Solicitud de prórroga',             clase: 'azul'  },
    'Revision_Compliance':               { texto: 'En Revisión - Compliance',      clase: 'azul'  },
    'Compliance_Aprobado':               { texto: 'Aprobado por Compliance',       clase: 'verde' },
    'Compliance_Observado':              { texto: 'Observado por Compliance',      clase: 'rojo'  },
    'Revision_Aprobadores':              { texto: 'En Revisión - Aprobadores',     clase: 'azul'  },
    'Revision_Aprobadores_NuevaRonda':   { texto: 'Nueva Ronda de Aprobación',     clase: 'azul'  },
    'Aprobadores_Comentario':            { texto: 'Comentarios de Aprobadores',    clase: 'rojo'  },
    'Iteracion_Usuario':                 { texto: 'Devuelto por Compliance',        clase: 'rojo'  },
    'Listo_Publicar':                    { texto: '✅ Listo para Publicar',         clase: 'verde' },
    'Publicado':                         { texto: '🎉 Publicado',                  clase: 'verde' },
    'Dado_De_Baja':                      { texto: '🗑️ Dado de Baja',              clase: 'rojo'  },
  };
  return mapa[paso] || { texto: paso, clase: 'gris' };
}

// ══════════════════════════════════════════════════════════════
// BARRA DE PROGRESO
// ══════════════════════════════════════════════════════════════


function renderizarProgreso(pasoActual) {
  const track = document.getElementById('stepsTrack');

  const ESTADOS_TERMINALES = ['Publicado', 'Dado_De_Baja'];
  const esTerminal = ESTADOS_TERMINALES.includes(pasoActual);
  const esBaja = pasoActual === 'Dado_De_Baja';

  const indiceActual = ORDEN_PASOS.indexOf(pasoActual);
  const hitosIndices = PASOS_FLUJO.map(h => ORDEN_PASOS.indexOf(h.key));

  track.innerHTML = PASOS_FLUJO.map((paso, i) => {
    const indiceHito = hitosIndices[i];
    const esUltimoHito = (i === PASOS_FLUJO.length - 1);
    let clase = '';

    if (esTerminal) {
      clase = 'done';
    } else if (indiceActual > indiceHito) {
      clase = 'done';
    } else if (indiceActual === indiceHito) {
      clase = 'active';
    }

    // En baja, el último círculo muestra el contexto de baja en vez de "Publicado"
    let label = paso.label;
    if (esBaja && esUltimoHito) {
      label = 'Baja';
    }

    return `
      <div class="step-item ${clase}">
        <div class="step-circle">${clase === 'done' ? '✓' : i + 1}</div>
        <div class="step-name">${label}</div>
      </div>
    `;
  }).join('');
}

// ══════════════════════════════════════════════════════════════
// RENDERIZAR SECCIONES POR ROL
// ══════════════════════════════════════════════════════════════

// ── DUEÑO / EMISOR ──
function renderizarSeccionUsuario(paso) {
  const seccion      = document.getElementById('seccionUsuario');
  const btnSolicitar = document.getElementById('btnSolicitarPublicacion');

  const pasosUsuario = [
    'Borrador_Usuario', 'Borrador',
    'Iteracion_Usuario', 'Compliance_Observado',
    'Aprobadores_Comentario',
    'Revision_Aprobadores', 'Revision_Aprobadores_NuevaRonda',
  ];

  const esRolConAcceso = ['dueno', 'usuario', 'compliance', 'todos'].includes(state.rolActual);
  if (!esRolConAcceso) { seccion.classList.add('hidden'); return; }

  seccion.classList.toggle('hidden', !pasosUsuario.includes(paso));

  // En Aprobadores_Comentario el emisor NO puede solicitar aún, solo atender comentarios
  const puedeEnviar = pasosUsuario.includes(paso) && paso !== 'Aprobadores_Comentario'
    && paso !== 'Revision_Aprobadores' && paso !== 'Revision_Aprobadores_NuevaRonda';
  btnSolicitar.disabled = !puedeEnviar;

  if (paso === 'Iteracion_Usuario' || paso === 'Compliance_Observado') {
    btnSolicitar.innerHTML = '<span>Reenviar para Publicación</span>';
  } else {
    btnSolicitar.innerHTML = '<span>Solicitar Publicación</span>';
  }

  // Mostrar lista de aprobadores para el emisor cuando está en fase de revisión
  const pasosConLista = [
    'Revision_Aprobadores', 'Revision_Aprobadores_NuevaRonda', 'Aprobadores_Comentario'
  ];
  if (pasosConLista.includes(paso)) {
    renderizarListaAprobadores('listaComentariosEmisor', 'seccionUsuario', state.metadata);
  } else {
    const c = document.getElementById('listaComentariosEmisor');
    if (c) c.innerHTML = '';
  }
}

// ── COMPLIANCE ──
function renderizarSeccionCompliance(paso) {
  const seccion     = document.getElementById('seccionCompliance');
  const btnFormalizar = document.getElementById('btnFormalizarRevision');
  const btnRechazar   = document.getElementById('btnRechazarCompliance');
  const btnPublicar   = document.getElementById('btnPublicar');
  const btnDarBaja    = document.getElementById('btnDarBaja');

  if (state.rolActual !== 'compliance' && state.rolActual !== 'todos') {
    seccion.classList.add('hidden'); return;
  }

  // Bug 2 y 6: compliance siempre visible en cualquier estado activo
  // (puede publicar saltándose aprobadores, dar de baja en cualquier momento, etc.)
  const pasosInactivos = ['Borrador_Usuario', 'Borrador', 'Publicado', 'Dado_De_Baja'];
  const visible = !pasosInactivos.includes(paso);
  seccion.classList.toggle('hidden', !visible);

  // Formalizar y Rechazar: solo cuando hay algo que revisar
  const puedeRevisar = ['Solicitud_Publicacion', 'Revision_Compliance'].includes(paso);
  btnFormalizar.disabled = !puedeRevisar;
  btnRechazar.disabled   = !puedeRevisar;

  // Bug 2: Publicar SIEMPRE habilitado para compliance (puede saltarse aprobadores)
  btnPublicar.disabled = false;
  btnDarBaja.disabled  = false;

  // Estilo especial cuando está Listo para Publicar (bug 5 y 6)
  if (paso === 'Listo_Publicar') {
    btnPublicar.classList.remove('btn-primary');
    btnPublicar.classList.add('btn-success');
  } else {
    btnPublicar.classList.remove('btn-success');
    btnPublicar.classList.add('btn-primary');
  }
  // ── Botón "Eliminar carátula" (solo demo + compliance/todos) ──
  const btnEliminarCaratula = document.getElementById('btnEliminarCaratula');
  if (btnEliminarCaratula) {
    const puedeEliminar = state.demoMode ||
      state.rolActual === 'compliance' ||
      state.rolActual === 'todos';
    btnEliminarCaratula.classList.toggle('hidden', !puedeEliminar);
  }

}

// ── APROBADORES ──
function renderizarSeccionAprobadores(paso, metadata) {
  const seccion = document.getElementById('seccionAprobadores');
  const btnSin  = document.getElementById('btnRevisadoSinComentarios');
  const btnCon  = document.getElementById('btnRevisadoConComentarios');

  const esAprobador  = state.rolActual === 'aprobador';
  const esCompliance = state.rolActual === 'compliance';
  const esTodos      = state.rolActual === 'todos';

  if (!esAprobador && !esCompliance && !esTodos) {
    seccion.classList.add('hidden'); return;
  }

  const pasosAprobador = ['Revision_Aprobadores', 'Revision_Aprobadores_NuevaRonda', 'Aprobadores_Comentario'];
  const visible = pasosAprobador.includes(paso);
  seccion.classList.toggle('hidden', !visible);

  const actual = metadata.aprobadoresConformes || 0;
  const total  = metadata.aprobadoresTotal     || 0;
  const pct    = total > 0 ? Math.min(actual / total * 100, 100) : 0;

  document.getElementById('aprobacionActual').textContent  = actual;
  document.getElementById('aprobacionTotal').textContent   = total;
  document.getElementById('aprobacionBarFill').style.width = pct + '%';

  // Botones activos solo en fases de revisión
  // Bug 4: si el aprobador ya aprobó pero el estado sigue en Revision_Aprobadores,
  // puede rectificar (cambiar a con comentarios)
  const miAprobacion = (metadata.aprobaciones || []).find(
    a => a.aprobadorEmail === state.emailUsuario
  );
  const yaAprobeSinComentarios = miAprobacion && miAprobacion.status === 'Aprobado';
  const puedeActuar = visible && paso !== 'Aprobadores_Comentario';

  btnSin.disabled = !puedeActuar;
  // Puede enviar con comentarios incluso si ya aprobó (rectificación), mientras siga en revisión
  btnCon.disabled = !puedeActuar;

  // Lista de aprobadores con estado para aprobadores/compliance
  renderizarListaAprobadores('listaComentarios', 'seccionAprobadores', metadata);
}

// ══════════════════════════════════════════════════════════════
// ACCIONES DE BOTONES
// ══════════════════════════════════════════════════════════════

// ── Helper: cambiar estado en SQL ──
async function cambiarEstado(nuevoEstado, observacion = null, metadata = null) {
  return apiCall('POST', 'doc/cambiarEstado', {
    spItemId:    state.spItemId || state.itemId,
    nuevoEstado,
    email:       state.emailUsuario,
    observacion,
    metadata,
  });
}

// ── Solicitar Publicación ──
async function accion_solicitarPublicacion() {
  const ok = await confirmar('El documento será enviado a revisión de Compliance.', '¿Solicitar Publicación?', '#4f8ef7');
  if (!ok) return;
  mostrarLoading('Enviando solicitud...');
  try {
    await cambiarEstado('Solicitud_Publicacion');
    mostrarToast('Solicitud enviada. Compliance ha sido notificado.', 'success');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Formalizar Revisión ──
async function accion_formalizarRevision() {
  const ok = await confirmar('Se enviará a Aprobadores. Se inicia el plazo de 3 días hábiles.', '¿Formalizar Revisión?', '#3a9c2a');
  if (!ok) return;
  mostrarLoading('Formalizando revisión...');
  try {
    await cambiarEstado('Revision_Aprobadores', 'Compliance aprobó y envía a aprobadores');
    mostrarToast('Revisión formalizada. Aprobadores notificados.', 'success');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Rechazar con Comentarios ──
async function accion_rechazarCompliance() {
  const ok = await confirmar('El documento será devuelto al responsable con sus comentarios.', '¿Rechazar y devolver?', '#e53935');
  if (!ok) return;
  mostrarLoading('Rechazando documento...');
  try {
    await cambiarEstado('Iteracion_Usuario', 'Compliance rechazó con comentarios');
    mostrarToast('Documento devuelto al responsable.', 'info');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Publicar Documento ──
async function accion_publicar() {
  const ok = await confirmar('Se creará la versión final, se generará el PDF y se moverá a Maestros.', '¿Publicar Documento?', '#3a9c2a');
  if (!ok) return;
  
  mostrarLoading('Publicando documento...');
  try {
    // 1. Cambia el estado a Publicado de forma normal
    await cambiarEstado('Publicado', 'Documento publicado por Compliance');
    
    // 2. Quitamos la ejecución directa de la carátula aquí para evitar el doble hilo
    mostrarToast('🎉 Documento publicado exitosamente.', 'success');
    await cargarMetadata();
  } catch (err) { 
    mostrarToast('Error al publicar: ' + err.message, 'error'); 
  } finally { 
    ocultarLoading(); 
  }
}

// ── Revisado Sin Comentarios ──
async function accion_revisadoSinComentarios() {
  const ok = await confirmar('Se registrará su aprobación sin comentarios.', '¿Confirmar aprobación?', '#3a9c2a');
  if (!ok) return;
  mostrarLoading('Registrando aprobación...');
  try {
    const resultado = await apiCall('POST', 'doc/incrementarAprobacion', {
      spItemId:       state.spItemId || state.itemId,
      email:          state.emailUsuario,
      conComentarios: false,
    });
    if (resultado.listoParaPublicar) {
      mostrarToast('✅ Umbral de aprobación alcanzado. Listo para Publicar.', 'success');
    } else {
      mostrarToast(`Aprobación registrada (${resultado.nuevoContador}/${resultado.total}).`, 'success');
    }
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Revisado Con Comentarios ──
async function accion_revisadoConComentarios() {
  const ok = await confirmar(
    'Se registrará que tienes comentarios. El emisor verá tu estado en la lista y podrá atenderlos.',
    '¿Registrar comentarios?', '#e53935'
  );
  if (!ok) return;
  mostrarLoading('Registrando comentarios...');
  try {
    // Registrar aprobación con comentarios en SQL
    await apiCall('POST', 'doc/incrementarAprobacion', {
      spItemId:       state.spItemId || state.itemId,
      email:          state.emailUsuario,
      conComentarios: true,
    });
    // Cambiar estado a Aprobadores_Comentario — el doc permanece aquí
    // hasta que el emisor atienda todos y los aprobadores confirmen conformidad
    await cambiarEstado('Aprobadores_Comentario', 'Aprobador registró comentarios');
    mostrarToast('Comentarios registrados. El emisor verá tu estado en la lista.', 'info');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ══════════════════════════════════════════════════════════════
// SIMULACIÓN DE ROL (solo para desarrollo)
// ══════════════════════════════════════════════════════════════
function cambiarRolSimulado(rol) {
  state.rolActual = rol;
  
  // Asignamos emails específicos para que la lógica de renderizado de botones 
  // (esRolConAcceso, etc.) se active correctamente
  const perfilesDemo = {
    'dueno':      'dueno@demo.com',
    'compliance': 'compliance@demo.com',
    'aprobador':  'aprobador@demo.com',
    'todos':      'admin@demo.com'
  };
  
  state.emailUsuario = perfilesDemo[rol] || 'dueno@demo.com';
  
  console.log(`[Demo] Cambiando a vista de: ${rol} (${state.emailUsuario})`);
  
  // Refrescamos la UI con la nueva perspectiva
  renderizarUI();
  mostrarToast(`Perspectiva de ${rol.toUpperCase()} activada`, 'info');
}

// ══════════════════════════════════════════════════════════════
// HELPERS DE UI
// ══════════════════════════════════════════════════════════════

function actualizarEstadoUI(texto, clase, pasoRaw) {
  document.getElementById('estadoTexto').textContent = texto;
  const dot = document.getElementById('estadoDot');
  dot.className = 'estado-dot ' + clase;
}

function mostrarLoading(texto = 'Procesando...') {
  document.getElementById('loadingText').textContent = texto;
  document.getElementById('loadingOverlay').classList.add('show');
}

function ocultarLoading() {
  document.getElementById('loadingOverlay').classList.remove('show');
}

let toastTimeout = null;
function mostrarToast(mensaje, tipo = 'info') {
  const toast  = document.getElementById('toast');
  const icon   = document.getElementById('toastIcon');
  const text   = document.getElementById('toastText');

  const iconos = { success: '✅', error: '❌', info: 'ℹ️' };
  icon.textContent = iconos[tipo] || 'ℹ️';
  text.textContent = mensaje;

  toast.className = `toast ${tipo} show`;

  if (toastTimeout) clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => {
    toast.classList.remove('show');
  }, tipo === 'error' ? 6000 : 4000);
}

// ══════════════════════════════════════════════════════════════
// DIÁLOGO DE CONFIRMACIÓN (reemplaza window.confirm)
// ══════════════════════════════════════════════════════════════
let _confirmarResolve = null;

function confirmar(mensaje, titulo = '¿Confirmar acción?', colorBtn = '#4f8ef7') {
  return new Promise((resolve) => {
    _confirmarResolve = resolve;
    document.getElementById('confirmTitulo').textContent  = titulo;
    document.getElementById('confirmMensaje').textContent = mensaje;
    document.getElementById('confirmBtn').style.background = colorBtn;
    document.getElementById('confirmOverlay').style.display = 'flex';
  });
}

function confirmarRespuesta(valor) {
  document.getElementById('confirmOverlay').style.display = 'none';
  if (_confirmarResolve) { _confirmarResolve(valor); _confirmarResolve = null; }
}

function cerrarModal(id) {
  document.getElementById(id).style.display = 'none';
}

// ── Solicitar Prórroga ──
function accion_solicitarProroga() {
  const hoy = new Date().toISOString().split('T')[0];
  document.getElementById('prorogaFecha').min = hoy;
  document.getElementById('prorogaFecha').value = '';
  document.getElementById('prorogaMotivo').value = '';
  document.getElementById('prorogaOverlay').style.display = 'flex';
}

async function confirmarProroga() {
  const fecha  = document.getElementById('prorogaFecha').value;
  const motivo = document.getElementById('prorogaMotivo').value.trim();
  if (!fecha) { mostrarToast('Debes indicar una fecha comprometida.', 'error'); return; }
  cerrarModal('prorogaOverlay');
  mostrarLoading('Enviando solicitud de prórroga...');
  try {
    await cambiarEstado('Solicitud_Publicacion', motivo || 'Solicitud de prórroga', { tipo: 'proroga', fechaCompromiso: fecha });
    mostrarToast('Solicitud de prórroga enviada a Compliance.', 'success');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Ratificar Vigencia ──
async function accion_ratificarVigencia() {
  const ok = await confirmar('Se notificará a Compliance que el documento no tiene cambios.', '¿Ratificar Vigencia?', '#3a9c2a');
  if (!ok) return;
  mostrarLoading('Ratificando vigencia...');
  try {
    await cambiarEstado('Solicitud_Publicacion', 'Ratificación de vigencia sin cambios', { tipo: 'ratificacion' });
    mostrarToast('Vigencia ratificada. Compliance ha sido notificado.', 'success');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Solicitar Baja ──
function accion_solicitarBaja() {
  document.getElementById('bajaMotivo').value = '';
  document.getElementById('bajaError').style.display = 'none';
  document.getElementById('bajaOverlay').style.display = 'flex';
}

async function confirmarBaja() {
  const motivo = document.getElementById('bajaMotivo').value.trim();
  if (!motivo) { document.getElementById('bajaError').style.display = 'block'; return; }
  cerrarModal('bajaOverlay');
  mostrarLoading('Enviando solicitud de baja...');
  try {
    await cambiarEstado('Solicitud_Publicacion', motivo, { tipo: 'baja', motivo });
    mostrarToast('Solicitud de baja enviada a Compliance.', 'info');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Dar de Baja (Compliance) ──
// ── Dar de Baja (Compliance) ──
async function accion_darBaja() {
  const ok = await confirmar('El documento será dado de baja definitivamente.', '¿Dar de Baja?', '#e53935');
  if (!ok) return;
  
  mostrarLoading('Procesando baja...');
  try {
    // 1. Registra el cambio de estado (guardará la palabra "baja" en la observación)
    await cambiarEstado('Dado_De_Baja', 'Documento dado de baja por Compliance', { tipo: 'baja_compliance' });
    
    // 2. INTERACCIÓN CON EL DOCUMENTO (Inyecta el banner rojo al inicio)
    await escribirBajaEnDocumento();
    
    mostrarToast('Documento dado de baja e inyectado en Word.', 'info');
    await cargarMetadata();
  } catch (err) { 
    mostrarToast('Error: ' + err.message, 'error'); 
  } finally { 
    ocultarLoading(); 
  }
}

// ── No Aplica (Aprobador) ──
async function accion_noAplica() {
  const ok = await confirmar('Indicarás que este documento no es de tu competencia.', '¿Marcar como No Aplica?', '#4f8ef7');
  if (!ok) return;
  mostrarLoading('Registrando...');
  try {
    await apiCall('POST', 'doc/incrementarAprobacion', {
      spItemId:       state.spItemId || state.itemId,
      email:          state.emailUsuario,
      conComentarios: false,
    });
    mostrarToast('Registrado como No Aplica.', 'success');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Atender Comentario (Emisor marca que atendió) ──
async function accion_atenderComentario(approvalId, aprobadorEmail) {
  const ok = await confirmar(
    `¿Marcar como atendidos los comentarios de ${aprobadorEmail}?`,
    '¿Atender comentario?', '#4f8ef7'
  );
  if (!ok) return;
  mostrarLoading('Registrando atención...');
  try {
    await apiCall('POST', 'comentarios/atender', {
      approvalId,
      emisorEmail:  state.emailUsuario,
      emisorNombre: state.emailUsuario,
    });
    mostrarToast('Comentario marcado como atendido.', 'success');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ── Confirmar Conformidad (Aprobador confirma si fue resuelto) ──
async function accion_confirmarConformidad(approvalId, conforme) {
  const titulo  = conforme ? '¿Confirmar conformidad?' : '¿Reabrir comentarios?';
  const mensaje = conforme
    ? 'Confirmarás que tus comentarios fueron resueltos correctamente.'
    : 'Indicarás que tus comentarios aún no han sido resueltos.';
  const ok = await confirmar(mensaje, titulo, conforme ? '#3a9c2a' : '#e53935');
  if (!ok) return;
  mostrarLoading('Registrando...');
  try {
    await apiCall('POST', 'comentarios/conformidad', {
      approvalId,
      aprobadorEmail: state.emailUsuario,
      conforme,
    });
    mostrarToast(conforme ? '✅ Conformidad registrada.' : '↩ Comentario reabierto.', conforme ? 'success' : 'info');
    await cargarMetadata();
  } catch (err) { mostrarToast('Error: ' + err.message, 'error'); }
  finally { ocultarLoading(); }
}

// ══════════════════════════════════════════════════════════════
// SIMULACIÓN DE ROL (solo para desarrollo)
// ══════════════════════════════════════════════════════════════
function cambiarRolSimulado(rol) {
  state.rolActual = rol;
  if (state.metadata) renderizarUI();
}

// ══════════════════════════════════════════════════════════════
// LISTA UNIFICADA DE APROBADORES — Estilo C (tabla con punto)
// Permisos: dueno→atender, aprobador→confirmar, compliance→ambos
// ══════════════════════════════════════════════════════════════
function renderizarListaAprobadores(contenedorId, seccionId, metadata) {
  if (!metadata) return;
  let contenedor = document.getElementById(contenedorId);
  if (!contenedor) {
    const body = document.querySelector('#' + seccionId + ' .rol-body');
    if (!body) return;
    contenedor = document.createElement('div');
    contenedor.id = contenedorId;
    body.appendChild(contenedor);
  }

  const aprobaciones = metadata.aprobaciones || [];
  if (!aprobaciones.length) {
    contenedor.innerHTML = '<div style="font-size:12px;color:#9ca3af;padding:8px 0;text-align:center;">No hay aprobadores asignados aún.</div>';
    return;
  }

  const rol = state.rolActual;
  const esDueno     = rol === 'dueno'     || rol === 'compliance' || rol === 'todos';
  const esAprobador = rol === 'aprobador' || rol === 'compliance' || rol === 'todos';

  // Config de estado: [color punto, texto, color texto]
  const cfg = function(status) {
    return ({
      'Aprobado':  ['#166534', 'Aprobado',                '#166534'],
      'Conforme':  ['#166534', 'Conforme',                '#166534'],
      'NoAplica':  ['#9ca3af', 'Sin observaciones',       '#9ca3af'],
      'Observado': ['#7f1d1d', 'Con comentarios',         '#7f1d1d'],
      'Reabierto': ['#7f1d1d', 'Reabierto',               '#7f1d1d'],
      'Atendido':  ['#92400e', 'Esperando confirmaci\u00f3n', '#92400e'],
      'Pendiente': ['#d1d5db', 'Pendiente',               '#9ca3af'],
    }[status] || ['#d1d5db', status, '#9ca3af']);
  };

  // Encabezado tabla
  var html = '<div style="margin-top:8px;border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;">';
  html += '<div style="display:grid;grid-template-columns:1fr auto;padding:6px 12px;background:#f9fafb;border-bottom:1px solid #e5e7eb;">';
  html += '<span style="font-size:10px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:0.06em;">Revisor</span>';
  html += '<span style="font-size:10px;font-weight:600;color:#6b7280;text-transform:uppercase;letter-spacing:0.06em;">Estado</span>';
  html += '</div>';

  aprobaciones.forEach(function(a, i) {
    var c = cfg(a.status);
    var dotColor   = c[0];
    var statusText = c[1];
    var textColor  = c[2];
    var puedeAtender   = esDueno     && (a.status === 'Observado' || a.status === 'Reabierto');
    var esSuComentario = a.aprobadorEmail === state.emailUsuario || rol === 'compliance' || rol === 'todos';
    var puedeConfirmar = esAprobador && esSuComentario && a.status === 'Atendido';
    var isLast = i === aprobaciones.length - 1;
    var rowBorder = isLast && !puedeAtender && !puedeConfirmar ? '' : 'border-bottom:1px solid #f3f4f6;';
    var rowBg = (a.status === 'Observado' || a.status === 'Reabierto') ? 'background:#fffbfb;' :
                 a.status === 'Atendido' ? 'background:#fffdf5;' : '';

    html += '<div style="' + rowBorder + rowBg + 'padding:9px 12px;">';
    html += '<div style="display:grid;grid-template-columns:1fr auto;align-items:center;">';
    html += '<div>';
    html += '<div style="font-size:12px;color:#111827;font-weight:500;">' + (a.aprobadorNombre || a.aprobadorEmail) + '</div>';
    html += '<div style="font-size:11px;color:#9ca3af;margin-top:1px;">' + a.aprobadorEmail + '</div>';
    html += '</div>';
    html += '<div style="display:flex;align-items:center;gap:5px;">';
    html += '<div style="width:7px;height:7px;border-radius:50%;background:' + dotColor + ';flex-shrink:0;"></div>';
    html += '<span style="font-size:11px;color:' + textColor + ';font-weight:500;white-space:nowrap;">' + statusText + '</span>';
    html += '</div></div>';

    if (puedeAtender) {
      html += '<button onclick="accion_atenderComentario(' + a.approvalId + ', \'' + a.aprobadorEmail + '\')" ';
      html += 'style="margin-top:8px;width:100%;padding:7px;font-size:11px;font-weight:500;border:none;';
      html += 'background:#1e3a5f;color:#fff;border-radius:5px;cursor:pointer;font-family:inherit;">';
      html += 'Marcar comentarios como atendidos</button>';
    }

    if (puedeConfirmar) {
      html += '<div style="display:flex;gap:6px;margin-top:8px;">';
      html += '<button onclick="accion_confirmarConformidad(' + a.approvalId + ', true)" ';
      html += 'style="flex:1;padding:7px;font-size:11px;font-weight:500;border:none;';
      html += 'background:#166534;color:#fff;border-radius:5px;cursor:pointer;font-family:inherit;">Conforme</button>';
      html += '<button onclick="accion_confirmarConformidad(' + a.approvalId + ', false)" ';
      html += 'style="flex:1;padding:7px;font-size:11px;font-weight:500;border:none;';
      html += 'background:#7f1d1d;color:#fff;border-radius:5px;cursor:pointer;font-family:inherit;">A\u00fan tengo observaciones</button>';
      html += '</div>';
    }

    html += '</div>';
  });

  html += '</div>';
  html += '<div style="margin-top:6px;font-size:10px;color:#9ca3af;text-align:center;">Los comentarios est\u00e1n en el archivo con control de cambios.</div>';
  contenedor.innerHTML = html;
}// ══════════════════════════════════════════════════════════════
// FUNCIONES DE LOGIN Y MODO DEMO
// Agregar estas funciones al final de taskpane.js
// ══════════════════════════════════════════════════════════════

/**
 * Autentica al usuario con su email
 */
async function autenticarUsuario(event) {
  event.preventDefault();
  
  const email = document.getElementById('emailInput').value.trim();
  if (!email) {
    mostrarErrorLogin('Por favor ingresa tu email corporativo');
    return false;
  }

  const errorEl = document.getElementById('loginError');
  const infoEl = document.getElementById('loginInfo');
  const formEl = document.getElementById('loginFormElement');
  const loadingEl = document.getElementById('loginLoading');
  
  errorEl.classList.remove('show');
  infoEl.classList.remove('show');
  formEl.style.display = 'none';
  loadingEl.classList.add('show');

  try {
    // Simulamos verificación de email (en producción, validarías con tu backend)
    await new Promise(r => setTimeout(r, 1500));

    // Validar email corporativo (ejemplo)
    if (!email.includes('@')) {
      throw new Error('Email inválido');
    }

    state.emailUsuario = email;
    state.demoMode = false;
    
    // Guardar en sessionStorage
    try { sessionStorage.setItem('gp_userEmail', email); } catch(e) {}
    
    console.log('[Auth] Usuario autenticado:', email);
    
    // Iniciar la aplicación
    await inicializarAplicacion();
    
    return false;
  } catch (err) {
    console.error('[Auth] Error:', err);
    mostrarErrorLogin(err.message || 'Error al autenticar. Intenta de nuevo.');
    formEl.style.display = 'flex';
    loadingEl.classList.remove('show');
  }
  
  return false;
}

/**
 * Inicia el modo demo/test sin autenticación
 */
async function iniciarModoDemo() {
  const errorEl = document.getElementById('loginError');
  const infoEl = document.getElementById('loginInfo');
  const formEl = document.getElementById('loginFormElement');
  const loadingEl = document.getElementById('loginLoading');
  
  errorEl.classList.remove('show');
  infoEl.classList.remove('show');
  formEl.style.display = 'none';
  loadingEl.classList.add('show');
  
  // Mostrar mensaje informativo
  infoEl.textContent = '⚡ Modo Demo activado - Puedes explorar toda la funcionalidad';
  infoEl.classList.add('show');

  try {
    // Pequeña pausa para UX
    await new Promise(r => setTimeout(r, 1200));

    // Configurar estado de demo
    state.emailUsuario = 'demo@test.local';
    state.demoMode = true;
    state.userId = 'DEMO-USER-001';
    state.documentId = 'DEMO-DOC-001';
    state.versionId = 'DEMO-VER-001';
    state.spItemId = '999';
    state.itemId = '999';
    state.nombreDocumento = 'Documento Demo - Modo Test';
    
    console.log('[Demo] Modo demo activado');
    
    // Iniciar la aplicación con datos demo
    await inicializarAplicacionDemo();
    
  } catch (err) {
    console.error('[Demo] Error:', err);
    mostrarErrorLogin('Error al iniciar modo demo');
    formEl.style.display = 'flex';
    loadingEl.classList.remove('show');
  }
}

/**
 * Inicializa la aplicación después de autenticación normal
 */
async function inicializarAplicacion() {
  try {
    mostrarLoading('Cargando aplicación...');

    // 1. Obtener contexto del documento
    await obtenerContextoDocumento();

    if (state.itemId && state.emailUsuario) {
      // 2. Resolver rol del usuario para este documento
      await resolverRolUsuario();
      // 3. Cargar metadata completa
      await cargarMetadata();
    } else if (!state.emailUsuario) {
      mostrarToast('Error: Email no disponible', 'error');
    } else {
      mostrarToast('No se pudo identificar el documento. Ábrelo desde SharePoint.', 'error');
      actualizarEstadoUI('—', 'gris', 'No identificado');
    }
    
    ocultarLoading();
    mostrarAppContainer();
    ocultarLoginScreen();

  } catch (err) {
    console.error('[InitApp]', err);
    mostrarToast('Error al inicializar la aplicación', 'error');
    ocultarLoading();
  }
}

/**
 * Inicializa la aplicación en modo demo con un escenario pre-cargado completo.
 */
async function inicializarAplicacionDemo() {
  try {
    mostrarLoading('Cargando modo demo...');
    state.demoMode = true;
    state.emailUsuario = 'dueno@demo.com';
    state.rolActual = 'dueno'; // Empezamos como Dueño
    state.rolResuelto = true;
    state.spItemId = '999';

    state.metadata = {
      title: 'Manual de Procedimientos v1',
      code: 'PRO-2026-001',
      pasoActual: 'Borrador_Usuario', // <--- PUNTO DE INICIO
      currentStateChangedAt: new Date().toISOString(),
      aprobadoresConformes: 0,
      aprobadoresTotal: 1, // Solo los que están en la lista de abajo con perfil aprobador
      aprobaciones: [
        { approvalId: 101, aprobadorNombre: 'Alexis Aprobador', aprobadorEmail: 'aprobador@demo.com', status: 'Observado', rol: 'aprobador' },
      ]
    };

    document.getElementById('docNombre').textContent = state.metadata.title;
    ocultarLoginScreen();
    mostrarAppContainer();
    renderizarUI();
    
    mostrarToast('Iniciado en modo Borrador. Use el selector inferior para cambiar de rol.', 'info');
  } catch (err) {
    console.error(err);
  } finally {
    ocultarLoading();
  }
}

// ══════════════════════════════════════════════════════════════
// CARÁTULA DE CONTROL DOCUMENTAL — v7
//
// Arquitectura por capas:
//
//   ┌─────────────────────────────────────┐
//   │ obtenerDatosCaratula()              │  ← decide fuente
//   │   ├─ mapearDesdeDemo(state)         │     (demo o backend)
//   │   └─ obtenerDesdeBackend(spItemId)  │
//   └────────────────┬────────────────────┘
//                    ▼
//   ┌─────────────────────────────────────┐
//   │ escribirPublicacionEnDocumento()    │  ← wrapper con lock
//   └────────────────┬────────────────────┘
//                    ▼
//   ┌─────────────────────────────────────┐
//   │ _ejecutarCaratula(datos)            │
//   │   ├─ _crearCaratula(ctx, datos)     │  ← primera vez
//   │   └─ _actualizarCaratula(ctx,datos) │  ← re-publicación
//   └─────────────────────────────────────┘
//
// La función de pintado NO conoce el origen de los datos.
// ══════════════════════════════════════════════════════════════
 
const TAG_RAIZ = "gp.caratula";
 
// Identificadores textuales de las tablas dinámicas (primera celda)
const HEADER_ELAB = "Elaborado por:";
const HEADER_REV = "Revisado por:";
const HEADER_HIST = "Aprobado por:";
 
let _caratulaPromise = null;
 
 

// ══════════════════════════════════════════════════════════════
// CAPA 1: OBTENCIÓN DE DATOS — abstrae demo vs online
// ══════════════════════════════════════════════════════════════
 
/**
 * Devuelve los datos de la carátula desde la fuente correcta.
 * El resto del código no sabe si vienen de demo o backend.
 */
async function obtenerDatosCaratula() {
  if (state.demoMode) {
    console.log('[Caratula] Datos desde DEMO');
    return mapearDesdeDemo(state);
  } else {
    console.log('[Caratula] Datos desde BACKEND');
    return await obtenerDesdeBackend(state.spItemId);
  }
}
 
/**
 * Convierte el string CSV ("juan@bf.cl, maria@bf.cl, Riesgo")
 * en un array limpio de strings.
 * - Trimea espacios
 * - Filtra vacíos
 */
function parsearCSV(valor) {
  if (!valor || typeof valor !== 'string') return [];
  return valor.split(',').map(s => s.trim()).filter(Boolean);
}
 
/**
 * Toma una lista de nombres crudos y los convierte en filas
 * con shape {nombre, gerencia, area, cargo}.
 *
 * NOTA: por ahora solo poblamos 'nombre'. Los otros campos
 * quedan vacíos hasta que tengamos el resolver de email→persona.
 */
function expandirParticipantes(listaNombres) {
  return listaNombres.map(nombre => ({
    nombre: nombre,
    gerencia: "",
    area: "",
    cargo: ""
  }));
}
 
/**
 * Mapper de datos cuando estamos en modo demo.
 * Lee desde state.metadata y state.colDuenos / state.colAprobadores
 * (o los nombres equivalentes que use tu state).
 */
function mapearDesdeDemo(state) {
  const md = state.metadata || {};
 
  // ─── DATOS DEMO de participantes ───
  // Si state.metadata.colDuenos / colAprobadores vienen poblados, se usan.
  // Si no, se usan estos defaults para que la carátula no quede vacía.
  const colDuenos = md.colDuenos || state.colDuenos
    || "Dueño Test 1, Dueño Test 2";
  const colAprobadores = md.colAprobadores || state.colAprobadores
    || "Aprobador Test 1, Aprobador Test 2, Aprobador Test 3";
 
  return {
    titulo:   md.title    || "Manual de Procedimientos v1",
    codigo:   md.code     || "PRO-2026-001",
    version:  md.version  || "1.0",
    gerencia: md.gerencia || "Operaciones / Compliance",
    fecha:    new Date().toLocaleDateString('es-CL'),
    aprobador: md.aprobador || "Compliance",
    detalleVersion: md.detalleVersion || "Versión inicial",
    duenos: expandirParticipantes(parsearCSV(colDuenos)),
    aprobadores: expandirParticipantes(parsearCSV(colAprobadores))
  };
}


/**
 * Mapper de datos cuando estamos online contra el backend.
 * Llama a los endpoints reales y normaliza al mismo shape que demo.
 */
async function obtenerDesdeBackend(spItemId) {
  const docInfo = await apiCall(`doc/metadata?spItemId=${spItemId}`);
  const colDuenos = docInfo.colDuenos || "";
  const colAprobadores = docInfo.colAprobadores || "";
 
  return {
    titulo:   docInfo.titulo  || docInfo.title  || "Procedimiento",
    codigo:   docInfo.codigo  || docInfo.code   || "PR-GEN-001",
    version:  docInfo.version || "1.0",
    gerencia: docInfo.gerencia || "",
    fecha:    new Date().toLocaleDateString('es-CL'),
    aprobador: docInfo.aprobador || "Compliance",
    detalleVersion: docInfo.detalleVersion || "Actualización de versión",
    duenos: expandirParticipantes(parsearCSV(colDuenos)),
    aprobadores: expandirParticipantes(parsearCSV(colAprobadores))
  };
}


/**
 * Muestra el contenedor de la aplicación
 */
function mostrarAppContainer() {
  const container = document.getElementById('appContainer');
  if (container) {
    container.style.display = 'flex';
    container.style.flexDirection = 'column';
  }
}

/**
 * Oculta el contenedor de la aplicación
 */
function ocultarAppContainer() {
  const container = document.getElementById('appContainer');
  if (container) container.style.display = 'none';
}

/**
 * Muestra la pantalla de login
 */
function mostrarLoginScreen() {
  const loginScreen = document.getElementById('loginScreen');
  if (loginScreen) loginScreen.classList.add('show');
}

/**
 * Oculta la pantalla de login
 */
function ocultarLoginScreen() {
  const loginScreen = document.getElementById('loginScreen');
  if (loginScreen) loginScreen.classList.remove('show');
}

/**
 * Muestra error en la pantalla de login
 */
function mostrarErrorLogin(mensaje) {
  const errorEl = document.getElementById('loginError');
  if (errorEl) {
    errorEl.textContent = '❌ ' + mensaje;
    errorEl.classList.add('show');
  }
}

/**
 * Oculta el loading overlay
 */
function ocultarLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.classList.remove('show');
}

/**
 * Cambio de rol simulado (para desarrollo/demo)
 */
function cambiarRolSimulado(rol) {
  state.rolActual = rol;
  
  // Mapeamos emails ficticios para que la lógica de "miAprobacion" funcione
  const emails = {
    'dueno': 'dueno@demo.com',
    'compliance': 'compliance@demo.com',
    'aprobador': 'aprobador@demo.com',
    'todos': 'admin@demo.com'
  };
  
  state.emailUsuario = emails[rol] || state.emailUsuario;
  
  console.log('[Demo] Perspectiva cambiada a:', rol, 'Usuario:', state.emailUsuario);
  mostrarToast(`Perspectiva: ${rol.toUpperCase()}`, 'success');
  renderizarUI();
}

// ══════════════════════════════════════════════════════════════
// FUNCIONES AUXILIARES PARA DEMO
// ══════════════════════════════════════════════════════════════

/**
 * Simula acciones de botones en modo demo
 */
function simularAccionDemo(nombreAccion) {
  if (!state.demoMode) return;
  
  const mensajes = {
    'solicitar': '✅ Solicitud de publicación enviada (Demo)',
    'revisar': '✅ Revisión realizada (Demo)',
    'aprobar': '✅ Documento aprobado (Demo)',
    'rechazar': '⚠️ Documento rechazado (Demo)',
    'publicar': '✅ Documento publicado (Demo)',
    'comentar': '✅ Comentarios agregados (Demo)',
    'proroga': '✅ Prórroga solicitada (Demo)',
    'baja': '✅ Baja de documento solicitada (Demo)'
  };
  
  mostrarToast(mensajes[nombreAccion] || '✅ Acción completada (Demo)', 'success');
}

/**
 * Renderiza un flujo completo interactivo para demostración Microsoft
 */
function renderizarFlujoCompletoDEMO() {
  const mainContent = document.querySelector('.main-content');
  if (!mainContent) return;

  let html = '<div style="margin-top: 20px; border-top: 1px solid #e5e7eb; padding-top: 20px;">';
  html += '<h3 style="margin-top: 0; font-size: 14px; font-weight: 700; color: #1f2937; text-transform: uppercase; letter-spacing: 0.05em;">📊 Flujo Completo de Aprobación (Ejemplo)</h3>';
  html += '<p style="font-size: 12px; color: #6b7280; margin: 8px 0 16px 0;">Haz clic en cada estado para ver cómo se vería en diferentes fases del proceso</p>';
  
  // Estados del flujo con ejemplos
  const estados = [
    {
      nombre: 'Borrador',
      descripcion: 'Documento en elaboración por el Dueño',
      color: '#9ca3af',
      completado: true,
      icono: '📝',
      ejemplo: 'borrador'
    },
    {
      nombre: 'Solicitud Publicación',
      descripcion: 'El Dueño envía a revisión de Compliance',
      color: '#3b82f6',
      completado: true,
      icono: '📤',
      ejemplo: 'solicitud'
    },
    {
      nombre: 'En Compliance',
      descripcion: 'El equipo legal revisa el documento',
      color: '#f59e0b',
      completado: false,
      icono: '🔍',
      ejemplo: 'compliance'
    },
    {
      nombre: 'Aprobación',
      descripcion: 'Los ejecutivos aprueban el documento',
      color: '#ec4899',
      completado: false,
      icono: '✏️',
      ejemplo: 'aprobacion'
    },
    {
      nombre: 'Listo',
      descripcion: 'Todas las aprobaciones completadas',
      color: '#8b5cf6',
      completado: false,
      icono: '✅',
      ejemplo: 'listo'
    },
    {
      nombre: 'Publicado',
      descripcion: 'Disponible en el portal corporativo',
      color: '#10b981',
      completado: false,
      icono: '🌐',
      ejemplo: 'publicado'
    }
  ];

  // Renderizar estados como botones interactivos
  html += '<div style="display: flex; flex-direction: column; gap: 10px; margin-top: 12px;">';
  
  estados.forEach((estado, idx) => {
    const isActive = estado.completado;
    const bgColor = isActive ? estado.color : '#f3f4f6';
    const textColor = isActive ? '#fff' : '#6b7280';
    const borderColor = estado.color;
    
    html += '<button onclick="mostrarEjemploEstadoDEMO(\'' + estado.ejemplo + '\')" style="';
    html += 'display: flex; align-items: center; gap: 12px;';
    html += 'padding: 12px; border-radius: 8px;';
    html += 'background: ' + bgColor + '; border-left: 4px solid ' + borderColor + ';';
    html += 'color: ' + textColor + '; font-size: 13px;';
    html += 'border: none; cursor: pointer; font-family: inherit; text-align: left;';
    html += 'transition: all 0.2s ease;';
    html += '"';
    html += 'onmouseover="this.style.transform=\'translateX(4px)\'; this.style.boxShadow=\'0 4px 12px rgba(0,0,0,0.1)\'"';
    html += 'onmouseout="this.style.transform=\'translateX(0)\'; this.style.boxShadow=\'none\'"';
    html += '>';
    
    html += '<span style="font-size: 18px;">' + estado.icono + '</span>';
    html += '<div style="flex: 1;">';
    html += '<div style="font-weight: 600;">' + estado.nombre + '</div>';
    html += '<div style="font-size: 12px; opacity: 0.8;">' + estado.descripcion + '</div>';
    html += '</div>';
    
    if (isActive) {
      html += '<span style="font-weight: 700; font-size: 12px;">✓</span>';
    } else {
      html += '<span style="font-size: 12px; opacity: 0.5;">→</span>';
    }
    
    html += '</button>';
  });
  
  html += '</div>';
  html += '</div>';

  mainContent.insertAdjacentHTML('beforeend', html);
}

/**
 * Muestra un ejemplo de cómo se vería en cada estado
 */
function mostrarEjemploEstadoDEMO(tipoEstado) {
  if (!state.demoMode) return;
  
  const ejemplos = {
    borrador: {
      estado: 'Borrador',
      paso: '1/6',
      titulo: '📝 Documento en Elaboración',
      descripcion: 'El Dueño está creando/editando el documento. Solo él puede hacer cambios.',
      rolVisible: 'dueno',
      botonesRol: [
        { texto: '📤 Solicitar Publicación', color: '#3b82f6', desc: 'Enviar a Compliance para revisión' },
        { texto: '🗑️ Descartar Cambios', color: '#ef4444', desc: 'Cancelar cambios sin guardar' }
      ],
      info: 'En este estado, solo el Dueño puede ver y editar el documento. El documento no ha sido enviado a revisión aún.'
    },
    solicitud: {
      estado: 'Solicitud de Publicación',
      paso: '2/6',
      titulo: '📤 Solicitado para Publicación',
      descripcion: 'El Dueño ha enviado el documento. Está esperando revisión de Compliance.',
      rolVisible: 'dueno',
      botonesRol: [
        { texto: '📅 Solicitar Prórroga', color: '#f59e0b', desc: 'Extender plazo de revisión' },
        { texto: '🗑️ Solicitar Baja', color: '#ef4444', desc: 'Eliminar el documento' }
      ],
      info: 'El documento está en cola. El Dueño puede solicitar prórroga si necesita más tiempo, o solicitar su baja.'
    },
    compliance: {
      estado: 'En Compliance',
      paso: '3/6',
      titulo: '🔍 En Revisión Legal',
      descripcion: 'El equipo de Compliance está revisando el documento para validar cumplimiento normativo.',
      rolVisible: 'compliance',
      botonesRol: [
        { texto: '✅ Formalizar Revisión', color: '#10b981', desc: 'Aprueba la revisión de Compliance' },
        { texto: '❌ Rechazar con Comentarios', color: '#ef4444', desc: 'Rechaza y devuelve para cambios' },
        { texto: '📝 Agregar Comentarios', color: '#8b5cf6', desc: 'Adiciona observaciones al documento' }
      ],
      info: 'En este estado, el equipo de Compliance revisa si el documento cumple con regulaciones. Puede aprobar, rechazar o solicitar cambios.'
    },
    aprobacion: {
      estado: 'Aprobación',
      paso: '4/6',
      titulo: '✏️ En Proceso de Aprobación',
      descripcion: 'Los ejecutivos aprueban el documento después de revisar los cambios de Compliance.',
      rolVisible: 'aprobador',
      botonesRol: [
        { texto: '✅ Aprobado', color: '#10b981', desc: 'Aprueba el documento' },
        { texto: '✏️ Con Comentarios', color: '#8b5cf6', desc: 'Aprueba pero agrega observaciones' },
        { texto: '⊘ Sin Observaciones', color: '#6b7280', desc: 'Aprueba sin comentarios' }
      ],
      info: 'Barra de aprobación: 0/2 aprobadores. Cada aprobador puede revisar y aprobar independientemente.'
    },
    listo: {
      estado: 'Listo',
      paso: '5/6',
      titulo: '✅ Completamente Aprobado',
      descripcion: 'Todas las aprobaciones están completas. El documento está listo para publicación.',
      rolVisible: 'compliance',
      botonesRol: [
        { texto: '📢 Publicar Documento', color: '#10b981', desc: 'Publica en el portal' },
        { texto: '🗑️ Dar de Baja', color: '#ef4444', desc: 'Elimina antes de publicar' }
      ],
      info: 'El documento ha pasado todas las etapas de revisión y aprobación. Solo falta publicarlo oficialmente.'
    },
    publicado: {
      estado: 'Publicado',
      paso: '6/6',
      titulo: '🌐 Documento Publicado',
      descripcion: 'El documento está disponible en el portal corporativo para todos los usuarios autorizados.',
      rolVisible: 'compliance',
      botonesRol: [
        { texto: '📋 Ver en Portal', color: '#3b82f6', desc: 'Abre el documento en el portal' },
        { texto: '🗑️ Dar de Baja', color: '#ef4444', desc: 'Archiva el documento' }
      ],
      info: 'El flujo de aprobación está completado. El documento ahora es oficial y vinculante.'
    }
  };
  
  const ejemplo = ejemplos[tipoEstado];
  if (!ejemplo) return;
  
  // Crear modal con ejemplo
  let html = '<div style="';
  html += 'position: fixed; inset: 0; background: rgba(0,0,0,0.6);';
  html += 'display: flex; align-items: center; justify-content: center;';
  html += 'z-index: 2000; padding: 20px;';
  html += '" onclick="this.remove()">';
  
  html += '<div style="';
  html += 'background: #fff; border-radius: 12px; max-width: 600px;';
  html += 'width: 100%; max-height: 85vh; overflow-y: auto; padding: 24px;';
  html += 'box-shadow: 0 20px 60px rgba(0,0,0,0.3);';
  html += '" onclick="event.stopPropagation()">';
  
  // Header
  html += '<div style="display: flex; align-items: start; justify-content: space-between; margin-bottom: 16px;">';
  html += '<div>';
  html += '<div style="font-size: 24px; margin-bottom: 8px;">✨ Paso ' + ejemplo.paso + '</div>';
  html += '<h2 style="margin: 0; font-size: 18px; font-weight: 700;">' + ejemplo.titulo + '</h2>';
  html += '</div>';
  html += '<button onclick="this.closest(\'div\').parentElement.remove()" style="background: none; border: none; font-size: 20px; cursor: pointer;">✕</button>';
  html += '</div>';
  
  html += '<p style="margin: 0 0 16px 0; font-size: 13px; color: #6b7280; line-height: 1.5;">' + ejemplo.descripcion + '</p>';
  
  // Información del estado
  html += '<div style="padding: 12px; background: #f0fdf4; border-radius: 6px; border-left: 3px solid #10b981; margin-bottom: 16px;">';
  html += '<div style="font-size: 12px; color: #166534;">' + ejemplo.info + '</div>';
  html += '</div>';
  
  // Rol visible
  html += '<div style="padding: 12px; background: #f3f4f6; border-radius: 6px; margin-bottom: 16px;">';
  html += '<div style="font-size: 11px; font-weight: 600; color: #6b7280; text-transform: uppercase; margin-bottom: 8px;">Rol que interactúa</div>';
  const rolLabels = { 'dueno': '📤 Dueño / Emisor', 'compliance': '🔍 Compliance Officer', 'aprobador': '✏️ Aprobador' };
  html += '<div style="font-size: 13px; font-weight: 600; color: #1f2937;">' + (rolLabels[ejemplo.rolVisible] || ejemplo.rolVisible) + '</div>';
  html += '</div>';
  
  // Acciones disponibles
  html += '<div style="margin-bottom: 16px;">';
  html += '<div style="font-size: 11px; font-weight: 600; color: #6b7280; text-transform: uppercase; margin-bottom: 10px;">Acciones disponibles en este estado</div>';
  html += '<div style="display: flex; flex-direction: column; gap: 8px;">';
  
  ejemplo.botonesRol.forEach(btn => {
    html += '<button onclick="mostrarToast(\'' + btn.texto + ' ✅\', \'success\'); this.closest(\'div\').parentElement.parentElement.remove();" style="';
    html += 'text-align: left; padding: 10px 12px; background: ' + btn.color + '20; border: 1px solid ' + btn.color + '40;';
    html += 'border-radius: 6px; cursor: pointer; font-family: inherit; transition: all 0.2s;';
    html += '" onmouseover="this.style.background=\'' + btn.color + '30\'" onmouseout="this.style.background=\'' + btn.color + '20\'">';
    html += '<div style="font-weight: 600; font-size: 12px; color: #1f2937; margin-bottom: 4px;">' + btn.texto + '</div>';
    html += '<div style="font-size: 11px; color: #6b7280;">' + btn.desc + '</div>';
    html += '</button>';
  });
  
  html += '</div>';
  html += '</div>';
  
  // Botones de cierre
  html += '<div style="display: flex; gap: 10px; padding-top: 16px; border-top: 1px solid #e5e7eb;">';
  html += '<button onclick="this.closest(\'div\').parentElement.remove()" style="flex: 1; padding: 10px; background: #f3f4f6; border: none; border-radius: 6px; cursor: pointer; font-family: inherit; font-weight: 500;">Cerrar</button>';
  html += '</div>';
  
  html += '</div>';
  html += '</div>';
  
  document.body.insertAdjacentHTML('beforeend', html);
}

// Agregar al export
window.mostrarEjemploEstadoDEMO = mostrarEjemploEstadoDEMO;



// Exportar funciones para que sean accesibles globalmente
window.autenticarUsuario = autenticarUsuario;
window.iniciarModoDemo = iniciarModoDemo;
window.cambiarRolSimulado = cambiarRolSimulado;
window.simularAccionDemo = simularAccionDemo;
window.renderizarFlujoCompletoDEMO = renderizarFlujoCompletoDEMO;

/**
 * Renderiza las acciones del Dueño al inicio
 */
function renderizarAccionesDuenoDEMO() {
  const mainContent = document.querySelector('.main-content');
  if (!mainContent) return;

  let html = '<div style="margin-top: 16px; padding: 16px; background: #f0f9ff; border-radius: 8px; border-left: 4px solid #3b82f6;">';
  html += '<h3 style="margin: 0 0 12px 0; font-size: 13px; font-weight: 700; color: #1e40af;">📤 Acciones Disponibles para el Dueño</h3>';
  
  html += '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">';
  
  // Solicitar Publicación
  html += '<div style="padding: 12px; background: #fff; border-radius: 6px; border: 1px solid #bfdbfe;">';
  html += '<div style="font-weight: 600; font-size: 12px; color: #1e40af; margin-bottom: 8px;">📤 Solicitar Publicación</div>';
  html += '<div style="font-size: 11px; color: #6b7280; margin-bottom: 10px;">Envía el documento a revisión de Compliance</div>';
  html += '<button onclick="mostrarEjemploAccionDEMO(\'solicitar\')" style="width: 100%; padding: 8px; background: #3b82f6; color: #fff; border: none; border-radius: 4px; font-size: 11px; cursor: pointer; font-family: inherit; font-weight: 500;">Ver Ejemplo</button>';
  html += '</div>';
  
  // Solicitar Prórroga
  html += '<div style="padding: 12px; background: #fff; border-radius: 6px; border: 1px solid #fed7aa;">';
  html += '<div style="font-weight: 600; font-size: 12px; color: #b45309; margin-bottom: 8px;">📅 Solicitar Prórroga</div>';
  html += '<div style="font-size: 11px; color: #6b7280; margin-bottom: 10px;">Extiende el plazo de revisión</div>';
  html += '<button onclick="mostrarEjemploAccionDEMO(\'proroga\')" style="width: 100%; padding: 8px; background: #f59e0b; color: #fff; border: none; border-radius: 4px; font-size: 11px; cursor: pointer; font-family: inherit; font-weight: 500;">Ver Ejemplo</button>';
  html += '</div>';
  
  // Solicitar Baja
  html += '<div style="padding: 12px; background: #fff; border-radius: 6px; border: 1px solid #fecaca; grid-column: 1 / -1;">';
  html += '<div style="font-weight: 600; font-size: 12px; color: #dc2626; margin-bottom: 8px;">🗑️ Solicitar Baja</div>';
  html += '<div style="font-size: 11px; color: #6b7280; margin-bottom: 10px;">Elimina el documento del sistema (acción permanente)</div>';
  html += '<button onclick="mostrarEjemploAccionDEMO(\'baja\')" style="width: 100%; padding: 8px; background: #ef4444; color: #fff; border: none; border-radius: 4px; font-size: 11px; cursor: pointer; font-family: inherit; font-weight: 500;">Ver Ejemplo</button>';
  html += '</div>';
  
  html += '</div>';
  html += '</div>';

  mainContent.insertAdjacentHTML('afterbegin', html);
}


/**
 * Inserta un banner de advertencia regulatorio al inicio del documento cuando se da de baja.
 * Esto demuestra interacción directa escribiendo en el Office Client (Content Pane).
 */
async function escribirBajaEnDocumento() {
  await Word.run(async (context) => {
    // 1. Obtenemos el cuerpo del documento
    const body = context.document.body;
    
    // 2. Insertamos un párrafo destacado al inicio absoluto ('Start')
    const parrafo = body.insertParagraph('⚠️ DOCUMENTO DADO DE BAJA DEFINITIVAMENTE POR EL EQUIPO DE COMPLIANCE', 'Start');
    
    // 3. Le aplicamos el formato con el estilo de alerta (Rojo oscuro)
    parrafo.font.name = 'Segoe UI';
    parrafo.font.size = 13;
    parrafo.font.bold = true;
    parrafo.font.color = '#7f1d1d'; // Rojo corporativo de baja
    parrafo.spacingBefore = 10;
    parrafo.spacingAfter = 10;
    
    // 4. Añadimos una línea divisoria gris justo debajo para separar el aviso del texto original
    const linea = parrafo.insertParagraph('────────────────────────────────────────────────────────', 'After');
    linea.font.color = '#cbd5e1';
    linea.spacingAfter = 15;

    // 5. Sincronizamos con Word para aplicar los cambios en caliente
    await context.sync();
    console.log('[Word API] Alerta de baja inyectada en el documento de Word.');
  });
}


// ══════════════════════════════════════════════════════════════
// CAPA 2: PUNTO DE ENTRADA — wrapper con lock anti-concurrencia
// ══════════════════════════════════════════════════════════════
 
function escribirPublicacionEnDocumento() {
  if (_caratulaPromise) {
    console.log('[Caratula] Invocación en curso, se reutiliza.');
    return _caratulaPromise;
  }
  _caratulaPromise = (async () => {
    const datos = await obtenerDatosCaratula();
    return await _ejecutarCaratula(datos);
  })().finally(() => { _caratulaPromise = null; });
  return _caratulaPromise;
}
 
 
// ══════════════════════════════════════════════════════════════
// CAPA 3: DECISIÓN CREATE vs UPDATE
// ══════════════════════════════════════════════════════════════
 
async function _ejecutarCaratula(datos) {
  await Word.run(async (context) => {
    const ccRaiz = context.document.contentControls.getByTag(TAG_RAIZ);
    ccRaiz.load("items/id");
    await context.sync();
 
    const existe = ccRaiz.items.length > 0;
    console.log(`[Caratula] Modo: ${existe ? 'UPDATE' : 'CREATE'} | Dueños: ${datos.duenos.length} | Aprobadores: ${datos.aprobadores.length}`);
 
    if (existe) {
      await _actualizarCaratula(context, datos);
    } else {
      await _crearCaratula(context, datos);
    }
    await context.sync();
    console.log(`[Word API] Carátula ${existe ? 'actualizada' : 'creada'}.`);
  });
}

// ══════════════════════════════════════════════════════════════
// LOCALIZAR TABLAS POR HEADER (primera celda)
// Es el mecanismo de "ID" para tablas dinámicas en Word Web.
// ══════════════════════════════════════════════════════════════
 
async function _localizarTablaPorHeader(context, textoHeader) {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
 
  // Cargar el texto de la primera celda de cada tabla
  const primerasCeldas = [];
  for (const tabla of tables.items) {
    const fila = tabla.rows.getFirst();
    const celda = fila.cells.getFirstOrNullObject();
    celda.body.load("text");
    primerasCeldas.push({ tabla, celda });
  }
  await context.sync();
 
  // Match por texto (trim para tolerar espacios)
  const buscar = textoHeader.trim();
  for (const { tabla, celda } of primerasCeldas) {
    if (celda.isNullObject) continue;
    if ((celda.body.text || "").trim() === buscar) {
      return tabla;
    }
  }
  return null;
}
 


// ══════════════════════════════════════════════════════════════
// MODO UPDATE — refresca valores fijos, agrega fila al historial,
//                recrea tablas de participantes
// ══════════════════════════════════════════════════════════════
 
async function _actualizarCaratula(context, d) {
  // ─── 1) Actualizar campos fijos por tag (CCs en celdas) ───
  const updates = {
    "gp.header.titulo":   d.titulo,
    "gp.header.codigo":   "Código: " + d.codigo,
    "gp.header.version":  "Versión: " + d.version,
    "gp.header.fecha":    "Fecha de publicación:\n" + d.fecha,
    "gp.header.gerencia": "Gerencia Responsable:\n" + d.gerencia,
    "gp.id.codigo":       d.codigo,
    "gp.id.version":      d.version,
    "gp.id.gerencia":     d.gerencia,
    "gp.id.fecha":        d.fecha
  };
 
  const ccsPorTag = {};
  for (const tag of Object.keys(updates)) {
    ccsPorTag[tag] = context.document.contentControls.getByTag(tag);
    ccsPorTag[tag].load("items");
  }
  await context.sync();
 
  for (const [tag, valor] of Object.entries(updates)) {
    const items = ccsPorTag[tag].items;
    if (items.length > 0) {
      items[0].insertText(valor, "Replace");
    }
  }
 
  // ─── 2) HISTORIAL: agregar fila nueva al final ───
  const tablaHist = await _localizarTablaPorHeader(context, HEADER_HIST);
  if (tablaHist) {
    tablaHist.addRows("End", 1, [[
      d.aprobador, d.gerencia, d.version, d.fecha, d.detalleVersion
    ]]);
    await context.sync();
    // Estilo de la fila nueva (la última)
    const rows = tablaHist.rows;
    rows.load("items");
    await context.sync();
    const filaNueva = rows.items[rows.items.length - 1];
    await _estilarFilaDatos(context, filaNueva, { ultimaItalica: true });
  } else {
    console.warn('[Caratula] No se encontró tabla de historial.');
  }
 
  // ─── 3) Sincronizar ELABORADO POR ───
  const tablaElab = await _localizarTablaPorHeader(context, HEADER_ELAB);
  if (tablaElab) {
    await _sincronizarFilasTabla(context, tablaElab, "Elaborado por:", d.duenos);
  } else {
    console.warn('[Caratula] No se encontró tabla "Elaborado por:".');
  }

  // ─── 4) Sincronizar REVISADO POR ───
  const tablaRev = await _localizarTablaPorHeader(context, HEADER_REV);
  if (tablaRev) {
    await _sincronizarFilasTabla(context, tablaRev, "Revisado por:", d.aprobadores);
  } else {
    console.warn('[Caratula] No se encontró tabla "Revisado por:".');
  }
}
 
/**
 * Sincroniza las filas de una tabla con la nueva lista de participantes.
 * En lugar de borrar la tabla completa (que falla en Word Online dentro
 * de un CC), opera fila por fila:
 *   - Si necesita más filas → addRows("End", N)
 *   - Si necesita menos → row.delete() de las sobrantes
 *   - Reescribe el contenido de las filas de datos existentes
 *
 * Estructura asumida de la tabla:
 *   Fila 0: header con [etiqueta, "Gerencia", "Área", "Cargo"]
 *   Filas 1..N: datos de cada participante
 */
async function _sincronizarFilasTabla(context, tabla, etiqueta, participantes) {
  const filas = participantes.length > 0
    ? participantes.map(p => [p.nombre, p.gerencia, p.area, p.cargo])
    : [["(sin asignados)", "", "", ""]];
 
  const filasDeseadas = filas.length;
 
  // Cargar filas actuales
  const rows = tabla.rows;
  rows.load("items");
  await context.sync();
 
  const filasActualesDatos = rows.items.length - 1;  // descontamos el header
  const diferencia = filasDeseadas - filasActualesDatos;
 
  // ─── Caso 1: agregar filas faltantes ───
  if (diferencia > 0) {
    const nuevasFilasData = filas.slice(filasActualesDatos);
    tabla.addRows("End", diferencia, nuevasFilasData);
    await context.sync();
    rows.load("items");
    await context.sync();
  }
  // ─── Caso 2: eliminar filas sobrantes ───
  else if (diferencia < 0) {
    const aBorrar = Math.abs(diferencia);
    for (let i = 0; i < aBorrar; i++) {
      const ultimaFila = rows.items[rows.items.length - 1 - i];
      ultimaFila.delete();
    }
    await context.sync();
    rows.load("items");
    await context.sync();
  }
 
  // ─── Reescribir contenido de las filas de datos ───
  // (se hace siempre, para refrescar nombres aunque la cantidad no cambie)
  for (let i = 0; i < filasDeseadas; i++) {
    const fila = rows.items[i + 1];  // +1 para saltar header
    const cells = fila.cells;
    cells.load("items");
    await context.sync();
    for (let j = 0; j < cells.items.length && j < 4; j++) {
      cells.items[j].body.insertText(filas[i][j] || "", "Replace");
    }
  }
 
  // Reaplicar estilos (por si las filas nuevas no los heredaron)
  await _aplicarFuenteTabla(context, tabla, 10);
  await _pintarHeaderRow(context, tabla.rows.getFirst());
}
 
 
// ══════════════════════════════════════════════════════════════
// MODO CREATE — inserta toda la carátula desde cero
// ══════════════════════════════════════════════════════════════

async function _crearCaratula(context, d) {
  const body = context.document.body;
  const VERDE_CLARO = "#E8F5EE";
  const VERDE = "#00843D";
  const TEXTO = "#1F2937";
 
  // Helper: envuelve celda en CC con tag
  const envolver = async (cell, tag, valor) => {
    const cc = cell.body.insertContentControl();
    cc.tag = tag;
    cc.title = tag;
    cc.appearance = "Hidden";
    cc.cannotEdit = false;
    // NO usamos cannotDelete porque puede fallar en Word Web
    cc.insertText(valor, "Replace");
    return cc;
  };
 
  // ─── 1) HEADER SUPERIOR ───
  const headerData = [
    ["BANCO FALABELLA", d.titulo, "Código: " + d.codigo],
    [
      "Versión: " + d.version + "\nDocumento Interno",
      "Gerencia Responsable:\n" + d.gerencia,
      "Fecha de publicación:\n" + d.fecha
    ]
  ];
  const tablaHeader = body.insertTable(2, 3, "Start", headerData);
  await context.sync();
  await _aplicarFuenteTabla(context, tablaHeader, 9);
 
  const filaH0 = tablaHeader.rows.getFirst();
  filaH0.shadingColor = VERDE_CLARO;
  const cellsH0 = filaH0.cells;
  cellsH0.load("items");
  await context.sync();
  cellsH0.items.forEach((cell, i) => {
    cell.body.font.color = VERDE;
    cell.body.font.bold = true;
    cell.body.font.size = i === 1 ? 13 : 10;
    if (i === 1) cell.horizontalAlignment = "Centered";
  });
 
  await envolver(cellsH0.items[1], "gp.header.titulo", d.titulo);
  await envolver(cellsH0.items[2], "gp.header.codigo", "Código: " + d.codigo);
 
  const filaH1 = filaH0.getNext();
  const cellsH1 = filaH1.cells;
  cellsH1.load("items");
  await context.sync();
  if (cellsH1.items[1]) cellsH1.items[1].horizontalAlignment = "Centered";
  await envolver(cellsH1.items[0], "gp.header.version", "Versión: " + d.version);
  await envolver(cellsH1.items[1], "gp.header.gerencia", "Gerencia Responsable:\n" + d.gerencia);
  await envolver(cellsH1.items[2], "gp.header.fecha", "Fecha de publicación:\n" + d.fecha);
 
  // ─── 2) TÍTULOS ───
  const t1 = tablaHeader.insertParagraph("1. DESCRIPCIÓN GENERAL DEL DOCUMENTO", "After");
  t1.font.name = "Segoe UI";
  t1.font.size = 14;
  t1.font.bold = true;
  t1.font.color = TEXTO;
  t1.spaceBefore = 24;
  t1.spaceAfter = 8;
 
  const t11 = t1.insertParagraph("1.1 Control del Documento", "After");
  t11.font.name = "Segoe UI";
  t11.font.size = 11;
  t11.font.bold = true;
  t11.font.color = TEXTO;
  t11.spaceBefore = 12;
  t11.spaceAfter = 6;
 
  // ─── 3) FICHA DE IDENTIFICACIÓN ───
  const identificacionData = [
    ["Ficha de Identificación del Documento", ""],
    ["Código", d.codigo],
    ["Versión", d.version],
    ["Gerencia Responsable", d.gerencia],
    ["Fecha de Publicación", d.fecha]
  ];
  const tablaId = t11.insertTable(5, 2, "After", identificacionData);
  await context.sync();
  await _aplicarFuenteTabla(context, tablaId, 10);
  await _pintarHeaderRow(context, tablaId.rows.getFirst());
 
  const rowsId = tablaId.rows;
  rowsId.load("items");
  await context.sync();
 
  const cellsArrId = [];
  for (let i = 1; i < rowsId.items.length; i++) {
    const cells = rowsId.items[i].cells;
    cells.load("items");
    cellsArrId.push(cells);
  }
  await context.sync();
 
  const tagsId = ["gp.id.codigo", "gp.id.version", "gp.id.gerencia", "gp.id.fecha"];
  const valoresId = [d.codigo, d.version, d.gerencia, d.fecha];
 
  for (let i = 0; i < cellsArrId.length; i++) {
    const cells = cellsArrId[i].items;
    cells[0].shadingColor = VERDE_CLARO;
    cells[0].body.font.bold = true;
    cells[0].body.font.color = VERDE;
    await envolver(cells[1], tagsId[i], valoresId[i]);
  }
 
  // ─── 4) HISTORIAL DE VERSIONES ───
  const histTitulo = tablaId.insertParagraph("Historial de Versiones", "After");
  histTitulo.font.name = "Segoe UI";
  histTitulo.font.size = 11;
  histTitulo.font.bold = true;
  histTitulo.font.color = TEXTO;
  histTitulo.spaceBefore = 16;
  histTitulo.spaceAfter = 6;
 
  const historialData = [
    ["Aprobado por:", "Gerencia", "Versión modificada", "Fecha", "Detalle de actualización"],
    [d.aprobador, d.gerencia, d.version, d.fecha, d.detalleVersion]
  ];
  const tablaHist = histTitulo.insertTable(2, 5, "After", historialData);
  await context.sync();
  await _aplicarFuenteTabla(context, tablaHist, 10);
  await _pintarHeaderRow(context, tablaHist.rows.getFirst());
 
  const filaHistDatos = tablaHist.rows.getFirst().getNext();
  await _estilarFilaDatos(context, filaHistDatos, { ultimaItalica: true });
 
  // ─── 5) TÍTULO 1.2 ───
  const t12 = tablaHist.insertParagraph("1.2 Autores y revisores del documento (versión actual)", "After");
  t12.font.name = "Segoe UI";
  t12.font.size = 11;
  t12.font.bold = true;
  t12.font.color = TEXTO;
  t12.spaceBefore = 16;
  t12.spaceAfter = 6;
 
  // ─── 6) ELABORADO POR (Dueños) ───
  const duenosFilas = d.duenos.length > 0
    ? d.duenos.map(p => [p.nombre, p.gerencia, p.area, p.cargo])
    : [["(sin asignados)", "", "", ""]];
  const elabData = [
    ["Elaborado por:", "Gerencia", "Área", "Cargo"],
    ...duenosFilas
  ];
  const tablaElab = t12.insertTable(elabData.length, 4, "After", elabData);
  await context.sync();
  await _aplicarFuenteTabla(context, tablaElab, 10);
  await _pintarHeaderRow(context, tablaElab.rows.getFirst());
 
  // ─── 7) REVISADO POR (Aprobadores) ───
  const aprobFilas = d.aprobadores.length > 0
    ? d.aprobadores.map(p => [p.nombre, p.gerencia, p.area, p.cargo])
    : [["(sin asignados)", "", "", ""]];
  const revData = [
    ["Revisado por:", "Gerencia", "Área", "Cargo"],
    ...aprobFilas
  ];
  const pSepar = tablaElab.insertParagraph("", "After");
  pSepar.spaceAfter = 4;
  const tablaRev = pSepar.insertTable(revData.length, 4, "After", revData);
  await context.sync();
  await _aplicarFuenteTabla(context, tablaRev, 10);
  await _pintarHeaderRow(context, tablaRev.rows.getFirst());
 
  // ─── 8) SALTO DE PÁGINA ───
  const pCierre = tablaRev.insertParagraph("", "After");
  pCierre.insertBreak(Word.BreakType.page, "After");
 
  // ─── 9) CC RAÍZ que envuelve TODA la carátula ───
  // Es lo único que envolvemos a nivel de "tabla múltiple", y lo
  // hacemos solo una vez (en CREATE), nunca lo recreamos.
  const rangoCaratula = tablaHeader.getRange().expandTo(pCierre.getRange());
  const ccRaiz = rangoCaratula.insertContentControl();
  ccRaiz.tag = TAG_RAIZ;
  ccRaiz.title = "Carátula GP";
  ccRaiz.appearance = "Hidden";
}

// ══════════════════════════════════════════════════════════════
// HELPERS DE ESTILO (compartidos entre CREATE y UPDATE)
// ══════════════════════════════════════════════════════════════
 
async function _aplicarFuenteTabla(context, table, sizeBase) {
  const TEXTO = "#1F2937";
  const tamano = sizeBase || 10;
  const rows = table.rows;
  rows.load("items");
  await context.sync();
  for (const row of rows.items) {
    const cells = row.cells;
    cells.load("items");
    await context.sync();
    cells.items.forEach(cell => {
      cell.body.font.name = "Segoe UI";
      cell.body.font.size = tamano;
      cell.body.font.color = TEXTO;
    });
  }
}
 
async function _pintarHeaderRow(context, row) {
  const VERDE = "#00843D";
  row.shadingColor = VERDE;
  const cells = row.cells;
  cells.load("items");
  await context.sync();
  cells.items.forEach(cell => {
    cell.body.font.name = "Segoe UI";
    cell.body.font.size = 10;
    cell.body.font.color = "#FFFFFF";
    cell.body.font.bold = true;
  });
}
 
async function _estilarFilaDatos(context, row, opts) {
  const VERDE = "#00843D";
  const TEXTO = "#1F2937";
  const cells = row.cells;
  cells.load("items");
  await context.sync();
  cells.items.forEach((cell, i) => {
    cell.body.font.name = "Segoe UI";
    cell.body.font.size = 10;
    cell.body.font.color = TEXTO;
    if (opts && opts.ultimaItalica && i === cells.items.length - 1) {
      cell.body.font.color = VERDE;
      cell.body.font.italic = true;
    }
  });
}
 // ── Acción: Eliminar carátula (solo demo/compliance) ──
async function accion_eliminarCaratula() {
  const confirmado = await confirmar(
    'Esta acción no se puede deshacer. El documento podrá generar una carátula nueva en la próxima publicación.',
    'Eliminar carátula del documento',
    '#dc2626'  // rojo para indicar acción destructiva
  );
  if (!confirmado) return;

  try {
    await eliminarCaratula();
    console.log('[Caratula] Eliminada por acción del usuario.');
  } catch (e) {
    console.error('[Caratula] Error al eliminar:', e);
  }
}
 
// ══════════════════════════════════════════════════════════════
// UTILIDAD DE DESARROLLO: elimina la carátula completa
// Usar desde consola: window.eliminarCaratula()
// ══════════════════════════════════════════════════════════════
 
async function eliminarCaratula() {
  await Word.run(async (context) => {
    // ─── 1) Localizar el CC raíz ───
    const ccRaiz = context.document.contentControls.getByTag(TAG_RAIZ);
    ccRaiz.load("items");
    await context.sync();
 
    if (ccRaiz.items.length === 0) {
      console.log('[Caratula] No hay carátula que eliminar.');
      return;
    }
 
    const cc = ccRaiz.items[0];
 
    // ─── 2) Vaciar todo el contenido del CC en una sola operación ───
    // En Word Online esto es mucho más estable que borrar tabla por tabla.
    // El servidor lo procesa como un único reemplazo atómico, no como
    // 20+ deletes que pueden corromper el estado de colaboración.
    try {
      cc.insertText("", "Replace");
      await context.sync();
    } catch (e) {
      // Si falla insertText, intentamos getRange().clear() como fallback
      console.warn('[Caratula] insertText falló, intento con clear:', e.message);
      try {
        cc.getRange().clear();
        await context.sync();
      } catch (e2) {
        console.error('[Caratula] No se pudo vaciar el contenido:', e2.message);
        throw e2;
      }
    }
 
    // ─── 3) Invalidar el CC raíz cambiando su tag ───
    // El CC físicamente sigue existiendo (Word Online no permite borrar
    // CCs que envuelven estructuras), pero ya no responde al getByTag(TAG_RAIZ)
    // por lo que la próxima publicación lo tratará como CREATE.
    // Como insertText vació el contenido, queda como un párrafo vacío oculto.
    try {
      cc.tag = TAG_RAIZ + "_eliminado_" + Date.now();
      cc.title = "Carátula eliminada";
      await context.sync();
      console.log('[Caratula] Eliminada correctamente.');
    } catch (e) {
      console.warn('[Caratula] No se pudo renombrar el tag:', e.message);
    }
  });
}
if (typeof window !== 'undefined') window.eliminarCaratula = eliminarCaratula;


/**
 * Muestra un ejemplo de acción sin hacer cambios reales
 */
function mostrarEjemploAccionDEMO(tipoAccion) {
  if (!state.demoMode) return;
  
  const ejemplos = {
    'solicitar': {
      titulo: '📤 Solicitar Publicación',
      descripcion: 'Se envía el documento a revisión del equipo de Compliance.',
      paso_actual: '2/5: Solicitud de Publicación',
      paso_siguiente: '3/5: En Compliance',
      cambios: [
        '✓ Estado cambia de "Borrador" a "Solicitud_Publicacion"',
        '✓ Se notifica al equipo de Compliance',
        '✓ El documento queda en estado de espera',
        '✓ El Dueño puede solicitar prórroga si es necesario'
      ]
    },
    'proroga': {
      titulo: '📅 Solicitar Prórroga',
      descripcion: 'Se extiende el plazo de revisión del documento.',
      paso_actual: '2/5: Solicitud de Publicación',
      paso_siguiente: '2/5: Solicitud de Publicación (Con prórroga)',
      cambios: [
        '✓ Se abre un diálogo para seleccionar nueva fecha',
        '✓ Se registra el motivo de la prórroga',
        '✓ Se notifica a todos los revisores',
        '✓ El plazo se extiende hasta la fecha especificada'
      ]
    },
    'baja': {
      titulo: '🗑️ Solicitar Baja',
      descripcion: 'El documento se elimina del sistema (acción irreversible).',
      paso_actual: '2/5: Solicitud de Publicación',
      paso_siguiente: 'Eliminado/Archivado',
      cambios: [
        '⚠️ Se solicita confirmación de baja',
        '⚠️ Es necesario ingresar motivo obligatoriamente',
        '⚠️ El documento se marca como "Dado de Baja"',
        '⚠️ Se crea registro de auditoría de la baja'
      ]
    }
  };
  
  const ejemplo = ejemplos[tipoAccion];
  if (!ejemplo) return;
  
  // Crear modal con ejemplo
  let html = '<div style="';
  html += 'position: fixed; inset: 0; background: rgba(0,0,0,0.5);';
  html += 'display: flex; align-items: center; justify-content: center;';
  html += 'z-index: 2000; padding: 20px;';
  html += '" onclick="this.remove()">';
  
  html += '<div style="';
  html += 'background: #fff; border-radius: 12px; max-width: 500px;';
  html += 'width: 100%; padding: 24px; box-shadow: 0 20px 60px rgba(0,0,0,0.3);';
  html += '" onclick="event.stopPropagation()">';
  
  html += '<div style="font-size: 24px; margin-bottom: 12px;">✨</div>';
  html += '<h2 style="margin: 0 0 12px 0; font-size: 18px; font-weight: 700;">' + ejemplo.titulo + '</h2>';
  html += '<p style="margin: 0 0 16px 0; font-size: 13px; color: #6b7280; line-height: 1.5;">' + ejemplo.descripcion + '</p>';
  
  html += '<div style="padding: 12px; background: #f3f4f6; border-radius: 6px; margin-bottom: 16px;">';
  html += '<div style="font-size: 11px; font-weight: 600; color: #6b7280; text-transform: uppercase; margin-bottom: 8px;">Paso Actual</div>';
  html += '<div style="font-size: 13px; font-weight: 600; color: #1f2937;">' + ejemplo.paso_actual + '</div>';
  html += '<div style="font-size: 11px; color: #9ca3af; margin-top: 8px;">↓</div>';
  html += '<div style="font-size: 11px; font-weight: 600; color: #6b7280; text-transform: uppercase; margin-top: 8px;">Paso Siguiente</div>';
  html += '<div style="font-size: 13px; font-weight: 600; color: #10b981;">' + ejemplo.paso_siguiente + '</div>';
  html += '</div>';
  
  html += '<div style="margin-bottom: 16px;">';
  html += '<div style="font-size: 11px; font-weight: 600; color: #6b7280; text-transform: uppercase; margin-bottom: 8px;">Cambios que ocurrirían</div>';
  html += '<div style="display: flex; flex-direction: column; gap: 8px;">';
  
  ejemplo.cambios.forEach(cambio => {
    html += '<div style="font-size: 12px; color: #374151; padding-left: 20px; position: relative;">';
    html += '<span style="position: absolute; left: 0;">' + (cambio.includes('⚠️') ? '⚠️' : '✓') + '</span>';
    html += cambio.substring(2);
    html += '</div>';
  });
  
  html += '</div>';
  html += '</div>';
  
  html += '<div style="display: flex; gap: 10px;">';
  html += '<button onclick="this.closest(\'div\').parentElement.remove()" style="flex: 1; padding: 10px; background: #f3f4f6; border: none; border-radius: 6px; cursor: pointer; font-family: inherit; font-weight: 500;">Cerrar</button>';
  html += '<button onclick="this.closest(\'div\').parentElement.remove(); simularAccionDemo(\'' + tipoAccion + '\')" style="flex: 1; padding: 10px; background: #3b82f6; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-family: inherit; font-weight: 500;">Simular Acción</button>';
  html += '</div>';
  
  html += '</div>';
  html += '</div>';
  
  document.body.insertAdjacentHTML('beforeend', html);
}

// Agregar al export
window.mostrarEjemploAccionDEMO = mostrarEjemploAccionDEMO;