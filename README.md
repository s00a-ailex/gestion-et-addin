# 📋 Gestión de Publicaciones — Word Add-in

Complemento de Word que agrega el panel de gestión del flujo de aprobación documental.
Se integra con la Azure Functions API (`atcazfunc.azurewebsites.net`).

---

## 📁 Estructura

```
gestion-publicaciones-addin/
├── manifest.xml        ← Definición del add-in (registrar en SharePoint)
├── taskpane.html       ← UI completa del panel
├── taskpane.js         ← Lógica, estado y llamadas a la API
├── commands.html       ← Requerido por el manifest
├── commands.js         ← Handlers de comandos del ribbon
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png
```

---

## ⚙️ Configuración antes de desplegar

### 1. Agregar la Function Key en taskpane.js

```javascript
const CONFIG = {
  API_BASE: 'https://atcazfunc.azurewebsites.net/api',
  FUNCTION_KEY: 'TU_FUNCTION_KEY_AQUI',  // ← reemplazar
};
```

Obtener la key desde:
**Portal Azure → atcazfunc → Funciones → [nombre función] → Claves de función → default**

O una key global:
**Portal Azure → atcazfunc → Claves de función de la aplicación → default**

---

### 2. Subir los archivos del add-in a Azure Functions

Los archivos HTML/JS deben estar accesibles en `https://atcazfunc.azurewebsites.net/`.
La forma más simple es servir archivos estáticos desde la Function App.

**Opción A — Carpeta `wwwroot` en Azure (recomendado):**
1. Portal Azure → atcazfunc → Herramientas avanzadas (Kudu) → Debug Console
2. Navegar a `site/wwwroot`
3. Subir: `taskpane.html`, `taskpane.js`, `commands.html`, `commands.js`, carpeta `assets/`

**Opción B — Azure Static Web Apps (si se quiere separar):**
Crear un Static Web App separado y actualizar las URLs en `manifest.xml`.

---

### 3. Íconos requeridos

Crear o colocar en la carpeta `assets/`:
- `icon-16.png`  (16×16 px)
- `icon-32.png`  (32×32 px)  
- `icon-80.png`  (80×80 px)

Puedes usar cualquier ícono PNG simple. Sin íconos, el add-in dará error al cargar.

---

## 🚀 Registrar en SharePoint (App Catalog)

1. Ir a:
   Necesito ser instalado por el administrador de Sharepoint de su organización.
   
   Si no existe el catálogo: **SharePoint Admin Center → Más características → Apps → App Catalog**

2. **Aplicaciones para Office → Cargar** → subir `manifest.xml`

3. **Hacer clic en "Confiar"** cuando lo solicite

4. En cualquier documento Word Online abierto desde SharePoint:
   - Pestaña **Insertar → Complementos → Complementos de la organización**
   - Seleccionar **Gestión de Publicaciones**

---

## 🔑 Cómo obtiene el Add-in el itemId del documento

El Add-in lee propiedades personalizadas del documento Word:

| Propiedad        | Descripción                              |
|------------------|------------------------------------------|
| `SP_ItemId`      | ID del item en la lista de SharePoint    |
| `SP_DriveId`     | Drive ID de la biblioteca                |
| `SP_DriveItemId` | Drive Item ID del archivo físico         |
| `SP_Nombre`      | Nombre del documento                     |

**Estas propiedades deben ser escritas por Power Automate al crear el documento.**

Flow recomendado en Power Automate:
```
Trigger: Al crear un archivo en la biblioteca
→ Obtener propiedades del archivo (DriveId, ItemId)  
→ Actualizar propiedades del archivo Word:
   SP_ItemId      = {ID del item de lista}
   SP_DriveId     = {Drive ID}
   SP_DriveItemId = {Drive Item ID}
   SP_Nombre      = {Nombre del archivo}
```

**Mientras no esté configurado Power Automate**, puedes pasar los valores por URL para desarrollo:
```
https://atcazfunc.azurewebsites.net/taskpane.html?itemId=42&driveId=b!xxx&driveItemId=01xxx&nombre=MiDocumento
```

---

## 👥 Gestión de roles (producción)

Actualmente el Add-in incluye un selector de rol para desarrollo.
En producción, reemplazar la lógica de `state.rolActual` con la identidad real del usuario M365:

```javascript
// En Office.onReady, después de obtener contexto:
const userEmail = Office.context.mailbox?.userProfile?.emailAddress
               || Office.context.roamingSettings?.get('userEmail');

// Luego consultar un endpoint de la API que devuelva el rol según el email
const rolInfo = await apiCall('GET', `sp/getRol?email=${userEmail}&itemId=${state.itemId}`);
state.rolActual = rolInfo.rol; // 'usuario' | 'compliance' | 'aprobador'
```

---

## 🗑️ Para producción: eliminar el selector de rol

En `taskpane.html`, comentar o eliminar el bloque:
```html
<!-- SELECTOR DE ROL (solo para desarrollo) -->
<div class="rol-selector" id="devRolSelector">
  ...
</div>
```
