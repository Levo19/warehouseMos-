# warehouseMos 🏭

PWA de gestión de almacén central para **InversionMos**. Aplicación hermana de MosExpress.

## Stack
- **Frontend**: Vanilla JS + Tailwind CSS → GitHub Pages
- **Backend**: Google Apps Script (Web App)
- **Base de datos**: Google Sheets (Spreadsheet propio, independiente de MosExpress)
- **Impresión**: PrintNode (etiquetas adhesivas ZPL para productos envasados)
- **Scanner**: ZXing via cámara del celular

## Setup inicial

### 1. Google Apps Script
1. Crear nuevo proyecto en script.google.com
2. Copiar todos los archivos de `gas/` al proyecto GAS
3. Ejecutar `setupWarehouse()` **una sola vez** — crea el Spreadsheet con datos de prueba
4. Copiar el `SPREADSHEET_ID` que aparece en el log
5. Desplegar como **Web App**: Ejecutar como _Yo_, Acceso _Cualquiera con el enlace_
6. Copiar la URL del Web App desplegado

### 2. Frontend (GitHub Pages)
1. En `index.html` buscar `window.WH_CONFIG`
2. Pegar la URL del Web App en `gasUrl: ''`
3. Push a `main` → GitHub Pages activa en Settings > Pages

### 3. PrintNode (etiquetas adhesivas)
1. Crear cuenta en printnode.com
2. Instalar cliente PrintNode en la PC que tiene la impresora de etiquetas
3. En la app → Config → ingresar API Key y Printer ID

## Factor de conversión

`factorConversion` en PRODUCTOS: unidades del derivado por 1 unidad del producto base.

| Base | Unidad | Derivado | Factor | Resultado |
|------|--------|----------|--------|-----------|
| Arroz granel | SACO | Arroz 1kg | 50 | 50 bolsas/saco |
| Comino granel | KG | Comino 500g | 2 | 2 bolsas/kg |
| Comino granel | KG | Comino 100g | 10 | 10 sobres/kg |

```
unidadesEsperadas = cantidadBase × factorConversion × (1 - mermaEsperadaPct / 100)
```

## Hojas del Spreadsheet (auto-generadas por setupWarehouse)

CONFIG · CATEGORIAS · PRODUCTOS · STOCK · LOTES_VENCIMIENTO · PROVEEDORES · PREINGRESOS · GUIAS · GUIA_DETALLE · MERMAS · AUDITORIAS · AJUSTES · ENVASADOS · PRODUCTO_NUEVO · ZONAS
