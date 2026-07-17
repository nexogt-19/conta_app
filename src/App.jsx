import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ── PARTIDAS VENTAS — PRE-CARGADAS (P270-P273) ───────────────────
// 4 partidas semanales de ventas junio 2026 — cuadradas ✅
const PARTIDAS_INICIALES = [{"num": 270, "fecha": "2026-06-06", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 1", "lineas": [{"cta": "1.1.17.004", "debe": 17299.2, "haber": 0, "conc": "Cobros BI — Semana 1"}, {"cta": "1.1.17.005", "debe": 18855.0, "haber": 0, "conc": "Cobros Banrural — Semana 1"}, {"cta": "4.1.01", "debe": 0, "haber": 32280.54, "conc": "Ventas Semana 1 — 17 facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 3873.66, "conc": "IVA Debito Fiscal 12% — Semana 1"}]}, {"num": 271, "fecha": "2026-06-16", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 2", "lineas": [{"cta": "1.1.17.004", "debe": 20315.4, "haber": 0, "conc": "Cobros BI — Semana 2"}, {"cta": "1.1.17.005", "debe": 79060.0, "haber": 0, "conc": "Cobros Banrural — Semana 2"}, {"cta": "4.1.01", "debe": 0, "haber": 88728.04, "conc": "Ventas Semana 2 — 18 facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 10647.36, "conc": "IVA Debito Fiscal 12% — Semana 2"}]}, {"num": 272, "fecha": "2026-06-23", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 3", "lineas": [{"cta": "1.1.17.004", "debe": 20373.6, "haber": 0, "conc": "Cobros BI — Semana 3"}, {"cta": "1.1.17.005", "debe": 10790.0, "haber": 0, "conc": "Cobros Banrural — Semana 3"}, {"cta": "4.1.01", "debe": 0, "haber": 27824.66, "conc": "Ventas Semana 3 — 16 facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 3338.94, "conc": "IVA Debito Fiscal 12% — Semana 3"}]}, {"num": 273, "fecha": "2026-06-24", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 4", "lineas": [{"cta": "1.1.17.005", "debe": 295.0, "haber": 0, "conc": "Cobros Banrural — Semana 4"}, {"cta": "1.1.03", "debe": 1435.2, "haber": 0, "conc": "CxC Pendiente — Semana 4"}, {"cta": "4.1.01", "debe": 0, "haber": 1544.82, "conc": "Ventas Semana 4 — 2 facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 185.38, "conc": "IVA Debito Fiscal 12% — Semana 4"}]}, {"num": 274, "fecha": "2026-06-09", "tipo": "GASTO_SEM", "concepto": "Gastos Semana 1 (11 facturas)", "lineas": [{"cta": "6.1.10", "debe": 307.76, "haber": 0, "conc": "Combustible — Semana 1"}, {"cta": "6.1.14", "debe": 89.24, "haber": 0, "conc": "Insumos de Oficina — Semana 1"}, {"cta": "6.1.17", "debe": 445.53, "haber": 0, "conc": "Programas y Software — Semana 1"}, {"cta": "6.2.06", "debe": 1448.46, "haber": 0, "conc": "Paqueteria y Fletes — Semana 1"}, {"cta": "1.1.06", "debe": 274.92, "haber": 0, "conc": "IVA CF — Semana 1"}, {"cta": "1.1.02", "debe": 0, "haber": 2337.11, "conc": "Pago gastos Semana 1"}, {"cta": "2.1.01", "debe": 0, "haber": 228.8, "conc": "Pago gastos Semana 1"}]}, {"num": 275, "fecha": "2026-06-16", "tipo": "GASTO_SEM", "concepto": "Gastos Semana 2 (3 facturas)", "lineas": [{"cta": "6.1.17", "debe": 234.97, "haber": 0, "conc": "Mantenimiento y Reparaciones — Semana 2"}, {"cta": "6.2.06", "debe": 225.64, "haber": 0, "conc": "Servicio de Paqueteria — Semana 2"}, {"cta": "1.1.06", "debe": 55.28, "haber": 0, "conc": "IVA CF — Semana 2"}, {"cta": "1.1.17.004", "debe": 0, "haber": 263.17, "conc": "Pago gastos Semana 2"}, {"cta": "1.1.02", "debe": 0, "haber": 252.72, "conc": "Pago gastos Semana 2"}]}];

// ── DATOS JUNIO 2026 — PRE-CARGADOS DESDE FEL ─────────────────
// Ventas: 53 activas + 2 anuladas | Compras: 23 registros
// Procesado: 26/06/2026 02:46
const DATOS_INICIALES = {
  ventas: [{"id": "v2j26", "fecha": "2026-06-24", "serie": "11A5A333", "tipo": "FACT", "cliente": "CF", "base": 1281.43, "iva": 153.77, "total": 1435.2, "estado": "ACTIVA", "cobros": [{"id": "cxc2", "tipo": "banco", "cta": "1.1.03", "doc": "PENDIENTE", "monto": 1435.2}]}, {"id": "v3j26", "fecha": "2026-06-24", "serie": "F0CF052B", "tipo": "FACT", "cliente": "CF", "base": 263.39, "iva": 31.61, "total": 295.0, "estado": "ACTIVA", "cobros": [{"id": "rr3x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "438822207", "monto": 295.0}]}, {"id": "v4j26", "fecha": "2026-06-23", "serie": "3DC5C874", "tipo": "FACT", "cliente": "CF", "base": 1339.29, "iva": 160.71, "total": 1500.0, "estado": "ACTIVA", "cobros": [{"id": "bi4x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "188466", "monto": 1500.0}]}, {"id": "v5j26", "fecha": "2026-06-22", "serie": "F7C22F08", "tipo": "FACT", "cliente": "CF", "base": 1744.29, "iva": 209.31, "total": 1953.6, "estado": "ACTIVA", "cobros": [{"id": "bi5x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "159399", "monto": 1953.6}]}, {"id": "v6j26", "fecha": "2026-06-22", "serie": "F86B7249", "tipo": "FACT", "cliente": "13945009", "base": 2062.5, "iva": 247.5, "total": 2310.0, "estado": "ACTIVA", "cobros": [{"id": "bi6x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56405909", "monto": 1150.0}, {"id": "bi6x26", "tipo": "banco", "cta": "1.1.17.004", "doc": "56405911", "monto": 1160.0}]}, {"id": "v7j26", "fecha": "2026-06-22", "serie": "3AA36C8B", "tipo": "FACT", "cliente": "69975132", "base": 1392.86, "iva": 167.14, "total": 1560.0, "estado": "ACTIVA", "cobros": [{"id": "bi7x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "226810", "monto": 1560.0}]}, {"id": "v8j26", "fecha": "2026-06-22", "serie": "01B9F9CA", "tipo": "FACT", "cliente": "13495364", "base": 1696.43, "iva": 203.57, "total": 1900.0, "estado": "ACTIVA", "cobros": [{"id": "rr8x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "347946258", "monto": 1900.0}]}, {"id": "v9j26", "fecha": "2026-06-22", "serie": "00523115", "tipo": "FACT", "cliente": "CF", "base": 526.79, "iva": 63.21, "total": 590.0, "estado": "ACTIVA", "cobros": [{"id": "rr9x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "329635163", "monto": 590.0}]}, {"id": "v10j26", "fecha": "2026-06-20", "serie": "96236AAE", "tipo": "FACT", "cliente": "7914172", "base": 2767.86, "iva": 332.14, "total": 3100.0, "estado": "ACTIVA", "cobros": [{"id": "bi10x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "efectivo", "monto": 3100.0}]}, {"id": "v11j26", "fecha": "2026-06-19", "serie": "9478A4A8", "tipo": "FACT", "cliente": "109694252", "base": 848.21, "iva": 101.79, "total": 950.0, "estado": "ACTIVA", "cobros": [{"id": "rr11x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "304267352", "monto": 950.0}]}, {"id": "v12j26", "fecha": "2026-06-19", "serie": "1F8CA694", "tipo": "FACT", "cliente": "CF", "base": 1441.96, "iva": 173.04, "total": 1615.0, "estado": "ACTIVA", "cobros": [{"id": "rr12x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "295339904", "monto": 1615.0}]}, {"id": "v13j26", "fecha": "2026-06-19", "serie": "DAEE010B", "tipo": "FACT", "cliente": "CF", "base": 1330.36, "iva": 159.64, "total": 1490.0, "estado": "ACTIVA", "cobros": [{"id": "rr13x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "106911539", "monto": 590.0}, {"id": "rr13x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "123003232", "monto": 900.0}]}, {"id": "v14j26", "fecha": "2026-06-19", "serie": "8AEAE612", "tipo": "FACT", "cliente": "102283249", "base": 3214.29, "iva": 385.71, "total": 3600.0, "estado": "ACTIVA", "cobros": [{"id": "rr14x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "269087161", "monto": 3600.0}]}, {"id": "v15j26", "fecha": "2026-06-18", "serie": "07B3D3B0", "tipo": "FACT", "cliente": "109808045", "base": 4464.29, "iva": 535.71, "total": 5000.0, "estado": "ACTIVA", "cobros": [{"id": "bi15x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "853002", "monto": 5000.0}]}, {"id": "v16j26", "fecha": "2026-06-18", "serie": "71543EFE", "tipo": "FACT", "cliente": "355062", "base": 1562.5, "iva": 187.5, "total": 1750.0, "estado": "ACTIVA", "cobros": [{"id": "bi16x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56406057", "monto": 1750.0}]}, {"id": "v17j26", "fecha": "2026-06-18", "serie": "42C1DB52", "tipo": "FACT", "cliente": "CF", "base": 218.75, "iva": 26.25, "total": 245.0, "estado": "ACTIVA", "cobros": [{"id": "rr17x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "250930055", "monto": 245.0}]}, {"id": "v18j26", "fecha": "2026-06-18", "serie": "55372C73", "tipo": "FACT", "cliente": "117808105", "base": 357.14, "iva": 42.86, "total": 400.0, "estado": "ACTIVA", "cobros": [{"id": "rr18x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "202261849", "monto": 400.0}]}, {"id": "v19j26", "fecha": "2026-06-18", "serie": "F0E1BCAE", "tipo": "FACT", "cliente": "35527102", "base": 2857.14, "iva": 342.86, "total": 3200.0, "estado": "ACTIVA", "cobros": [{"id": "bi19x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56406056", "monto": 3200.0}]}, {"id": "v20j26", "fecha": "2026-06-16", "serie": "BBBB5A99", "tipo": "FACT", "cliente": "94078203", "base": 1066.96, "iva": 128.04, "total": 1195.0, "estado": "ACTIVA", "cobros": [{"id": "rr20x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "139521616", "monto": 1195.0}]}, {"id": "v21j26", "fecha": "2026-06-16", "serie": "7E316CB1", "tipo": "FACT", "cliente": "CF", "base": 398.57, "iva": 47.83, "total": 446.4, "estado": "ACTIVA", "cobros": [{"id": "bi21x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "222692", "monto": 446.4}]}, {"id": "v22j26", "fecha": "2026-06-15", "serie": "9B5297FC", "tipo": "FACT", "cliente": "306372010", "base": 1457.14, "iva": 174.86, "total": 1632.0, "estado": "ACTIVA", "cobros": [{"id": "bi22x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "172021", "monto": 1632.0}]}, {"id": "v23j26", "fecha": "2026-06-15", "serie": "34F7963D", "tipo": "FACT", "cliente": "CF", "base": 1410.71, "iva": 169.29, "total": 1580.0, "estado": "ACTIVA", "cobros": [{"id": "bi23x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "12146671", "monto": 1580.0}]}, {"id": "v24j26", "fecha": "2026-06-15", "serie": "08FA8FB8", "tipo": "FACT", "cliente": "CF", "base": 263.39, "iva": 31.61, "total": 295.0, "estado": "ACTIVA", "cobros": [{"id": "bi24x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "51622", "monto": 295.0}]}, {"id": "v25j26", "fecha": "2026-06-15", "serie": "8F93A123", "tipo": "FACT", "cliente": "CF", "base": 892.86, "iva": 107.14, "total": 1000.0, "estado": "ACTIVA", "cobros": [{"id": "rr25x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "44132624", "monto": 1000.0}]}, {"id": "v26j26", "fecha": "2026-06-15", "serie": "25D9DC13", "tipo": "FACT", "cliente": "13945009", "base": 2053.57, "iva": 246.43, "total": 2300.0, "estado": "ACTIVA", "cobros": [{"id": "bi26x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404451", "monto": 2300.0}]}, {"id": "v27j26", "fecha": "2026-06-15", "serie": "C695EC3B", "tipo": "FACT", "cliente": "98337211", "base": 267.86, "iva": 32.14, "total": 300.0, "estado": "ACTIVA", "cobros": [{"id": "bi27x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404453", "monto": 300.0}]}, {"id": "v28j26", "fecha": "2026-06-15", "serie": "824C3063", "tipo": "FACT", "cliente": "50637460", "base": 803.57, "iva": 96.43, "total": 900.0, "estado": "ACTIVA", "cobros": [{"id": "bi28x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404457", "monto": 900.0}]}, {"id": "v29j26", "fecha": "2026-06-13", "serie": "8D2EB768", "tipo": "FACT", "cliente": "115998020", "base": 14875.0, "iva": 1785.0, "total": 16660.0, "estado": "ACTIVA", "cobros": [{"id": "rr29x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "100882549", "monto": 16660.0}]}, {"id": "v30j26", "fecha": "2026-06-13", "serie": "9B43E674", "tipo": "FACT", "cliente": "46523626", "base": 1276.79, "iva": 153.21, "total": 1430.0, "estado": "ACTIVA", "cobros": [{"id": "bi30x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404452", "monto": 1430.0}]}, {"id": "v31j26", "fecha": "2026-06-11", "serie": "43D239C7", "tipo": "FACT", "cliente": "115998020", "base": 44406.25, "iva": 5328.75, "total": 49735.0, "estado": "ACTIVA", "cobros": [{"id": "bi31x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404460", "monto": 200.0}, {"id": "rr31x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "38076486", "monto": 49535.0}]}, {"id": "v32j26", "fecha": "2026-06-11", "serie": "14D3D855", "tipo": "FACT", "cliente": "115998020", "base": 16935.0, "iva": 2032.2, "total": 18967.2, "estado": "ANULADO", "cobros": []}, {"id": "v33j26", "fecha": "2026-06-11", "serie": "0060FFF1", "tipo": "FACT", "cliente": "115998020", "base": 49169.64, "iva": 5900.36, "total": 55070.0, "estado": "ANULADO", "cobros": []}, {"id": "v34j26", "fecha": "2026-06-11", "serie": "2C752484", "tipo": "FACT", "cliente": "13945009", "base": 7026.79, "iva": 843.21, "total": 7870.0, "estado": "ACTIVA", "cobros": [{"id": "rr34x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295124", "monto": 1035.0}, {"id": "rr34x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295123", "monto": 2900.0}, {"id": "rr34x41", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295122", "monto": 2900.0}, {"id": "rr34x44", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295121", "monto": 1035.0}]}, {"id": "v35j26", "fecha": "2026-06-11", "serie": "50B87C03", "tipo": "FACT", "cliente": "90134346", "base": 1200.0, "iva": 144.0, "total": 1344.0, "estado": "ACTIVA", "cobros": [{"id": "bi35x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "130778", "monto": 1344.0}]}, {"id": "v36j26", "fecha": "2026-06-11", "serie": "EBEDAB74", "tipo": "FACT", "cliente": "89056876", "base": 1339.29, "iva": 160.71, "total": 1500.0, "estado": "ACTIVA", "cobros": [{"id": "rr36x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "2079830743", "monto": 1500.0}]}, {"id": "v37j26", "fecha": "2026-06-11", "serie": "8E794B90", "tipo": "FACT", "cliente": "48956449", "base": 4414.29, "iva": 529.71, "total": 4944.0, "estado": "ACTIVA", "cobros": [{"id": "bi37x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "167359", "monto": 4944.0}]}, {"id": "v38j26", "fecha": "2026-06-11", "serie": "8B336FBE", "tipo": "FACT", "cliente": "5504538", "base": 4414.29, "iva": 529.71, "total": 4944.0, "estado": "ACTIVA", "cobros": [{"id": "bi38x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "167359", "monto": 4944.0}]}, {"id": "v39j26", "fecha": "2026-06-11", "serie": "07FC2E6C", "tipo": "FACT", "cliente": "CF", "base": 1160.71, "iva": 139.29, "total": 1300.0, "estado": "ACTIVA", "cobros": [{"id": "rr39x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1890192014", "monto": 1300.0}]}, {"id": "v40j26", "fecha": "2026-06-06", "serie": "3096928C", "tipo": "FACT", "cliente": "107462656", "base": 1071.43, "iva": 128.57, "total": 1200.0, "estado": "ACTIVA", "cobros": [{"id": "rr40x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "2005710533", "monto": 1200.0}]}, {"id": "v41j26", "fecha": "2026-06-05", "serie": "5458ECB0", "tipo": "FACT", "cliente": "CF", "base": 441.43, "iva": 52.97, "total": 494.4, "estado": "ACTIVA", "cobros": [{"id": "bi41x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "229586", "monto": 494.4}]}, {"id": "v42j26", "fecha": "2026-06-05", "serie": "BFDAEB4A", "tipo": "FACT", "cliente": "57140626", "base": 1855.71, "iva": 222.69, "total": 2078.4, "estado": "ACTIVA", "cobros": [{"id": "bi42x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "229586", "monto": 2078.4}]}, {"id": "v43j26", "fecha": "2026-06-05", "serie": "5B792CDD", "tipo": "FACT", "cliente": "55359337", "base": 3102.86, "iva": 372.34, "total": 3475.2, "estado": "ACTIVA", "cobros": [{"id": "bi43x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "129700", "monto": 3475.2}]}, {"id": "v44j26", "fecha": "2026-06-05", "serie": "95291FD3", "tipo": "FACT", "cliente": "13945009", "base": 8062.5, "iva": 967.5, "total": 9030.0, "estado": "ACTIVA", "cobros": [{"id": "rr44x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820493", "monto": 905.0}, {"id": "rr44x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820494", "monto": 905.0}, {"id": "rr44x41", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820495", "monto": 5520.0}, {"id": "rr44x44", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820496", "monto": 1700.0}]}, {"id": "v45j26", "fecha": "2026-06-04", "serie": "9877425D", "tipo": "FACT", "cliente": "CF", "base": 1250.0, "iva": 150.0, "total": 1400.0, "estado": "ACTIVA", "cobros": [{"id": "bi45x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404455", "monto": 1400.0}]}, {"id": "v46j26", "fecha": "2026-06-04", "serie": "43090455", "tipo": "FACT", "cliente": "30677114", "base": 3125.0, "iva": 375.0, "total": 3500.0, "estado": "ACTIVA", "cobros": [{"id": "bi46x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "174780", "monto": 3500.0}]}, {"id": "v47j26", "fecha": "2026-06-03", "serie": "AC795D67", "tipo": "FACT", "cliente": "108841014", "base": 1414.29, "iva": 169.71, "total": 1584.0, "estado": "ACTIVA", "cobros": [{"id": "bi47x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "158398", "monto": 1584.0}]}, {"id": "v48j26", "fecha": "2026-06-03", "serie": "7E380EDB", "tipo": "FACT", "cliente": "CF", "base": 655.71, "iva": 78.69, "total": 734.4, "estado": "ACTIVA", "cobros": [{"id": "bi48x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "158398", "monto": 734.4}]}, {"id": "v49j26", "fecha": "2026-06-03", "serie": "81A3472A", "tipo": "FACT", "cliente": "103805311", "base": 527.14, "iva": 63.26, "total": 590.4, "estado": "ACTIVA", "cobros": [{"id": "bi49x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "176574", "monto": 590.4}]}, {"id": "v50j26", "fecha": "2026-06-03", "serie": "F61EFC67", "tipo": "FACT", "cliente": "CF", "base": 484.29, "iva": 58.11, "total": 542.4, "estado": "ACTIVA", "cobros": [{"id": "bi50x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "281266", "monto": 542.4}]}, {"id": "v51j26", "fecha": "2026-06-02", "serie": "7398F13C", "tipo": "FACT", "cliente": "117808105", "base": 4285.71, "iva": 514.29, "total": 4800.0, "estado": "ACTIVA", "cobros": [{"id": "rr51x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1854380599", "monto": 2800.0}, {"id": "rr51x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "1854356679", "monto": 2000.0}]}, {"id": "v52j26", "fecha": "2026-06-02", "serie": "30EC02BE", "tipo": "FACT", "cliente": "CF", "base": 816.96, "iva": 98.04, "total": 915.0, "estado": "ACTIVA", "cobros": [{"id": "rr52x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1816504941", "monto": 915.0}]}, {"id": "v53j26", "fecha": "2026-06-02", "serie": "B01FD61B", "tipo": "FACT", "cliente": "CF", "base": 1964.29, "iva": 235.71, "total": 2200.0, "estado": "ACTIVA", "cobros": [{"id": "rr53x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1816504941", "monto": 2200.0}]}, {"id": "v54j26", "fecha": "2026-06-02", "serie": "2BB26C8B", "tipo": "FACT", "cliente": "CF", "base": 415.18, "iva": 49.82, "total": 465.0, "estado": "ACTIVA", "cobros": [{"id": "rr54x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1837379343", "monto": 465.0}]}, {"id": "v55j26", "fecha": "2026-06-02", "serie": "514F71C4", "tipo": "FACT", "cliente": "68723296", "base": 2589.29, "iva": 310.71, "total": 2900.0, "estado": "ACTIVA", "cobros": [{"id": "bi55x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "164937", "monto": 2900.0}]}, {"id": "v56j26", "fecha": "2026-06-02", "serie": "271C48E7", "tipo": "FACT", "cliente": "CF", "base": 218.75, "iva": 26.25, "total": 245.0, "estado": "ACTIVA", "cobros": [{"id": "rr56x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1797726191", "monto": 245.0}]}],
  compras: [{"id": "c2j26", "fecha": "2026-06-25", "serie": "774BDFF3", "tipo": "FPEQ", "proveedor": "JULIO JOSE , PALACIOS CASTRO", "cta": "6.1.18", "base": 5000.0, "iva": 0, "total": 5000.0, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c3j26", "fecha": "2026-06-25", "serie": "D3B904BC", "tipo": "FPEQ", "proveedor": "JESSICA MADELINE , GARCÍA MARTÍNEZ DE PALACIO", "cta": "6.1.18", "base": 5000.0, "iva": 0, "total": 5000.0, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c4j26", "fecha": "2026-06-25", "serie": "184474D3", "tipo": "FESP", "proveedor": "CORPORACIÓN NEXO GLOBAL, SOCIEDAD ANÓNIMA", "cta": "6.1.10", "base": 15446.43, "iva": 1853.57, "total": 17300.0, "isrFE": 772.32, "netoProv": 14674.11, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c5j26", "fecha": "2026-06-23", "serie": "7B2A8EA4", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 55.41, "iva": 6.65, "total": 62.06, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c6j26", "fecha": "2026-06-23", "serie": "D5EAF8CD", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 59.61, "iva": 7.15, "total": 66.76, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c7j26", "fecha": "2026-06-23", "serie": "77CD80EE", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 94.52, "iva": 11.34, "total": 105.86, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c8j26", "fecha": "2026-06-21", "serie": "1AF8BA9D", "tipo": "FACT", "proveedor": "DISTRIBUIDORA DE VEHICULOS IMPORTADOS SOCIEDA", "cta": "6.1.10", "base": 1619.65, "iva": 194.36, "total": 1814.01, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c9j26", "fecha": "2026-06-19", "serie": "F022E098", "tipo": "FACT", "proveedor": "UNO GUATEMALA, SOCIEDAD ANONIMA", "cta": "6.1.10", "base": 376.78, "iva": 45.21, "total": 487.99, "impPet": 66.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c10j26", "fecha": "2026-06-19", "serie": "F463A670", "tipo": "FACT", "proveedor": "KEILA PAOLA LOPEZ ORELLANA", "cta": "6.1.14", "base": 959.82, "iva": 115.18, "total": 1075.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c11j26", "fecha": "2026-06-16", "serie": "0D3AA43C", "tipo": "FACT", "proveedor": "POSFILE SOCIEDAD ANONIMA", "cta": "6.1.17", "base": 234.97, "iva": 28.2, "total": 263.17, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c12j26", "fecha": "2026-06-15", "serie": "ED7D7111", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 175.0, "iva": 21.0, "total": 196.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c13j26", "fecha": "2026-06-15", "serie": "59CCE986", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 50.64, "iva": 6.08, "total": 56.72, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c14j26", "fecha": "2026-06-09", "serie": "C0BC36FF", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 58.22, "iva": 6.99, "total": 65.21, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c15j26", "fecha": "2026-06-09", "serie": "41961CDA", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 60.55, "iva": 7.27, "total": 67.82, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c16j26", "fecha": "2026-06-06", "serie": "93949C02", "tipo": "FACT", "proveedor": "UNO GUATEMALA, SOCIEDAD ANONIMA", "cta": "6.1.10", "base": 307.76, "iva": 36.93, "total": 398.32, "impPet": 53.63, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c17j26", "fecha": "2026-06-06", "serie": "BE52BA5D", "tipo": "FACT", "proveedor": "PRICESMART (GUATEMALA) SOCIEDAD ANONIMA", "cta": "6.1.14", "base": 89.24, "iva": 10.71, "total": 99.95, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c18j26", "fecha": "2026-06-05", "serie": "4C9B028C", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 50.91, "iva": 6.11, "total": 57.02, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c19j26", "fecha": "2026-06-05", "serie": "78DA4639", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 50.44, "iva": 6.05, "total": 56.49, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c20j26", "fecha": "2026-06-05", "serie": "C4736B3A", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 131.19, "iva": 15.74, "total": 146.93, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c21j26", "fecha": "2026-06-04", "serie": "A003B67E", "tipo": "FACT", "proveedor": "CARGO EXPRESO, SOCIEDAD ANONIMA", "cta": "6.2.06", "base": 892.86, "iva": 107.14, "total": 1000.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c22j26", "fecha": "2026-06-03", "serie": "E484ED92", "tipo": "FACT", "proveedor": "TATMON, SOCIEDAD ANONIMA", "cta": "6.1.17", "base": 266.96, "iva": 32.04, "total": 299.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c23j26", "fecha": "2026-06-02", "serie": "CC6AF47C", "tipo": "FACT", "proveedor": "CARGO EXPRESO, SOCIEDAD ANONIMA", "cta": "6.2.06", "base": 204.29, "iva": 24.51, "total": 228.8, "impPet": 0.0, "pagoCta": "2.1.01", "estado": "ACTIVA", "registrado": false}, {"id": "c24j26", "fecha": "2026-06-01", "serie": "0695A96E", "tipo": "FACT", "proveedor": "SERVICIOS INNOVADORES DE COMUNICACIÓN Y ENTRE", "cta": "6.1.17", "base": 178.57, "iva": 21.43, "total": 200.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}]
};

// ── CATÁLOGO ───────────────────────────────────────────────────
// ── CATÁLOGO DE PROVEEDORES ───────────────────────────────────
const PROV_CAT_DEFAULT = {
  "HETERIA":"6.2.06","CARGO EXPRESO":"6.2.06","UNO GUATEMALA":"6.2.04",
  "AINNOVA":"6.2.04","SERCHI":"6.2.04","POSFILE":"6.1.17","TATMON":"6.1.17",
  "DISTRIBUIDORA DE VEH":"6.1.10","PRICESMART":"6.1.14","KEILA":"6.1.14",
  "NOVEX":"6.1.15","JESSICA GARCIA":"6.1.18","JULIO PALACIOS":"6.1.18",
  "SERVICIOS INNOV":"6.1.13","CLARO":"6.1.13","TIGO":"6.1.13",
  "COMUNICACIONES CELULARES":"6.1.13","VIVIAN":"6.1.10","CPMDG":"6.2.02",
};
function ctaDesdeProveedor(prov,provCat={}){
  const pu=(prov||"").toUpperCase();
  for(const[k,v] of Object.entries(provCat)) if(pu.includes(k.toUpperCase())) return v;
  for(const[k,v] of Object.entries(PROV_CAT_DEFAULT)) if(pu.includes(k.toUpperCase())) return v;
  return "6.1.14";
}

const CUENTAS = {
  "1.1.02":"Caja Chica","1.1.03":"Cuentas por Cobrar",
  "1.1.04":"Inventario","1.1.05":"Anticipos a Proveedores",
  "1.1.06":"IVA Credito Fiscal","1.1.08":"IVA Retenido x Cobrar",
  "1.1.17.004":"Banco Industrial Q","1.1.17.005":"Banrural Q",
  "1.1.17.006":"Banco Industrial USD","1.2.01":"Mobiliario y Equipo",
  "2.1.01":"Cuentas por Pagar","2.1.02":"Prestamos por Pagar",
  "2.1.03":"IVA Debito Fiscal","2.1.04":"Anticipos de Clientes",
  "2.1.05":"Depositos x Identificar","2.1.06":"ISR Retenido x Pagar",
  "3.1.01":"Aportacion de Socios","3.1.02":"Capital Pagado",
  "4.1.01":"Ventas de Bienes","5.1.01":"Costo de Ventas",
  "6.1.10":"Mantenimiento y Arrend. Vehiculos",
  "6.1.13":"Telefonia e Internet",
  "6.1.14":"Papeleria y Utiles de Oficina",
  "6.1.15":"Servicios Contables",
  "6.1.16":"Servicios Legales y Notariales",
  "6.1.17":"Mantenimiento y Reparaciones",
  "6.1.18":"Servicios Tecnicos",
  "6.1.19":"Gastos de Afiliacion",
  "6.2.02":"Comisiones sobre Ventas",
  "6.2.03":"Viaticos y Gastos de Viaje",
  "6.2.04":"Combustible y Lubricantes",
  "6.2.06":"Servicio de Paqueteria",
  "6.2.05":"Mantenimiento y Arrend. Vehiculos",
  "6.3.01":"Ajuste de Inventario",
  "6.4.01":"Impuesto Sobre la Renta",
};

const BANCOS_PAGO = [
  {cta:"1.1.17.004",nom:"Banco Industrial Q"},
  {cta:"1.1.17.005",nom:"Banrural Q"},
  {cta:"1.1.02",nom:"Caja Chica"},
  {cta:"2.1.01",nom:"Credito / CxP"},
  {cta:"1.1.06",nom:"IVA CF (reclasificar)"},
];

const MEDIOS_PAGO = [
  {cta:"1.1.17.004",nom:"Banco Industrial Q",icon:"🏦",tipo:"banco"},
  {cta:"1.1.17.005",nom:"Banrural Q",icon:"🏦",tipo:"banco"},
  {cta:"1.1.17.006",nom:"BI USD",icon:"💵",tipo:"banco"},
  {cta:"1.1.02",nom:"Caja / Efectivo",icon:"💵",tipo:"efectivo"},
  {cta:"VISA",nom:"Tarjeta VISA/Debito",icon:"💳",tipo:"visa"},
  {cta:"2.1.05",nom:"Dep. x Identificar",icon:"❓",tipo:"deposito"},
  {cta:"2.1.01",nom:"Credito / CxP",icon:"📋",tipo:"credito"},
];

const GASTOS_CAT = [
  // 6.1.xx Gastos Administrativos
  {cta:"6.1.10",nom:"Mantenimiento y Arrend. Vehiculos",iva:"normal"},
  {cta:"6.1.13",nom:"Telefonia e Internet",iva:"normal"},
  {cta:"6.1.14",nom:"Papeleria y Utiles de Oficina",iva:"normal"},
  {cta:"6.1.15",nom:"Servicios Contables",iva:"normal"},
  {cta:"6.1.16",nom:"Servicios Legales y Notariales",iva:"normal"},
  {cta:"6.1.17",nom:"Mantenimiento y Reparaciones",iva:"normal"},
  {cta:"6.1.18",nom:"Servicios Tecnicos",iva:"fpeq"},
  {cta:"6.1.19",nom:"Gastos de Afiliacion",iva:"normal"},
  // 6.2.xx Gastos de Venta
  {cta:"6.2.02",nom:"Comisiones sobre Ventas",iva:"normal"},
  {cta:"6.2.03",nom:"Viaticos y Gastos de Viaje",iva:"normal"},
  {cta:"6.2.04",nom:"Combustible y Lubricantes",iva:"petro"},
  {cta:"6.2.06",nom:"Servicio de Paqueteria",iva:"normal"},
  // Otros
  {cta:"6.3.01",nom:"Ajuste de Inventario",iva:"fpeq"},
  {cta:"6.4.01",nom:"Impuesto Sobre la Renta",iva:"fpeq"},
  // Activos (compras de equipo)
  {cta:"1.2.01",nom:"Mobiliario y Equipo",iva:"normal"},
];

const SEM_RANGES=[
  {label:"Semana 1",d1:1,d2:9},
  {label:"Semana 2",d1:10,d2:16},
  {label:"Semana 3",d1:17,d2:23},
  {label:"Semana 4",d1:24,d2:31},
];

const q=v=>`Q${Number(v||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const r2=v=>Math.round(Number(v||0)*100)/100;
const uid=()=>Math.random().toString(36).slice(2,9);

// ── COBROS HELPERS ─────────────────────────────────────────────
function calcTotalCobros(cobros){
  return r2(cobros.reduce((s,c)=>{
    if(c.tipo==="visa") return s+r2(Number(c.visaNet||0))+r2(Number(c.visaRet||0))+r2(Number(c.visaCom||0))+r2(Number(c.visaMem||0));
    return s+r2(Number(c.monto||0));
  },0));
}

function cobrosToLineas(cobros,cliente,serie){
  const L=[];
  cobros.forEach(c=>{
    const m=r2(Number(c.monto||0));
    if(c.tipo==="banco"||c.tipo==="efectivo"){
      if(m>0) L.push({cta:c.cta,debe:m,haber:0,conc:`${CUENTAS[c.cta]||c.cta}${c.doc?` doc:${c.doc}`:""} - ${cliente}`});
    }
    if(c.tipo==="visa"){
      const vn=r2(Number(c.visaNet||0)),vr=r2(Number(c.visaRet||0));
      const vc=r2(Number(c.visaCom||0)),vm=r2(Number(c.visaMem||0));
      if(vn>0) L.push({cta:"1.1.17.004",debe:vn,haber:0,conc:`VISA neto dep.${c.doc?` ${c.doc}`:""} - ${serie||cliente}`});
      if(vr>0) L.push({cta:"1.1.08",debe:vr,haber:0,conc:"IVA retenido VISA 15%"});
      if(vc>0) L.push({cta:"6.2.02",debe:vc,haber:0,conc:"Comision VISA bruta"});
      if(vm>0) L.push({cta:"6.1.19",debe:vm,haber:0,conc:"Membresia VISA"});
    }
    if(c.tipo==="deposito"){
      if(m>0) L.push({cta:"2.1.05",debe:m,haber:0,conc:`Dep. x identificar aplicado${c.doc?` doc:${c.doc}`:""}`});
      if(m<0) L.push({cta:"2.1.05",debe:0,haber:r2(-m),conc:`Exceso deposito${c.doc?` doc:${c.doc}`:""}`});
    }
    if(c.tipo==="credito"){
      if(m>0) L.push({cta:"2.1.01",debe:m,haber:0,conc:`Cobro CxC - ${cliente}`});
    }
  });
  return L;
}

// ── GENERADORES ────────────────────────────────────────────────

// Partida semanal consolidada de GASTOS — una sola partida por semana
function genPartidaGastosSem(gastos, num, semLabel){
  if(!gastos.length) return null;
  const fecha = gastos[0].fecha;
  const L = [];
  const isFesp = g => g.tipo==="FESP";

    // Separar gastos por tipo
  const fespGastos = gastos.filter(isFesp);
  const normGastos = gastos.filter(g=>!isFesp(g));

  // DEBE: gastos normales y FPEQ agrupados por cuenta
  const porCta = {};
  normGastos.forEach(g=>{
    if(!porCta[g.cta]) porCta[g.cta]={base:0};
    porCta[g.cta].base=r2(porCta[g.cta].base+g.base);
  });
  // FESP: usar base s/IVA (total/1.12)
  fespGastos.forEach(g=>{
    const baseFesp=g.base; // ya está calculado como total/1.12
    if(!porCta[g.cta]) porCta[g.cta]={base:0};
    porCta[g.cta].base=r2(porCta[g.cta].base+baseFesp);
  });
  Object.entries(porCta).sort().forEach(([cta,d])=>{
    L.push({cta,debe:d.base,haber:0,conc:`${CUENTAS[cta]||cta} — ${semLabel}`});
  });

  // IVA CF: solo de normales + FESP (no FPEQ)
  const tCF=r2(normGastos.reduce((a,g)=>a+g.iva,0)+fespGastos.reduce((a,g)=>a+g.iva,0));
  if(tCF>0) L.push({cta:"1.1.06",debe:tCF,haber:0,conc:`IVA CF — ${semLabel}`});

  // ISR FESP: 5% retenido del proveedor — va solo al HABER (pasivo), NO al DEBE
  // No es gasto de Nexo Global — se retiene del pago al proveedor y se paga a SAT
  const tISR=r2(fespGastos.reduce((a,g)=>a+r2(g.isrFE||r2(g.base*0.05)),0));

  // HABER banco: normales pagan base+iva, FESP paga neto (base-ISR)
  const porPago={};
  normGastos.forEach(g=>{
    const k=g.pagoCta||"1.1.02";
    porPago[k]=r2((porPago[k]||0)+g.base+g.iva);
  });
  fespGastos.forEach(g=>{
    const k=g.pagoCta||"1.1.17.004";
    const neto=g.netoProv||r2(g.base-r2(g.isrFE||r2(g.base*0.05)));
    porPago[k]=r2((porPago[k]||0)+neto);
  });
  Object.entries(porPago).forEach(([cta,tot])=>{
    L.push({cta,debe:0,haber:tot,conc:`Pago gastos ${semLabel}`});
  });

  // Pasivos FESP: IVA DT y ISR por pagar a SAT
  const tIVAFE=r2(fespGastos.reduce((a,g)=>a+g.iva,0));
  if(tIVAFE>0) L.push({cta:"2.1.03",debe:0,haber:tIVAFE,conc:`IVA DT ret. FE — ${semLabel}`});
  if(tISR>0)   L.push({cta:"2.1.06",debe:0,haber:tISR,conc:`ISR Retenido FE x Pagar — ${semLabel}`});

  const concepto = `Gastos ${semLabel} (${gastos.length} facturas)`;
  return {num, fecha, tipo:"GASTO_SEM", concepto, lineas:L,
    detalle:gastos.map(g=>({prov:g.proveedor,tipo:g.tipo,total:r2(g.base+g.iva)}))};
}

function genPartidaVenta(v,num){
  const L=cobrosToLineas(v.cobros||[],v.cliente,v.serie);
  L.push({cta:"4.1.01",debe:0,haber:r2(v.base),conc:`Venta ${v.serie||""} - ${v.cliente}`});
  L.push({cta:"2.1.03",debe:0,haber:r2(v.iva),conc:"IVA Debito Fiscal 12%"});
  return{num,fecha:v.fecha,tipo:"VENTA",concepto:`Venta ${v.serie||""} - ${v.cliente}`,lineas:L};
}

function genPartidaSemanal(vtas,num,semLabel){
  if(!vtas.length) return null;
  const allC=vtas.flatMap(v=>v.cobros||[]);
  const L=cobrosToLineas(allC,`${semLabel} (${vtas.length} facturas)`,"");
  const tB=r2(vtas.reduce((a,v)=>a+v.base,0));
  const tI=r2(vtas.reduce((a,v)=>a+v.iva,0));
  L.push({cta:"4.1.01",debe:0,haber:tB,conc:`Ventas ${semLabel} - ${vtas.length} facturas`});
  L.push({cta:"2.1.03",debe:0,haber:tI,conc:`IVA Debito Fiscal 12% - ${semLabel}`});
  return{num,fecha:vtas[0].fecha,tipo:"VENTA_SEM",concepto:`Ventas ${semLabel}`,lineas:L,
    detalle:vtas.map(v=>({serie:v.serie,cliente:v.cliente,total:r2(v.base+v.iva)}))};
}

function genPartidaGasto(c,num){
  const L=[];
  const isFesp=c.tipo==="FESP";

  if(isFesp){
    // FESP: base=total/1.12, ISR=5% base (retenido del proveedor, no gasto de Nexo)
    const isr=c.isrFE||r2(c.base*0.05);
    const neto=c.netoProv||r2(c.base-isr); // lo que se paga al proveedor
    L.push({cta:c.cta,debe:r2(c.base),haber:0,conc:`${CUENTAS[c.cta]||c.cta} - ${c.proveedor}`});
    if(c.iva>0) L.push({cta:"1.1.06",debe:r2(c.iva),haber:0,conc:`IVA CF FE - ${c.proveedor}`});
    L.push({cta:c.pagoCta||"1.1.17.004",debe:0,haber:neto,conc:`Pago neto proveedor - ${c.proveedor}`});
    if(c.iva>0) L.push({cta:"2.1.03",debe:0,haber:r2(c.iva),conc:`IVA DT ret. FE x Pagar`});
    if(isr>0)   L.push({cta:"2.1.06",debe:0,haber:isr,conc:`ISR Retenido FE x Pagar (5%)`});
  } else {
    if(c.base>0) L.push({cta:c.cta,debe:r2(c.base),haber:0,conc:`${CUENTAS[c.cta]||c.cta} - ${c.proveedor}`});
    if(c.iva>0)  L.push({cta:"1.1.06",debe:r2(c.iva),haber:0,conc:`IVA CF - ${c.proveedor}`});
    L.push({cta:c.pagoCta||"1.1.02",debe:0,haber:r2(c.base+c.iva),conc:`Pago - ${c.proveedor}`});
  }
  return{num,fecha:c.fecha,tipo:isFesp?"FESP":"COMPRA",
    concepto:`${CUENTAS[c.cta]||c.cta} - ${c.proveedor}`,lineas:L};
}

// ── PARSER FEL ─────────────────────────────────────────────────
function parseFecha(f){
  if(!f) return "";
  const s=String(f);
  const m1=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m1) return `${m1[1]}-${m1[2]}-${m1[3]}`;
  const m2=s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if(m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;
  if(typeof f==="number"){
    try{const d=XLSX.SSF.parse_date_code(f);
      return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;}catch(e){}
  }
  return s.slice(0,10);
}

// Parser para formato ClickToolsGT / EMITIDAS — múltiples columnas de depósito
function parseFELVentas(ws){
  const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  // Detectar si es formato estándar SAT (encabezado en fila 5) o formato ClickTools (fila 1)
  const f0=String(data[0]?.[0]||"").toLowerCase();
  const esClickTools=f0.includes("fecha")||f0.includes("emisi");
  const inicio=esClickTools?1:4;

  const toNum=v=>{try{const n=Number(v);return isNaN(n)?0:n;}catch{return 0;}};

  return data.slice(inicio).filter(r=>{
    if(!r||!r[0]) return false;
    const gt=toNum(r[18]);
    if(!gt) return false;
    const fecha=String(r[0]);
    return fecha.includes("-")||fecha.includes("/"); // tiene fecha válida
  }).map((r,i)=>{
    const gt=toNum(r[18]),iva=toNum(r[19]);
    const estado=String(r[16]||"").toLowerCase();
    const cobros=[];

    if(esClickTools){
      // Formato ClickTools: BI en cols 23,26,29,32 (idx 22,25,28,31) | RR en 35,38,41,44,47 (idx 34,37,40,43,46)
      // VISA: neto=col53(52), ret=col58(57), com=col55+57(54+56), mem=col52(51)
      // "PENDIENTE DEPOSITO" → registrar como CxC (1.1.03)
      let tienePendiente=false;
      [[22,23],[25,26],[28,29],[31,32]].forEach(([mi,ri])=>{
        const raw=r[mi]; const m=toNum(raw);
        if(m>0) cobros.push({id:uid(),tipo:"banco",cta:"1.1.17.004",
          doc:String(r[ri]||"").slice(0,12),monto:m});
        else if(raw&&String(raw).toUpperCase().includes("PENDIENTE")) tienePendiente=true;
      });
      if(tienePendiente) cobros.push({id:uid(),tipo:"banco",cta:"1.1.03",
        doc:"PENDIENTE",monto:toNum(r[18])});
      [[34,35],[37,38],[40,41],[43,44],[46,47]].forEach(([mi,ri])=>{
        const m=toNum(r[mi]);
        if(m>0) cobros.push({id:uid(),tipo:"banco",cta:"1.1.17.005",
          doc:String(r[ri]||"").slice(0,12),monto:m});
      });
      const vDep=toNum(r[52]),vRet=toNum(r[57]),vCom=r2(toNum(r[54])+toNum(r[56])),vMem=toNum(r[51]);
      if(vDep>0) cobros.push({id:uid(),tipo:"visa",cta:"VISA",doc:String(r[53]||"").slice(0,12),
        visaNet:vDep,visaRet:vRet,visaCom:vCom,visaMem:vMem});
    } else {
      // Formato SAT estándar
      const bi=toNum(r[33]),br=toNum(r[45]);
      const vn=toNum(r[66]),vm=toNum(r[65]),vc=r2(toNum(r[68])+toNum(r[70])),vr=toNum(r[71]);
      if(bi>0) cobros.push({id:uid(),tipo:"banco",cta:"1.1.17.004",doc:"",monto:bi});
      if(br>0) cobros.push({id:uid(),tipo:"banco",cta:"1.1.17.005",doc:"",monto:br});
      if(vn>0) cobros.push({id:uid(),tipo:"visa",cta:"VISA",doc:"",visaNet:vn,visaRet:vr,visaCom:vc,visaMem:vm});
    }
    return{id:`v${i}${Date.now()}`,fecha:parseFecha(r[0]),
      serie:String(r[3]||"").slice(0,8).toUpperCase(),tipo:String(r[2]||"FACT"),
      cliente:String(r[13]||"CF").slice(0,45),base:r2(gt-iva),iva:r2(iva),total:r2(gt),
      estado:estado.includes("anul")?"ANULADO":"ACTIVA",cobros};
  });
}

// Parser para formato ClickToolsGT / recibidas — con columna de método de pago en texto
function parseFELCompras(ws){
  const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  const f0=String(data[0]?.[0]||"").toLowerCase();
  const esClickTools=f0.includes("fecha")||f0.includes("emisi");
  const inicio=esClickTools?1:3;
  const toNum=v=>{try{const n=Number(v);return isNaN(n)?0:n;}catch{return 0;}};

  // Mapa método de pago texto → cuenta contable
  const pagoACta=(p)=>{
    const s=String(p||"").toLowerCase();
    if(s.includes("industrial")||s.includes("bi ")) return "1.1.17.004";
    if(s.includes("banrural")||s.includes("rr ")) return "1.1.17.005";
    if(s.includes("caja")) return "1.1.02";
    if(s.includes("pagar")||s.includes("cxp")||s.includes("credito")) return "2.1.01";
    return "1.1.02"; // default caja chica
  };

  return data.slice(inicio).filter(r=>{
    if(!r||!r[0]) return false;
    const gt=toNum(r[18]); if(!gt) return false;
    const fecha=String(r[0]);
    return fecha.includes("-")||fecha.includes("/");
  }).map((r,i)=>{
    const toNum2=v=>{try{const n=Number(v);return isNaN(n)?0:n;}catch{return 0;}};
    const gt=toNum2(r[18]),iva=toNum2(r[19]),ip=toNum2(r[22]);
    const prov=String(r[9]||"").slice(0,45),tipo=String(r[2]||"FACT");
    const pu=prov.toUpperCase();
    // Determinar cuenta de gasto por proveedor
    let cta="6.1.14";
    if(pu.includes("UNO")||pu.includes("PUMA"))cta="6.1.10";
    else if(pu.includes("JESSICA")||pu.includes("GARCIA")||pu.includes("JULIO")||pu.includes("PALACIOS"))cta="6.1.18";
    else if(pu.includes("POSFILE")||pu.includes("TATMON")||pu.includes("SERVICIOS INNOVAD"))cta="6.1.17";
    else if(pu.includes("CARGO EXPRESO")||pu.includes("HETERIA"))cta="6.2.06";
    else if(pu.includes("NOVEX"))cta="6.1.15";
    else if(pu.includes("PROCESAMIENTO")||pu.includes("CPMDG"))cta="6.2.02";
    else if(pu.includes("CLARO")||pu.includes("TIGO"))cta="6.1.13";
    else if(pu.includes("VIVIAN"))cta="6.2.05";
    else if(pu.includes("DISTRIBUIDORA DE VEHICULOS")||pu.includes("PRICESMART"))cta="6.1.14";
    else if(pu.includes("KEILA")||pu.includes("LOPEZ"))cta="6.1.14";
    const cat=GASTOS_CAT.find(g=>g.cta===cta);
    let base,ivaC;
    if(cat?.iva==="fpeq"||tipo==="FPEQ"){base=r2(gt);ivaC=0;}
    else if(tipo==="FESP"){base=r2(gt);ivaC=r2(gt*0.12);} // Factura Especial (IVA se agrega aparte)
    else if(cat?.iva==="petro"&&ip>0){const np=r2(gt-ip);base=r2(np/1.12);ivaC=r2(np/1.12*0.12);}
    else{base=r2(gt/1.12);ivaC=r2(gt/1.12*0.12);}
    // Método de pago desde columna 24 (texto) si existe, sino inferir de proveedor
    const pagoTexto=esClickTools?String(r[23]||""):"";
    const pagoCta=pagoTexto?pagoACta(pagoTexto):"1.1.02";
    const estado2=String(r[16]||"").toLowerCase();
    return{id:`c${i}${Date.now()}`,fecha:parseFecha(r[0]),
      serie:String(r[3]||"").slice(0,8).toUpperCase(),tipo,proveedor:prov,
      cta,base,iva:ivaC,total:r2(gt),impPet:ip,pagoCta,
      estado:estado2.includes("anul")?"ANULADO":"ACTIVA"};
  });
}

// ── PERSISTENCIA — Google Sheets + Storage local ──────────────
const GAS_URL="https://script.google.com/macros/s/AKfycbyQlGGElFc20CwdqgB4978ZfBF53UbxQTcIl2n3A2Gafe8pYSQ9fJWJETYIdO7FRVIi/exec";
// En Vercel: usar proxy /api/sheets para evitar CORS
const API_BASE = typeof window!=="undefined" && window.location?.hostname!=="localhost" 
  && !window.location?.hostname?.includes("claude") ? "" : "";
const SHEETS_ENDPOINT = "/api/sheets";

// Storage local (respaldo temporal)
async function guardar(k,d){try{await window.storage.set(k,JSON.stringify(d));}catch(e){}}
async function cargar(k){try{const r=await window.storage.get(k);return r?JSON.parse(r.value):null;}catch(e){return null;}}

// Google Sheets (permanente) — usando no-cors para POST y cors para GET
async function guardarGS(tabla,data){
  try{
    // Usamos no-cors para evitar bloqueo CORS en POST
    // Los datos se guardan aunque no podamos leer la respuesta
    const body=JSON.stringify({tabla,datos:[[JSON.stringify(data)]]});
    await fetch(GAS_URL,{
      method:"POST",
      mode:"no-cors",
      headers:{"Content-Type":"text/plain"}, // text/plain evita preflight CORS
      body
    });
    console.log("GS guardado:",tabla,data.length||"config");
  }catch(e){console.warn("GS guardar error:",e);}
}

async function cargarGS(tabla){
  try{
    // GET no tiene problema de CORS
    const r=await fetch(`${GAS_URL}?tabla=${tabla}`,{mode:"cors"});
    if(!r.ok) return null;
    const rows=await r.json();
    if(rows&&rows[0]&&rows[0][0]) return JSON.parse(rows[0][0]);
    return null;
  }catch(e){
    // Si falla CORS en GET, intentar con jsonp-like approach
    console.warn("GS cargar error:",e);
    return null;
  }
}

// ── EXPORT — FORMATO IGUAL AL EXCEL EXISTENTE ─────────────────
// Columnas: N° | FECHA | CUENTA | DESCRIPCION | CONCEPTO | PARCIAL | DEBE | HABER
// ── ESTADOS FINANCIEROS DESDE PARTIDAS — VERSIÓN DETALLADA ──────
function calcEF(partidas, periodos=null){
  const acc={}, accAll={};
  partidas.forEach(p=>p.lineas.forEach(l=>{
    if(!accAll[l.cta]) accAll[l.cta]={d:0,h:0};
    accAll[l.cta].d=r2(accAll[l.cta].d+l.debe);
    accAll[l.cta].h=r2(accAll[l.cta].h+l.haber);
    if(periodos){const f=(p.fecha||"").slice(0,7);if(f<periodos.desde||f>periodos.hasta)return;}
    if(!acc[l.cta]) acc[l.cta]={d:0,h:0};
    acc[l.cta].d=r2(acc[l.cta].d+l.debe);
    acc[l.cta].h=r2(acc[l.cta].h+l.haber);
  }));
  const h=(c)=>r2(acc[c]?.h||0), d=(c)=>r2(acc[c]?.d||0), sal=(c)=>r2((accAll[c]?.d||0)-(accAll[c]?.h||0));

  // ── P&L con NCREs ────────────────────────────────────────
  const ventas_brutas=h("4.1.01"), ncre_emitidas=d("4.1.01"), ventas_netas=r2(ventas_brutas-ncre_emitidas);

  // ── CV desglose desde inventario ─────────────────────────
  const inv_mov=partidas.flatMap(p=>p.lineas.filter(l=>l.cta==="1.1.04").map(l=>({...l,fecha:p.fecha})));
  const en_per=(f)=>!periodos||(((f||"").slice(0,7)>=periodos.desde)&&((f||"").slice(0,7)<=periodos.hasta));
  const antes_per=(f)=>periodos&&(f||"").slice(0,7)<periodos.desde;

  const inv_ini_saldo=r2(inv_mov.filter(l=>antes_per(l.fecha)).reduce((a,l)=>a+l.debe-l.haber,0));
  const inv_ini=periodos?inv_ini_saldo:r2(sal("1.1.04")+d("5.1.01"));
  const inv_per=inv_mov.filter(l=>en_per(l.fecha));
  const c_local=r2(inv_per.filter(l=>l.debe>0&&!["IMPORT","LIQUIDAC"].some(k=>(l.conc||"").toUpperCase().includes(k))).reduce((a,l)=>a+l.debe,0));
  const c_ext=r2(inv_per.filter(l=>l.debe>0&&["IMPORT","LIQUIDAC"].some(k=>(l.conc||"").toUpperCase().includes(k))).reduce((a,l)=>a+l.debe,0));
  const ncre_recibidas_inv=r2(inv_per.filter(l=>l.haber>0&&["DEVOL","NCRE"].some(k=>(l.conc||"").toUpperCase().includes(k))).reduce((a,l)=>a+l.haber,0));

  const cv_registrado=d("5.1.01"), inv_final=sal("1.1.04");
  const ub=r2(ventas_netas-cv_registrado);

  // ── Gastos con NCREs recibidas ───────────────────────────
  const gastos_ctas=[
    {cta:"6.1.10",nom:"Mantenimiento y Arrend. Vehiculos"},
    {cta:"6.1.13",nom:"Telefonia e Internet"},
    {cta:"6.1.14",nom:"Papeleria y Utiles de Oficina"},
    {cta:"6.1.15",nom:"Servicios Contables"},
    {cta:"6.1.16",nom:"Servicios Legales y Notariales"},
    {cta:"6.1.17",nom:"Mantenimiento y Reparaciones"},
    {cta:"6.1.18",nom:"Servicios Tecnicos"},
    {cta:"6.1.19",nom:"Gastos de Afiliacion"},
    {cta:"6.2.02",nom:"Comisiones sobre Ventas"},
    {cta:"6.2.03",nom:"Viaticos y Gastos de Viaje"},
    {cta:"6.2.04",nom:"Combustible y Lubricantes"},
    {cta:"6.2.06",nom:"Servicio de Paqueteria"},
    {cta:"6.3.01",nom:"Ajuste de Inventario"},
  ];
  const gastos=gastos_ctas.map(({cta,nom})=>({cta,nom,debe:d(cta),haber:h(cta),neto:r2(d(cta)-h(cta))}));
  const tot_gas=r2(gastos.reduce((a,g)=>a+g.neto,0));
  const ncre_en_gastos=r2(gastos.reduce((a,g)=>a+g.haber,0));

  const uo=r2(ub-tot_gas), isr_pag=r2(d("6.4.01")-h("6.4.01"));
  const isr_dev=r2(Math.max(uo,0)*0.25), reserva=r2(Math.max(uo,0)*0.05);
  const un=r2(uo-isr_pag-reserva);

  // ── Balance ──────────────────────────────────────────────
  const bancos=r2(sal("1.1.17.004")+sal("1.1.17.005")+sal("1.1.17.006")+sal("1.1.02"));
  const act_c=r2(bancos+sal("1.1.03")+sal("1.1.04")+sal("1.1.05")+sal("1.1.06")+sal("1.1.08")+sal("1.1.09"));
  const act_f=r2(sal("1.2.01")), tot_act=r2(act_c+act_f);
  const cxp=r2(-sal("2.1.01")), prest=r2(-sal("2.1.02")), iva_dt=r2(-sal("2.1.03"));
  const ant_c=r2(-sal("2.1.04")), dep_xi=r2(-sal("2.1.05")), isr_r=r2(-sal("2.1.06"));
  const tot_pas=r2(cxp+prest+iva_dt+ant_c+dep_xi+isr_r);
  const capital=r2(-sal("3.1.01")-sal("3.1.02"));
  const un_bal=r2(uo-isr_pag), tot_cap=r2(capital+un_bal), tot_pc=r2(tot_pas+tot_cap);

  return{ventas_brutas,ncre_emitidas,ventas_netas,
    inv_ini,c_local,c_ext,ncre_recibidas_inv,
    cv_registrado,inv_final,ub,gastos,ncre_en_gastos,tot_gas,
    uo,isr_pag,isr_dev,reserva,un,
    margen_bruto:ventas_netas?r2(ub/ventas_netas):0,
    margen_op:ventas_netas?r2(uo/ventas_netas):0,
    bancos,act_c,act_f,tot_act,cxp,prest,iva_dt,ant_c,dep_xi,isr_r,
    tot_pas,capital,un_bal,tot_cap,tot_pc,dif_balance:r2(tot_act-tot_pc)};
}


function calcEFDummy(partidas, periodos){return calcEF(partidas,periodos);}

// ── ESTADOS FINANCIEROS EN TIEMPO REAL ────────────────────────
function EstadosFinancierosView({partidas}){
  const[vista,setVista]=useState("pyl");
  const[showRef,setShowRef]=useState(false);
  const ef=useMemo(()=>calcEF(partidas),[partidas]);

  const Fil=({v,lbl})=>(
    <button onClick={()=>setVista(v)}
      className={`px-3 py-1.5 rounded-lg text-xs font-bold ${vista===v?"bg-blue-700 text-white":"bg-gray-100 text-gray-600"}`}>
      {lbl}
    </button>
  );

  const Row=({lbl,val,neg=false,bold,sub,pct,color,ref,indent=0})=>(
    <div className={`flex items-center px-3 py-1.5 gap-2 ${sub?"bg-blue-50 border-t border-b border-blue-100":""}`}>
      {indent>0&&<span style={{width:indent*10}} className="shrink-0"/>}
      <div className="flex-1">
        <span className={`text-xs ${sub?"text-blue-800 font-bold":bold?"font-semibold text-gray-800":"text-gray-700"} ${color||""}`}>
          {neg&&<span className="text-red-500 mr-1">(−)</span>}{lbl}
        </span>
        {showRef&&ref&&<span className="text-xs text-blue-400 ml-1 font-mono">[{ref}]</span>}
      </div>
      <div className="shrink-0 text-right">
        <span className={`text-xs font-mono ${sub?"text-blue-800 font-bold":bold?"font-semibold":""} ${color||""} ${val<0?"text-red-600":""}`}>
          {val===null||val===undefined?"":Math.abs(val)<0.01?"Q 0.00":val<0?`(${q(Math.abs(val))})`:q(val)}
        </span>
        {pct!==undefined&&Math.abs(pct)>0.001&&<span className="text-xs text-gray-400 ml-1">{(pct*100).toFixed(1)}%</span>}
      </div>
    </div>
  );
  const Sec=({lbl})=><div className="bg-gray-800 text-white px-3 py-1.5 text-xs font-bold uppercase tracking-wide mt-1">{lbl}</div>;
  const Div=()=><div className="border-t border-gray-100 my-0.5"/>;

  return(
    <div className="space-y-3">
      <div className="flex gap-2 flex-wrap items-center">
        <Fil v="pyl" lbl="📊 Resultados"/>
        <Fil v="balance" lbl="⚖️ Balance"/>
        <Fil v="isr" lbl="🏦 ISR"/>
        <button onClick={()=>setShowRef(p=>!p)}
          className={`px-3 py-1.5 rounded-lg text-xs font-bold ml-auto border ${showRef?"bg-amber-500 text-white border-amber-500":"bg-amber-50 text-amber-700 border-amber-300"}`}>
          {showRef?"✅ Ref. Mayor ON":"🔍 Mostrar ref. Mayor"}
        </button>
      </div>
      {showRef&&<div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2 text-xs text-amber-800">
        💡 Código <span className="font-mono text-blue-600">[entre corchetes]</span> = cuenta del Libro Mayor que origina ese número.
      </div>}

      {vista==="pyl"&&(
        <div className="border rounded-xl overflow-hidden">
          <div className="bg-blue-900 text-white px-3 py-2 flex justify-between">
            <span className="font-bold text-sm">ESTADO DE RESULTADOS</span>
            <span className="text-xs text-blue-300">Acumulado</span>
          </div>
          <Sec lbl="Ingresos del Período"/>
          <Row lbl="Ventas de Bienes" val={ef.ventas_brutas} ref="4.1.01 HABER" pct={1}/>
          {ef.ncre_emitidas>0.01&&<Row lbl="Devoluciones / NCRE Emitidas" val={ef.ncre_emitidas} neg ref="4.1.01 DEBE" color="text-red-600" indent={1}/>}
          <Row lbl="VENTAS NETAS" val={ef.ventas_netas} sub bold ref="4.1.01 neto"/><Div/>

          <Sec lbl="Costo de Ventas"/>
          <Row lbl="Inventario Inicial" val={ef.inv_ini} ref="1.1.04 inicio" indent={1}/>
          {ef.c_local>0.01&&<Row lbl="(+) Compras Locales" val={ef.c_local} ref="1.1.04 DEBE local" indent={2}/>}
          {ef.c_ext>0.01&&<Row lbl="(+) Importaciones" val={ef.c_ext} ref="1.1.04 DEBE import." indent={2}/>}
          {ef.ncre_recibidas_inv>0.01&&<Row lbl="(−) NCRE Recibidas Compras" val={ef.ncre_recibidas_inv} neg ref="1.1.04 HABER devol." color="text-red-600" indent={2}/>}
          <Row lbl="Mercadería Disponible" val={r2(ef.inv_ini+ef.c_local+ef.c_ext-ef.ncre_recibidas_inv)} bold indent={1}/>
          <Row lbl="(−) Inventario Final" val={ef.inv_final} neg ref="1.1.04 saldo final" color="text-red-600" indent={2}/>
          <Row lbl="COSTO DE LO VENDIDO" val={ef.cv_registrado} sub bold ref="5.1.01 DEBE"/><Div/>
          <Row lbl="UTILIDAD BRUTA EN VENTAS" val={ef.ub} sub bold pct={ef.margen_bruto}
            color={ef.ub>=0?"text-green-800":"text-red-800"}/><Div/>

          <Sec lbl="Gastos de Operación"/>
          {ef.gastos.map((g,i)=>g.neto>0.01?(
            <div key={i}>
              <Row lbl={g.nom} val={g.neto} ref={g.cta} indent={1}
                pct={ef.ventas_netas?g.neto/ef.ventas_netas:0}/>
              {g.haber>0.01&&<Row lbl={"(−) Notas de Crédito Recibidas — "+g.cta} val={g.haber}
                neg ref={g.cta+" HABER"} color="text-red-500" indent={2}/>}
            </div>
          ):null)}
          <Row lbl="TOTAL GASTOS" val={ef.tot_gas} sub bold ref="6.x.xx neto"/><Div/>

          <Row lbl="UTILIDAD DE OPERACIÓN" val={ef.uo} sub bold pct={ef.margen_op}
            color={ef.uo>=0?"text-green-800":"text-red-800"}/>
          <Row lbl="ISR pagado en libros" val={ef.isr_pag} ref="6.4.01 DEBE" neg indent={1}/>
          <Row lbl="Reserva Legal 5% (estimada)" val={ef.reserva} neg indent={1}/>
          <Row lbl="UTILIDAD NETA DEL PERÍODO" val={ef.un} sub bold
            color={ef.un>=0?"text-green-800":"text-red-800"}/>
          <div className="bg-gray-50 px-3 py-2 text-xs text-gray-500 grid grid-cols-2 gap-1">
            <span>Margen bruto: <b>{(ef.margen_bruto*100).toFixed(1)}%</b></span>
            <span>Margen op: <b>{(ef.margen_op*100).toFixed(1)}%</b></span>
            <span>ISR devengado 25%: <b>{q(ef.isr_dev)}</b></span>
            <span>ISR pagado libros: <b>{q(ef.isr_pag)}</b></span>
          </div>
        </div>
      )}

      {vista==="balance"&&(
        <div className="border rounded-xl overflow-hidden">
          <div className="bg-blue-900 text-white px-3 py-2 flex justify-between">
            <span className="font-bold text-sm">BALANCE GENERAL</span>
            <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${Math.abs(ef.dif_balance)<0.10?"bg-green-500":"bg-red-500"} text-white`}>
              {Math.abs(ef.dif_balance)<0.10?"✅ CUADRA":"⚠️ Dif Q"+ef.dif_balance.toFixed(2)}
            </span>
          </div>
          <Sec lbl="Activo"/>
          <Row lbl="Caja y Bancos" val={ef.bancos} ref="1.1.02+1.1.17.x" indent={1}/>
          <Row lbl="Inventario Final" val={ef.inv_final} ref="1.1.04" indent={1}/>
          {ef.act_c-ef.bancos-ef.inv_final>0.01&&<Row lbl="Otros Activos Corrientes" val={r2(ef.act_c-ef.bancos-ef.inv_final)} ref="1.1.03/05/06/08/09" indent={1}/>}
          <Row lbl="Total Activo Circulante" val={ef.act_c} bold/>
          {ef.act_f>0.01&&<Row lbl="Mobiliario y Equipo" val={ef.act_f} ref="1.2.01" indent={1}/>}
          <Row lbl="TOTAL ACTIVO" val={ef.tot_act} sub bold/>
          <Div/>
          <Sec lbl="Pasivo"/>
          {ef.cxp>0.01&&<Row lbl="Cuentas por Pagar" val={ef.cxp} ref="2.1.01" indent={1}/>}
          {ef.prest>0.01&&<Row lbl="Prestamos por Pagar" val={ef.prest} ref="2.1.02" indent={1}/>}
          {Math.abs(ef.iva_dt)>0.01&&<Row lbl="IVA Debito Fiscal" val={ef.iva_dt} ref="2.1.03" indent={1}/>}
          {ef.ant_c>0.01&&<Row lbl="Anticipos de Clientes" val={ef.ant_c} ref="2.1.04" indent={1}/>}
          {ef.dep_xi>0.01&&<Row lbl="Depositos x Identificar" val={ef.dep_xi} ref="2.1.05" indent={1}/>}
          {ef.isr_r>0.01&&<Row lbl="ISR Retenido x Pagar" val={ef.isr_r} ref="2.1.06" indent={1}/>}
          <Row lbl="TOTAL PASIVO" val={ef.tot_pas} bold/>
          <Div/>
          <Sec lbl="Capital y Reservas"/>
          <Row lbl="Capital Socios" val={ef.capital} ref="3.1.x" indent={1}/>
          <Row lbl="Utilidad del Período" val={ef.un_bal} ref="P&L acumulado" indent={1} color="text-blue-700"/>
          <Row lbl="TOTAL CAPITAL" val={ef.tot_cap} bold/>
          <Row lbl="TOTAL PASIVO + CAPITAL" val={ef.tot_pc} sub bold/>
        </div>
      )}
      {vista==="isr"&&<ISRDashboard partidas={partidas} ef={ef}/>}
    </div>
  );
}


// ── ISR DASHBOARD ──────────────────────────────────────────────
function ISRDashboard({partidas,ef}){
  // Calcular ISR por trimestre
  const trimestres = useMemo(()=>{
    const trim={Q1:{meses:["2026-01","2026-02","2026-03"]},
               Q2:{meses:["2026-04","2026-05","2026-06"]},
               Q3:{meses:["2026-07","2026-08","2026-09"]},
               Q4:{meses:["2026-10","2026-11","2026-12"]}};
    return Object.entries(trim).map(([k,{meses}])=>{
      const pvs=partidas.filter(p=>meses.includes((p.fecha||"").slice(0,7)));
      const ventas=r2(pvs.flatMap(p=>p.lineas).filter(l=>l.cta==="4.1.01").reduce((a,l)=>a-l.haber,0));
      const cv=r2(pvs.flatMap(p=>p.lineas).filter(l=>l.cta==="5.1.01").reduce((a,l)=>a+l.debe,0));
      const gastos=r2(pvs.flatMap(p=>p.lineas).filter(l=>l.cta.startsWith("6.")&&l.cta!=="6.4.01").reduce((a,l)=>a+l.debe,0));
      const isr_pag=r2(pvs.flatMap(p=>p.lineas).filter(l=>l.cta==="6.4.01").reduce((a,l)=>a+l.debe,0));
      const uo=r2(ventas-cv-gastos);
      const isr_dev=r2(Math.max(uo,0)*0.25);
      const pendiente=r2(isr_dev-isr_pag);
      return{k,ventas,cv,uo,isr_dev,isr_pag,pendiente,meses};
    }).filter(t=>t.ventas>0||t.isr_pag>0);
  },[partidas]);

  return(
    <div className="space-y-3">
      <div className="bg-amber-50 border border-amber-200 rounded-xl p-3">
        <p className="text-xs font-bold text-amber-800 mb-1">💡 Régimen ISR: Sobre Utilidades 25% trimestral</p>
        <p className="text-xs text-amber-600">ISR devengado = 25% × Utilidad Operativa del trimestre. Pagar los primeros 10 días del mes siguiente.</p>
      </div>
      {trimestres.map(t=>(
        <div key={t.k} className={`border rounded-xl overflow-hidden ${t.pendiente>0.01?"border-red-300":"border-green-300"}`}>
          <div className={`px-3 py-2 flex justify-between items-center ${t.pendiente>0.01?"bg-red-50":"bg-green-50"}`}>
            <span className="font-bold text-sm">{t.k} — {t.meses[0].slice(5)} a {t.meses[2].slice(5)}</span>
            <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${t.pendiente>0.01?"bg-red-600 text-white":"bg-green-600 text-white"}`}>
              {t.pendiente>0.01?`⚠️ Pendiente Q${q(t.pendiente)}`:"✅ Pagado"}
            </span>
          </div>
          <div className="p-3 grid grid-cols-2 gap-2 text-xs">
            <div className="bg-white rounded p-2">
              <div className="text-gray-400">Ventas</div>
              <div className="font-bold">{q(t.ventas)}</div>
            </div>
            <div className="bg-white rounded p-2">
              <div className="text-gray-400">Util. Operativa</div>
              <div className={`font-bold ${t.uo>=0?"text-green-700":"text-red-700"}`}>{q(t.uo)}</div>
            </div>
            <div className="bg-blue-50 rounded p-2">
              <div className="text-gray-400">ISR devengado (25%)</div>
              <div className="font-bold text-blue-700">{q(t.isr_dev)}</div>
            </div>
            <div className="bg-white rounded p-2">
              <div className="text-gray-400">ISR pagado (libros)</div>
              <div className="font-bold text-green-700">{q(t.isr_pag)}</div>
            </div>
          </div>
          {t.pendiente>0.01&&(
            <div className="bg-red-50 px-3 py-2 text-xs text-red-700 font-semibold">
              ⚠️ Pendiente de pagar: Q{q(t.pendiente)} — Registrar partida: DEBE 6.4.01 / HABER Banco
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// ── CONCILIACIÓN AUTOMÁTICA DIARIO vs MAYOR ────────────────────
function ConciliacionView({partidas}){
  const [filtro,setFiltro]=useState("");
  const conciliacion=useMemo(()=>{
    // Calcular saldo por cuenta desde partidas (Diario)
    const acc={};
    partidas.forEach(p=>p.lineas.forEach(l=>{
      if(!acc[l.cta]) acc[l.cta]={cta:l.cta,d:0,h:0,movs:0};
      acc[l.cta].d=r2(acc[l.cta].d+l.debe);
      acc[l.cta].h=r2(acc[l.cta].h+l.haber);
      acc[l.cta].movs++;
    }));
    return Object.values(acc).sort((a,b)=>a.cta.localeCompare(b.cta)).map(c=>({
      ...c,
      saldo:r2(c.d-c.h),
      nom:CUENTAS[c.cta]||c.cta,
      ok:true // El Mayor auto-generado SIEMPRE cuadra con el Diario
    }));
  },[partidas]);

  const filtradas=filtro?conciliacion.filter(c=>c.cta.includes(filtro)||c.nom.toLowerCase().includes(filtro.toLowerCase())):conciliacion;
  const totD=r2(filtradas.reduce((a,c)=>a+c.d,0));
  const totH=r2(filtradas.reduce((a,c)=>a+c.h,0));

  return(
    <div className="space-y-3">
      <div className="bg-green-50 border border-green-200 rounded-xl p-3 text-xs text-green-800">
        ✅ <b>El Libro Mayor se genera automáticamente del Diario</b> — siempre cuadran por definición. Esta vista muestra el resumen de saldos por cuenta.
      </div>
      <input placeholder="Buscar cuenta..." value={filtro} onChange={e=>setFiltro(e.target.value)}
        className="w-full border rounded-xl px-3 py-2 text-sm"/>
      <div className="border rounded-xl overflow-hidden">
        <div className="bg-gray-800 text-white px-3 py-2 grid grid-cols-6 text-xs font-bold">
          <span className="col-span-2">Cuenta</span>
          <span className="text-right">Debe</span>
          <span className="text-right">Haber</span>
          <span className="text-right">Saldo</span>
          <span className="text-right">Movs</span>
        </div>
        <div className="max-h-96 overflow-y-auto divide-y text-xs">
          {filtradas.map((c,i)=>(
            <div key={i} className={`grid grid-cols-6 px-3 py-1.5 ${i%2?"bg-white":"bg-gray-50"}`}>
              <div className="col-span-2">
                <div className="font-semibold text-blue-700">{c.cta}</div>
                <div className="text-gray-500 text-xs truncate">{c.nom}</div>
              </div>
              <span className="text-right self-center">{q(c.d)}</span>
              <span className="text-right self-center">{q(c.h)}</span>
              <span className={`text-right self-center font-bold ${c.saldo>=0?"text-gray-800":"text-red-600"}`}>
                {c.saldo>=0?q(c.saldo):`(${q(Math.abs(c.saldo))})`}
              </span>
              <span className="text-right self-center text-gray-400">{c.movs}</span>
            </div>
          ))}
        </div>
        <div className="grid grid-cols-6 px-3 py-2 bg-blue-900 text-white text-xs font-bold">
          <span className="col-span-2">TOTALES</span>
          <span className="text-right">{q(totD)}</span>
          <span className="text-right">{q(totH)}</span>
          <span className="text-right">{q(r2(totD-totH))}</span>
          <span className="text-right">{filtradas.reduce((a,c)=>a+c.movs,0)}</span>
        </div>
      </div>
    </div>
  );
}

function LibroMayorView({partidas}){
  const ctas=useMemo(()=>{const s=new Set();partidas.forEach(p=>p.lineas.forEach(l=>s.add(l.cta)));return[...s].sort();},[partidas]);
  const[sel,setSel]=useState(ctas[0]||"4.1.01");
  const movs=useMemo(()=>{let s=0;return partidas.flatMap(p=>
    p.lineas.filter(l=>l.cta===sel).map(l=>{s=r2(s+l.debe-l.haber);return{...l,pnum:p.num,fecha:p.fecha,saldo:s};}));},[partidas,sel]);
  const td=r2(movs.reduce((a,m)=>a+m.debe,0)),th=r2(movs.reduce((a,m)=>a+m.haber,0));
  return(
    <div className="space-y-3">
      <div className="flex gap-2">
        <Sel className="flex-1" value={sel} onChange={setSel}
          options={ctas.map(c=>({value:c,label:`${c} — ${CUENTAS[c]||c}`}))}/>
        <Btn size="sm" color="green" onClick={()=>xlsxMayor(partidas)}>⬇</Btn>
      </div>
      <div className="bg-blue-900 text-white rounded-lg px-3 py-2 text-sm font-bold">{sel} — {CUENTAS[sel]||sel}</div>
      {movs.length===0?<p className="text-center text-gray-400 py-6 text-sm">Sin movimientos</p>:(
        <div className="border rounded-lg overflow-hidden text-xs">
          <div className="grid grid-cols-5 bg-gray-100 px-3 py-1.5 font-semibold text-gray-600">
            <span>P#</span><span>Fecha</span><span className="col-span-2">Concepto</span><span className="text-right">Saldo</span>
          </div>
          {movs.map((m,i)=>(
            <div key={i} className={`grid grid-cols-5 px-3 py-1 ${i%2?"bg-white":"bg-gray-50"}`}>
              <span className="text-blue-700 font-semibold">P{m.pnum}</span>
              <span>{m.fecha}</span>
              <span className="col-span-2 truncate text-gray-600">{m.conc}</span>
              <span className={`text-right font-bold ${m.saldo>=0?"text-gray-800":"text-red-600"}`}>{q(Math.abs(m.saldo))}</span>
            </div>
          ))}
          <div className="grid grid-cols-5 px-3 py-1.5 bg-blue-50 font-bold border-t">
            <span className="col-span-4">SALDO FINAL</span>
            <span className={`text-right ${r2(td-th)>=0?"text-green-700":"text-red-600"}`}>{q(r2(td-th))}</span>
          </div>
        </div>
      )}
    </div>
  );
}

// ── APP PRINCIPAL ─────────────────────────────────────────────
export default function App(){
  const[tab,setTab]=useState("dashboard");
  const[ventas,setVentas]=useState([]);
  const[compras,setCompras]=useState([]);
  const[partidas,setPartidas]=useState([]);
  const[startNum,setStartNum]=useState(270);
  const[modoVenta,setModoVenta]=useState("semanal");
  const[provCat,setProvCat]=useState({}); // catálogo proveedor→cuenta aprendido // "semanal" | "individual"
  const[confirmarBorrar,setConfirmarBorrar]=useState(false);
  const[ready,setReady]=useState(false);
  const[storageOk,setStorageOk]=useState(false);

  useEffect(()=>{
    (async()=>{
      try{
        // CARGA — datos pre-procesados + partidas adicionales desde Google Sheets
        const v=DATOS_INICIALES.ventas;
        const c=DATOS_INICIALES.compras.map(x=>({...x,registrado:false}));

        // Cargar partidas adicionales desde Google Sheets (Vercel) o storage local
        let lpGS=null,lpLocal=null,lsn=null,lmv=null;
        try{lpGS=await cargarGS("partidas");}catch(e){}
        try{lpLocal=await cargar("nx4-p");}catch(e){}
        try{lsn=await cargar("nx4-sn");}catch(e){}
        try{lmv=await cargar("nx4-mv");}catch(e){}

        const lpRaw=(lpGS&&lpGS.length>0)?lpGS:(lpLocal&&lpLocal.length>0)?lpLocal:[];
        const baseNums=new Set(PARTIDAS_INICIALES.map(x=>x.num));
        const adicionales=lpRaw.filter(x=>!baseNums.has(x.num));
        const p=[...PARTIDAS_INICIALES,...adicionales];
        console.log("Cargado:",v.length,"v",c.length,"c",p.length,"p (",adicionales.length,"adicionales, fuente:",lpGS?"GS":"local",")");

        setVentas(v);
        setCompras(c);
        setPartidas(p);
        if(lsn)setStartNum(Number(lsn)||270);
        if(lmv)setModoVenta(lmv);
        setStorageOk(true);
        console.log(`Cargado: ${v.length} ventas, ${c.length} compras, ${p.length} partidas`);
      }catch(e){}
      setReady(true);
    })();
  },[]);

  useEffect(()=>{
    if(!ready)return;
    (async()=>{
      // Guardar en storage local (inmediato)
      // Solo guardar partidas adicionales (P274+) — P270-P273 están en el código
      const baseNums2=new Set(PARTIDAS_INICIALES.map(x=>x.num));
      const partidasExtra=partidas.filter(x=>!baseNums2.has(x.num));
      await Promise.all([guardar("nx4-p",partidasExtra),guardar("nx4-sn",startNum),guardar("nx4-mv",modoVenta)]);
      // Guardar en Google Sheets (permanente)
      await Promise.all([
        guardarGS("ventas",ventas),
        guardarGS("compras",compras),
        guardarGS("partidas",partidas),
        guardarGS("config",{startNum,modoVenta}),
      ]);
    })();
  },[ventas,compras,partidas,startNum,modoVenta,ready]);

  // Número siguiente: startNum + partidas.length
  const nextNum=startNum+partidas.length;

  const addVenta=useCallback(v=>{
    if(modoVenta==="individual"){
      const num=startNum+partidas.length;
      const p=genPartidaVenta(v,num);
      setVentas(prev=>[...prev,v]);
      setPartidas(prev=>[...prev,p]);
    } else {
      // Modo semanal: acumula ventas como datos, la partida se crea en "Semanal"
      setVentas(prev=>[...prev,v]);
    }
    setTab("dashboard");
  },[partidas,startNum,modoVenta]);

  const addCompra=useCallback(c=>{
    if(modoVenta==="individual"){
      // Modo individual: crea partida inmediata
      const num=startNum+partidas.length;
      const p=genPartidaGasto(c,num);
      setCompras(prev=>[...prev,c]);
      setPartidas(prev=>[...prev,p]);
    } else {
      // Modo semanal: acumula el gasto sin crear partida
      // Se consolida en Semanal igual que las ventas
      setCompras(prev=>[...prev,{...c,registrado:false}]);
    }
    setTab("semanal");
  },[partidas,startNum,modoVenta]);

  const importVentas=useCallback(rows=>{
    setVentas(prev=>{
      const s=new Set(prev.map(v=>v.serie).filter(Boolean));
      return[...prev,...rows.filter(r=>!r.serie||!s.has(r.serie))];
    });
  },[]);

  const importCompras=useCallback(rows=>{
    setCompras(prev=>{
      const s=new Set(prev.map(c=>c.serie).filter(Boolean));
      const nuevas=rows.filter(r=>!r.serie||!s.has(r.serie)).map(r=>({...r,registrado:false}));
      return[...prev,...nuevas];
    });
  },[]);

  // Registrar partidas de gastos de una semana
  const registrarGastosSemana=useCallback((gastosSem, semLabel)=>{
    // Una sola partida consolidada por semana (igual que ventas)
    const num = startNum+partidas.length;
    const p = genPartidaGastosSem(gastosSem, num, semLabel||"Semana");
    if(!p) return;
    const seriesReg=new Set(gastosSem.map(c=>c.id));
    setCompras(prev=>prev.map(c=>seriesReg.has(c.id)?{...c,registrado:true}:c));
    setPartidas(prev=>{
      const todas=[...prev,p];
      setTimeout(()=>xlsxCompatible(todas),500);
      return todas;
    });
    setTab("diario");
  },[partidas,startNum]);

  const eliminarVenta=useCallback(id=>{
    setVentas(prev=>prev.filter(v=>v.id!==id));
  },[]);

  const eliminarGasto=useCallback(id=>{
    setCompras(prev=>prev.map(c=>c.id===id?{...c,registrado:true}:c));
    // Marcar como registrado lo oculta de la lista de pendientes
  },[]);

  const registrarSem=useCallback(p=>{
    if(!p)return;
    // ── VALIDACIÓN ──────────────────────────────
    const td=r2(p.lineas.reduce((a,l)=>a+l.debe,0));
    const th=r2(p.lineas.reduce((a,l)=>a+l.haber,0));
    if(Math.abs(td-th)>0.10){
      alert(`⚠️ La partida NO cuadra: DEBE Q${td.toFixed(2)} ≠ HABER Q${th.toFixed(2)}. Diferencia: Q${(td-th).toFixed(2)}`);
      return;
    }
    const num=startNum+partidas.length;
    const existe=partidas.find(x=>x.num===num);
    if(existe){
      if(!window.confirm(`⚠️ Ya existe la partida P${num}. ¿Continuar de todas formas?`)) return;
    }
    setPartidas(prev=>{
      const nuevas=[...prev,{...p,num}];
      setTimeout(()=>xlsxCompatible(nuevas),500);
      return nuevas;
    });
    setTab("diario");
  },[partidas,startNum]);

  const limpiar=()=>{
    setVentas([]);setCompras([]);setPartidas([]);
    ["nx4-v","nx4-c","nx4-p","nx4-sn","nx4-mv",
     "nx2-v","nx2-c","nx2-p","nx-v","nx-c","nx-p"
    ].forEach(k=>window.storage?.delete(k).catch(()=>{}));
    setConfirmarBorrar(false);
  };

  const vA=ventas.filter(v=>v.estado!=="ANULADO");
  const cA=compras.filter(c=>c.estado!=="ANULADO");
  const tV=r2(vA.reduce((a,v)=>a+v.base,0));
  const tDF=r2(vA.reduce((a,v)=>a+v.iva,0));
  const tG=r2(cA.reduce((a,c)=>a+c.base,0));
  const tCF=r2(cA.reduce((a,c)=>a+c.iva,0));
  const ivaPos=r2(tDF-tCF);
  const uo=r2(tV-r2(cA.filter(c=>c.cta==="5.1.01").reduce((a,c)=>a+c.base,0))-tG);

  // ISR Q2 proyección
  const UO_ABR_MAY=46354.08;
  const uoQ2=r2(UO_ABR_MAY+uo);
  const isrQ2=r2(uoQ2*0.25);

  const TABS=[
    {id:"dashboard",icon:"📊",label:"Inicio"},
    {id:"importar",icon:"📥",label:"Importar"},
    {id:"venta",icon:"💰",label:"Venta"},
    {id:"gasto",icon:"🧾",label:"Gasto"},
    {id:"semanal",icon:"📅",label:"Semanal"},
    {id:"banco",icon:"🏦",label:"Banco"},
    {id:"diario",icon:"📖",label:"Diario"},
    {id:"mayor",icon:"📋",label:"Mayor"},
    {id:"ef",icon:"📈",label:"E.F."},
    {id:"conciliar",icon:"✅",label:"Conciliar"},
    {id:"config",icon:"⚙️",label:"Config"},
  ];

  return(
    <div className="min-h-screen bg-gray-100 flex flex-col" style={{fontFamily:"Calibri,sans-serif"}}>
      <div className="bg-blue-900 text-white px-4 py-2.5 shadow">
        <div className="flex justify-between items-center">
          <div>
            <p className="font-bold text-sm">Corporación Nexo Global S.A.</p>
            <p className="text-blue-300 text-xs">Junio 2026 · Partidas desde P{startNum} · v4</p>
          </div>
          <div className="flex items-center gap-2">
            {storageOk&&<span className="text-green-400 text-xs" title="Guardando en Google Sheets">☁️💾</span>}
            <span className="text-blue-200 text-xs font-semibold">P{nextNum} próxima</span>
            <button onClick={()=>setConfirmarBorrar(true)} className="text-xs text-red-300 border border-red-400 px-2 py-0.5 rounded">🗑</button>
          </div>
        </div>
      </div>

      <div className="bg-white border-b overflow-x-auto shadow-sm">
        <div className="flex">
          {TABS.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)}
              className={`flex flex-col items-center px-3 py-2 text-xs font-medium whitespace-nowrap border-b-2 transition-colors ${
                tab===t.id?"border-blue-700 text-blue-700 bg-blue-50":"border-transparent text-gray-500 hover:text-gray-700"}`}>
              <span className="text-sm">{t.icon}</span><span className="mt-0.5">{t.label}</span>
            </button>
          ))}
        </div>
      </div>

      <div className="flex-1 p-4 max-w-xl mx-auto w-full pb-8">

        {tab==="dashboard"&&(
          <div className="space-y-4">
            {/* Banner continuidad */}
            <div className="bg-blue-900 text-white rounded-xl p-3 text-xs">
              <div className="flex justify-between items-center">
                <div>
                  <p className="font-bold text-sm">Junio 2026</p>
                  <p className="text-blue-300">Continuación desde P{startNum} · {partidas.length} partidas · {vA.length} ventas</p>
                </div>
                <div className="text-right">
                  <p className="text-blue-300 text-xs">Modo activo:</p>
                  <p className="font-bold">{modoVenta==="semanal"?"📅 Semanal":"📋 Individual"}</p>
                </div>
              </div>
            </div>
            {/* ISR Alert en Dashboard */}
            {(()=>{
              const ef=calcEF(partidas);
              const isr_pen=r2(Math.max(ef.uo,0)*0.25-ef.isr_pag);
              return isr_pen>100?(
                <div className="bg-red-50 border border-red-300 rounded-xl p-3">
                  <p className="text-xs font-bold text-red-800">⚠️ ISR Pendiente de pago</p>
                  <p className="text-xs text-red-600 mt-0.5">
                    ISR devengado Q{q(r2(Math.max(ef.uo,0)*0.25))} — Pagado Q{q(ef.isr_pag)} — <b>Pendiente Q{q(isr_pen)}</b>
                  </p>
                  <Btn size="sm" color="red" onClick={()=>setTab("ef")} className="mt-2">Ver detalle ISR</Btn>
                </div>
              ):null;
            })()}

            {/* Mini EF en Dashboard */}
            {partidas.length>0&&(()=>{
              const ef=calcEF(partidas);
              return(
                <div className="grid grid-cols-3 gap-2">
                  {[
                    {lbl:"Ventas",val:ef.ventas,color:"text-blue-700"},
                    {lbl:"Util. Bruta",val:ef.ub,color:ef.ub>=0?"text-green-700":"text-red-700"},
                    {lbl:"Util. Operativa",val:ef.uo,color:ef.uo>=0?"text-green-700":"text-red-700"},
                  ].map(({lbl,val,color})=>(
                    <div key={lbl} className="bg-white border rounded-xl p-2 text-center" onClick={()=>setTab("ef")}>
                      <div className="text-xs text-gray-400">{lbl}</div>
                      <div className={`text-sm font-bold ${color}`}>{q(val)}</div>
                    </div>
                  ))}
                </div>
              );
            })()}

            {/* Indicador gastos pendientes */}
            {compras.filter(c=>!c.registrado&&c.estado!=="ANULADO").length>0&&(
              <div className="bg-orange-50 border border-orange-300 rounded-xl p-3 flex items-center justify-between">
                <div>
                  <p className="text-xs font-bold text-orange-800">
                    🧾 {compras.filter(c=>!c.registrado&&c.estado!=="ANULADO").length} gastos pendientes de registrar
                  </p>
                  <p className="text-xs text-orange-600 mt-0.5">Vaya a 📅 Semanal → sección naranja al final</p>
                </div>
                <Btn size="sm" color="orange" onClick={()=>setTab("semanal")}>Ir →</Btn>
              </div>
            )}
            {/* Aviso exportar partidas */}
            {partidas.length>0&&(
              <div className="bg-green-50 border border-green-300 rounded-xl p-3 flex items-center justify-between">
                <div>
                  <p className="text-xs font-bold text-green-800">✅ {partidas.length} partidas registradas (P{startNum}–P{startNum+partidas.length-1})</p>
                  <p className="text-xs text-green-600 mt-0.5">Se exporta automáticamente al registrar. También puede exportar manualmente.</p>
                </div>
                <Btn size="sm" color="green" onClick={()=>xlsxCompatible(partidas)}>⬇</Btn>
              </div>
            )}
            {/* Alerta ventas sin consolidar (modo semanal) */}
            {modoVenta==="semanal"&&vA.length>0&&(
              <div className="bg-amber-50 border border-amber-300 rounded-xl p-3 flex items-center justify-between">
                <div>
                  <p className="text-xs font-bold text-amber-800">⚠️ {vA.length} ventas acumuladas sin partida</p>
                  <p className="text-xs text-amber-600 mt-0.5">Ir a Semanal para crear las partidas consolidadas</p>
                </div>
                <Btn size="sm" color="yellow" onClick={()=>setTab("semanal")}>Semanal →</Btn>
              </div>
            )}

            <div className="grid grid-cols-2 gap-3">
              <Card title="Ventas junio (base)" value={q(tV)} sub={`${vA.length} facturas`} color="green"/>
              <Card title="Gastos junio" value={q(tG)} sub={`${cA.length} registros`} color="orange"/>
            </div>

            {/* IVA posición */}
            <div className={`rounded-xl p-3 border ${ivaPos>0?"bg-red-50 border-red-200":"bg-green-50 border-green-200"}`}>
              <div className="flex justify-between text-sm">
                <span>{ivaPos>0?"IVA A PAGAR JUNIO":"SALDO A FAVOR JUNIO"}</span>
                <span className={`font-bold ${ivaPos>0?"text-red-600":"text-green-600"}`}>{q(Math.abs(ivaPos))}</span>
              </div>
              <p className="text-xs text-gray-500 mt-1">DF {q(tDF)} − CF {q(tCF)}</p>
            </div>

            {/* ISR Q2 en vivo */}
            <div className="bg-white rounded-xl border p-4">
              <p className="text-xs font-bold text-gray-600 uppercase mb-3">ISR Trimestre 2 (Abr+May+Jun) — 25%</p>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between border-b pb-1">
                  <span className="text-gray-500">Abr+May UO (real)</span>
                  <span className="font-semibold">{q(UO_ABR_MAY)}</span>
                </div>
                <div className="flex justify-between border-b pb-1">
                  <span className="text-gray-500">Junio UO (acumulado)</span>
                  <span className={`font-semibold ${uo>=0?"text-green-700":"text-red-600"}`}>{q(uo)}</span>
                </div>
                <div className="flex justify-between border-b pb-1">
                  <span className="text-gray-700 font-bold">UO Q2 Total</span>
                  <span className="font-bold">{q(uoQ2)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-700 font-bold">ISR Q2 estimado (25%)</span>
                  <span className="font-bold text-red-600 text-base">{q(isrQ2)}</span>
                </div>
              </div>
              <div className="mt-3 bg-gray-50 rounded-lg p-2 text-xs text-gray-500 text-center">
                Cada Q1,000 de gasto extra en junio → ISR Q2 baja <b>Q250</b>
              </div>
            </div>

            {/* Acciones */}
            <div className="bg-white rounded-xl border p-4">
              <div className="grid grid-cols-2 gap-2 mb-3">
                <Btn color="green" onClick={()=>setTab("venta")}>💰 Nueva Venta</Btn>
                <Btn color="orange" onClick={()=>setTab("gasto")}>🧾 Nuevo Gasto</Btn>
                <Btn color="blue" onClick={()=>setTab("importar")}>📥 Importar FEL</Btn>
                <Btn color="purple" onClick={()=>setTab("semanal")}>📅 Semanal</Btn>
              </div>
              <Btn full color="teal" onClick={()=>xlsxCompatible(partidas)}>
                ⬇ Exportar P{startNum}–P{nextNum-1} para pegar en Excel existente
              </Btn>
            </div>
          </div>
        )}

        {tab==="importar"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">📥</span>
              <div><h2 className="font-bold text-gray-800">Importar FEL Junio</h2>
                <p className="text-xs text-gray-500">Excel con hojas FEL_Ventas y FEL_Compras</p></div>
            </div>
            <ImportarFEL onVentas={importVentas} onCompras={importCompras}/>
          </div>
        )}

        {tab==="venta"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">💰</span>
              <div><h2 className="font-bold text-gray-800">Registrar Venta</h2>
                <p className="text-xs text-gray-500">Cobros mixtos disponibles</p></div>
            </div>
            {/* Selector de modo */}
            <div className="bg-gray-50 border rounded-xl p-3 space-y-2">
              <p className="text-xs font-bold text-gray-600 uppercase tracking-wide">Modo de registro</p>
              <div className="grid grid-cols-2 gap-2">
                <button onClick={()=>setModoVenta("semanal")}
                  className={`py-2.5 px-3 rounded-xl text-xs font-semibold border-2 transition-all text-left ${
                    modoVenta==="semanal"
                      ?"border-blue-600 bg-blue-50 text-blue-700"
                      :"border-gray-200 bg-white text-gray-500"}`}>
                  <div className="text-lg mb-0.5">📅</div>
                  <div className="font-bold">Semanal</div>
                  <div className="text-xs font-normal opacity-75">Acumula ventas — la partida se crea al consolidar en Semanal</div>
                </button>
                <button onClick={()=>setModoVenta("individual")}
                  className={`py-2.5 px-3 rounded-xl text-xs font-semibold border-2 transition-all text-left ${
                    modoVenta==="individual"
                      ?"border-green-600 bg-green-50 text-green-700"
                      :"border-gray-200 bg-white text-gray-500"}`}>
                  <div className="text-lg mb-0.5">📋</div>
                  <div className="font-bold">Individual</div>
                  <div className="text-xs font-normal opacity-75">Crea 1 partida por factura — para corte de caja diario</div>
                </button>
              </div>
              {modoVenta==="semanal"&&(
                <div className="bg-blue-50 rounded-lg px-3 py-2 text-xs text-blue-700 flex gap-2">
                  <span>ℹ️</span>
                  <span>La venta se guarda como dato. Vaya a <b>📅 Semanal</b> al final de la semana para crear la partida consolidada.</span>
                </div>
              )}
              {modoVenta==="individual"&&(
                <div className="bg-green-50 rounded-lg px-3 py-2 text-xs text-green-700 flex gap-2">
                  <span>ℹ️</span>
                  <span>Se crea <b>P{nextNum}</b> por esta venta. No use el Semanal después o duplicará las partidas.</span>
                </div>
              )}
            </div>
            <FormVenta num={nextNum} onSave={addVenta} modo={modoVenta}/>
          </div>
        )}

        {tab==="gasto"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">🧾</span>
              <div>
                <h2 className="font-bold text-gray-800">Registrar Gasto</h2>
                <p className="text-xs text-gray-500">
                  {modoVenta==="semanal"
                    ?"Modo semanal — se acumula en 📅 Semanal para consolidar"
                    :"Modo individual — crea partida inmediata P"+nextNum}
                </p>
              </div>
            </div>
            {modoVenta==="semanal"&&(
              <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-700 flex gap-2">
                <span>💡</span>
                <span>El gasto se agrega a la semana correspondiente. Vaya a <b>📅 Semanal</b> para ver todos los gastos y registrar la partida consolidada.</span>
              </div>
            )}
            <FormGasto num={nextNum} onSave={(c)=>{
              // Aprender cuenta del proveedor
              if(c.proveedor&&c.cta){
                const key=c.proveedor.toUpperCase().slice(0,20);
                setProvCat(prev=>({...prev,[key]:c.cta}));
              }
              addCompra(c);
            }} provCat={provCat}/>
          </div>
        )}

        {tab==="semanal"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">📅</span>
              <div><h2 className="font-bold text-gray-800">Partidas Semanales Junio</h2></div>
            </div>
            <PartidasSemanales ventas={vA} compras={compras} nextNum={nextNum}
              onRegistrar={registrarSem} onRegistrarGastos={registrarGastosSemana}
              onEliminar={eliminarVenta} onEliminarGasto={eliminarGasto}/>
          </div>
        )}

        {tab==="banco"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">🏦</span>
              <div><h2 className="font-bold text-gray-800">Conciliación Bancaria</h2>
                <p className="text-xs text-gray-500">Resumen de depósitos por cuenta y boleta</p></div>
            </div>
            <ConciliacionBanco ventas={vA}/>
          </div>
        )}

        {tab==="diario"&&(
          <div className="space-y-3">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                <span className="text-xl">📖</span>
                <h2 className="font-bold text-gray-800">Libro Diario — Junio</h2>
              </div>
              <span className="text-xs bg-blue-100 text-blue-700 px-2 py-1 rounded-full font-semibold">{partidas.length} partidas</span>
            </div>
            <LibroDiarioView partidas={partidas}/>
          </div>
        )}

        {tab==="mayor"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-xl">📋</span>
              <h2 className="font-bold text-gray-800">Libro Mayor — Junio</h2>
            </div>
            <LibroMayorView partidas={partidas}/>
          </div>
        )}

        {tab==="ef"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">📈</span>
              <div><h2 className="font-bold text-gray-800">Estados Financieros</h2>
                <p className="text-xs text-gray-500">P&L y Balance calculados en tiempo real desde las partidas</p></div>
            </div>
            <EstadosFinancierosView partidas={partidas}/>
          </div>
        )}

        {tab==="conciliar"&&(
          <div className="space-y-3">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">✅</span>
              <div><h2 className="font-bold text-gray-800">Conciliación Automática</h2>
                <p className="text-xs text-gray-500">Saldos por cuenta — Mayor generado automáticamente del Diario</p></div>
            </div>
            <ConciliacionView partidas={partidas}/>
          </div>
        )}

        {tab==="config"&&(
          <div className="space-y-4">
            <div className="flex items-center gap-2 mb-2">
              <span className="text-2xl">⚙️</span>
              <div><h2 className="font-bold text-gray-800">Configuración</h2>
                <p className="text-xs text-gray-500">Ajuste el número inicial de partida</p></div>
            </div>
            <div className="bg-white rounded-xl border p-4 space-y-4">
              <Inp label="Número de primera partida de junio" type="number"
                value={startNum} onChange={v=>setStartNum(Number(v)||270)}/>
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-xs text-blue-700 space-y-1">
                <p className="font-bold">ℹ️ Cómo funciona la continuidad:</p>
                <p>• El Excel histórico tiene P1–P{startNum-1} (hasta mayo)</p>
                <p>• Esta app registra P{startNum} en adelante (junio)</p>
                <p>• Al exportar genera el formato exacto del Excel existente</p>
                <p>• Copie y pegue al final del archivo histórico</p>
              </div>
              <div className="bg-orange-50 border border-orange-200 rounded-lg p-3 text-xs">
                <p className="font-bold text-orange-700 mb-1">📋 Instrucciones para unir los libros:</p>
                <ol className="text-orange-600 space-y-1 list-decimal ml-4">
                  <li>Exporte desde el botón "⬇ Exportar para pegar"</li>
                  <li>Abra el archivo <b>Libros_Contables_FINAL.xlsx</b></li>
                  <li>Vaya a la hoja <b>Libro_Diario</b>, fila 1080</li>
                  <li>Borre la fila de <b>TOTALES GENERALES</b> actual</li>
                  <li>Pegue los datos exportados (fila 4 en adelante)</li>
                  <li>El nuevo TOTALES GENERALES ya está incluido</li>
                </ol>
              </div>
              <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-xs text-green-700">
                <p className="font-bold">✅ Estado actual:</p>
                <p>• Partidas en app: <b>{partidas.length}</b></p>
                <p>• Rango: <b>P{startNum} — P{nextNum-1}</b></p>
                <p>• Auto-guardado: <b>{storageOk?"Activo ✓":"Sin conexión"}</b></p>
              </div>
            </div>
          </div>
        )}

      </div>

      {/* Modal confirmación borrar */}
      {confirmarBorrar&&(
        <div className="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl p-6 max-w-sm w-full shadow-2xl">
            <div className="text-center mb-4">
              <span className="text-4xl">⚠️</span>
              <h3 className="font-bold text-gray-800 text-lg mt-2">¿Borrar todos los datos?</h3>
              <p className="text-sm text-gray-500 mt-1">Se eliminarán todas las ventas, gastos y partidas registradas. Esta acción no se puede deshacer.</p>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <button onClick={()=>setConfirmarBorrar(false)}
                className="py-3 rounded-xl border-2 border-gray-300 text-gray-700 font-bold text-sm hover:bg-gray-50">
                Cancelar
              </button>
              <button onClick={limpiar}
                className="py-3 rounded-xl bg-red-600 text-white font-bold text-sm hover:bg-red-700">
                Sí, borrar todo
              </button>
            </div>
          </div>
        </div>
      )}
      <div className="bg-white border-t px-4 py-2 text-center">
        <p className="text-xs text-gray-400">Nexo Global · NIT 120767147 · Junio 2026 · v4 · P{startNum}+</p>
      </div>
    </div>
  );
}
