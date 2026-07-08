import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ── PARTIDAS VENTAS — PRE-CARGADAS (P270-P273) ───────────────────
// 4 partidas semanales de ventas junio 2026 — cuadradas ✅
const PARTIDAS_INICIALES = [{"num": 270, "fecha": "2026-06-06", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 1", "lineas": [{"cta": "1.1.17.004", "debe": 17299.2, "haber": 0, "conc": "Cobros BI — Semana 1"}, {"cta": "1.1.17.005", "debe": 18855.0, "haber": 0, "conc": "Cobros Banrural — Semana 1"}, {"cta": "4.1.01", "debe": 0, "haber": 32280.54, "conc": "Ventas Semana 1 — 17 facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 3873.66, "conc": "IVA Debito Fiscal 12% — Semana 1"}]}, {"num": 271, "fecha": "2026-06-16", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 2", "lineas": [{"cta": "1.1.17.004", "debe": 20315.4, "haber": 0, "conc": "Cobros BI — Semana 2"}, {"cta": "1.1.17.005", "debe": 79060.0, "haber": 0, "conc": "Cobros Banrural — Semana 2"}, {"cta": "4.1.01", "debe": 0, "haber": 88728.04, "conc": "Ventas Semana 2 — 18 facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 10647.36, "conc": "IVA Debito Fiscal 12% — Semana 2"}]}, {"num": 272, "fecha": "2026-06-23", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 3", "lineas": [{"cta": "1.1.17.004", "debe": 20373.6, "haber": 0, "conc": "Cobros BI — Semana 3"}, {"cta": "1.1.17.005", "debe": 10790.0, "haber": 0, "conc": "Cobros Banrural — Semana 3"}, {"cta": "4.1.01", "debe": 0, "haber": 27824.66, "conc": "Ventas Semana 3 — 16 facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 3338.94, "conc": "IVA Debito Fiscal 12% — Semana 3"}]}, {"num": 273, "fecha": "2026-06-30", "tipo": "VENTA_SEM", "concepto": "Ventas Semana 4", "lineas": [{"cta": "1.1.17.004", "debe": 8429.0, "haber": 0, "conc": "Cobros BI — Semana 4"}, {"cta": "1.1.17.005", "debe": 10775.0, "haber": 0, "conc": "Cobros Banrural — Semana 4"}, {"cta": "4.1.01", "debe": 0, "haber": 17146.42, "conc": "Ventas Semana 4 — facturas"}, {"cta": "2.1.03", "debe": 0, "haber": 2057.58, "conc": "IVA Debito Fiscal 12% — Semana 4"}]}];

// ── DATOS JUNIO 2026 — PRE-CARGADOS DESDE FEL ─────────────────
// Ventas: 53 activas + 2 anuladas | Compras: 23 registros
// Procesado: 26/06/2026 02:46
const DATOS_INICIALES = {
  ventas: [{"id": "v2j26", "fecha": "2026-06-24", "serie": "11A5A333", "tipo": "FACT", "cliente": "CF", "base": 1281.43, "iva": 153.77, "total": 1435.2, "estado": "ACTIVA", "cobros": [{"id": "cxc2", "tipo": "banco", "cta": "1.1.03", "doc": "PENDIENTE", "monto": 1435.2}]}, {"id": "v3j26", "fecha": "2026-06-24", "serie": "F0CF052B", "tipo": "FACT", "cliente": "CF", "base": 263.39, "iva": 31.61, "total": 295.0, "estado": "ACTIVA", "cobros": [{"id": "rr3x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "438822207", "monto": 295.0}]}, {"id": "v4j26", "fecha": "2026-06-23", "serie": "3DC5C874", "tipo": "FACT", "cliente": "CF", "base": 1339.29, "iva": 160.71, "total": 1500.0, "estado": "ACTIVA", "cobros": [{"id": "bi4x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "188466", "monto": 1500.0}]}, {"id": "v5j26", "fecha": "2026-06-22", "serie": "F7C22F08", "tipo": "FACT", "cliente": "CF", "base": 1744.29, "iva": 209.31, "total": 1953.6, "estado": "ACTIVA", "cobros": [{"id": "bi5x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "159399", "monto": 1953.6}]}, {"id": "v6j26", "fecha": "2026-06-22", "serie": "F86B7249", "tipo": "FACT", "cliente": "13945009", "base": 2062.5, "iva": 247.5, "total": 2310.0, "estado": "ACTIVA", "cobros": [{"id": "bi6x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56405909", "monto": 1150.0}, {"id": "bi6x26", "tipo": "banco", "cta": "1.1.17.004", "doc": "56405911", "monto": 1160.0}]}, {"id": "v7j26", "fecha": "2026-06-22", "serie": "3AA36C8B", "tipo": "FACT", "cliente": "69975132", "base": 1392.86, "iva": 167.14, "total": 1560.0, "estado": "ACTIVA", "cobros": [{"id": "bi7x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "226810", "monto": 1560.0}]}, {"id": "v8j26", "fecha": "2026-06-22", "serie": "01B9F9CA", "tipo": "FACT", "cliente": "13495364", "base": 1696.43, "iva": 203.57, "total": 1900.0, "estado": "ACTIVA", "cobros": [{"id": "rr8x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "347946258", "monto": 1900.0}]}, {"id": "v9j26", "fecha": "2026-06-22", "serie": "00523115", "tipo": "FACT", "cliente": "CF", "base": 526.79, "iva": 63.21, "total": 590.0, "estado": "ACTIVA", "cobros": [{"id": "rr9x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "329635163", "monto": 590.0}]}, {"id": "v10j26", "fecha": "2026-06-20", "serie": "96236AAE", "tipo": "FACT", "cliente": "7914172", "base": 2767.86, "iva": 332.14, "total": 3100.0, "estado": "ACTIVA", "cobros": [{"id": "bi10x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "efectivo", "monto": 3100.0}]}, {"id": "v11j26", "fecha": "2026-06-19", "serie": "9478A4A8", "tipo": "FACT", "cliente": "109694252", "base": 848.21, "iva": 101.79, "total": 950.0, "estado": "ACTIVA", "cobros": [{"id": "rr11x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "304267352", "monto": 950.0}]}, {"id": "v12j26", "fecha": "2026-06-19", "serie": "1F8CA694", "tipo": "FACT", "cliente": "CF", "base": 1441.96, "iva": 173.04, "total": 1615.0, "estado": "ACTIVA", "cobros": [{"id": "rr12x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "295339904", "monto": 1615.0}]}, {"id": "v13j26", "fecha": "2026-06-19", "serie": "DAEE010B", "tipo": "FACT", "cliente": "CF", "base": 1330.36, "iva": 159.64, "total": 1490.0, "estado": "ACTIVA", "cobros": [{"id": "rr13x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "106911539", "monto": 590.0}, {"id": "rr13x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "123003232", "monto": 900.0}]}, {"id": "v14j26", "fecha": "2026-06-19", "serie": "8AEAE612", "tipo": "FACT", "cliente": "102283249", "base": 3214.29, "iva": 385.71, "total": 3600.0, "estado": "ACTIVA", "cobros": [{"id": "rr14x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "269087161", "monto": 3600.0}]}, {"id": "v15j26", "fecha": "2026-06-18", "serie": "07B3D3B0", "tipo": "FACT", "cliente": "109808045", "base": 4464.29, "iva": 535.71, "total": 5000.0, "estado": "ACTIVA", "cobros": [{"id": "bi15x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "853002", "monto": 5000.0}]}, {"id": "v16j26", "fecha": "2026-06-18", "serie": "71543EFE", "tipo": "FACT", "cliente": "355062", "base": 1562.5, "iva": 187.5, "total": 1750.0, "estado": "ACTIVA", "cobros": [{"id": "bi16x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56406057", "monto": 1750.0}]}, {"id": "v17j26", "fecha": "2026-06-18", "serie": "42C1DB52", "tipo": "FACT", "cliente": "CF", "base": 218.75, "iva": 26.25, "total": 245.0, "estado": "ACTIVA", "cobros": [{"id": "rr17x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "250930055", "monto": 245.0}]}, {"id": "v18j26", "fecha": "2026-06-18", "serie": "55372C73", "tipo": "FACT", "cliente": "117808105", "base": 357.14, "iva": 42.86, "total": 400.0, "estado": "ACTIVA", "cobros": [{"id": "rr18x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "202261849", "monto": 400.0}]}, {"id": "v19j26", "fecha": "2026-06-18", "serie": "F0E1BCAE", "tipo": "FACT", "cliente": "35527102", "base": 2857.14, "iva": 342.86, "total": 3200.0, "estado": "ACTIVA", "cobros": [{"id": "bi19x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56406056", "monto": 3200.0}]}, {"id": "v20j26", "fecha": "2026-06-16", "serie": "BBBB5A99", "tipo": "FACT", "cliente": "94078203", "base": 1066.96, "iva": 128.04, "total": 1195.0, "estado": "ACTIVA", "cobros": [{"id": "rr20x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "139521616", "monto": 1195.0}]}, {"id": "v21j26", "fecha": "2026-06-16", "serie": "7E316CB1", "tipo": "FACT", "cliente": "CF", "base": 398.57, "iva": 47.83, "total": 446.4, "estado": "ACTIVA", "cobros": [{"id": "bi21x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "222692", "monto": 446.4}]}, {"id": "v22j26", "fecha": "2026-06-15", "serie": "9B5297FC", "tipo": "FACT", "cliente": "306372010", "base": 1457.14, "iva": 174.86, "total": 1632.0, "estado": "ACTIVA", "cobros": [{"id": "bi22x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "172021", "monto": 1632.0}]}, {"id": "v23j26", "fecha": "2026-06-15", "serie": "34F7963D", "tipo": "FACT", "cliente": "CF", "base": 1410.71, "iva": 169.29, "total": 1580.0, "estado": "ACTIVA", "cobros": [{"id": "bi23x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "12146671", "monto": 1580.0}]}, {"id": "v24j26", "fecha": "2026-06-15", "serie": "08FA8FB8", "tipo": "FACT", "cliente": "CF", "base": 263.39, "iva": 31.61, "total": 295.0, "estado": "ACTIVA", "cobros": [{"id": "bi24x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "51622", "monto": 295.0}]}, {"id": "v25j26", "fecha": "2026-06-15", "serie": "8F93A123", "tipo": "FACT", "cliente": "CF", "base": 892.86, "iva": 107.14, "total": 1000.0, "estado": "ACTIVA", "cobros": [{"id": "rr25x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "44132624", "monto": 1000.0}]}, {"id": "v26j26", "fecha": "2026-06-15", "serie": "25D9DC13", "tipo": "FACT", "cliente": "13945009", "base": 2053.57, "iva": 246.43, "total": 2300.0, "estado": "ACTIVA", "cobros": [{"id": "bi26x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404451", "monto": 2300.0}]}, {"id": "v27j26", "fecha": "2026-06-15", "serie": "C695EC3B", "tipo": "FACT", "cliente": "98337211", "base": 267.86, "iva": 32.14, "total": 300.0, "estado": "ACTIVA", "cobros": [{"id": "bi27x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404453", "monto": 300.0}]}, {"id": "v28j26", "fecha": "2026-06-15", "serie": "824C3063", "tipo": "FACT", "cliente": "50637460", "base": 803.57, "iva": 96.43, "total": 900.0, "estado": "ACTIVA", "cobros": [{"id": "bi28x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404457", "monto": 900.0}]}, {"id": "v29j26", "fecha": "2026-06-13", "serie": "8D2EB768", "tipo": "FACT", "cliente": "115998020", "base": 14875.0, "iva": 1785.0, "total": 16660.0, "estado": "ACTIVA", "cobros": [{"id": "rr29x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "100882549", "monto": 16660.0}]}, {"id": "v30j26", "fecha": "2026-06-13", "serie": "9B43E674", "tipo": "FACT", "cliente": "46523626", "base": 1276.79, "iva": 153.21, "total": 1430.0, "estado": "ACTIVA", "cobros": [{"id": "bi30x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404452", "monto": 1430.0}]}, {"id": "v31j26", "fecha": "2026-06-11", "serie": "43D239C7", "tipo": "FACT", "cliente": "115998020", "base": 44406.25, "iva": 5328.75, "total": 49735.0, "estado": "ACTIVA", "cobros": [{"id": "bi31x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404460", "monto": 200.0}, {"id": "rr31x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "38076486", "monto": 49535.0}]}, {"id": "v32j26", "fecha": "2026-06-11", "serie": "14D3D855", "tipo": "FACT", "cliente": "115998020", "base": 16935.0, "iva": 2032.2, "total": 18967.2, "estado": "ANULADO", "cobros": []}, {"id": "v33j26", "fecha": "2026-06-11", "serie": "0060FFF1", "tipo": "FACT", "cliente": "115998020", "base": 49169.64, "iva": 5900.36, "total": 55070.0, "estado": "ANULADO", "cobros": []}, {"id": "v34j26", "fecha": "2026-06-11", "serie": "2C752484", "tipo": "FACT", "cliente": "13945009", "base": 7026.79, "iva": 843.21, "total": 7870.0, "estado": "ACTIVA", "cobros": [{"id": "rr34x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295124", "monto": 1035.0}, {"id": "rr34x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295123", "monto": 2900.0}, {"id": "rr34x41", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295122", "monto": 2900.0}, {"id": "rr34x44", "tipo": "banco", "cta": "1.1.17.005", "doc": "45295121", "monto": 1035.0}]}, {"id": "v35j26", "fecha": "2026-06-11", "serie": "50B87C03", "tipo": "FACT", "cliente": "90134346", "base": 1200.0, "iva": 144.0, "total": 1344.0, "estado": "ACTIVA", "cobros": [{"id": "bi35x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "130778", "monto": 1344.0}]}, {"id": "v36j26", "fecha": "2026-06-11", "serie": "EBEDAB74", "tipo": "FACT", "cliente": "89056876", "base": 1339.29, "iva": 160.71, "total": 1500.0, "estado": "ACTIVA", "cobros": [{"id": "rr36x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "2079830743", "monto": 1500.0}]}, {"id": "v37j26", "fecha": "2026-06-11", "serie": "8E794B90", "tipo": "FACT", "cliente": "48956449", "base": 4414.29, "iva": 529.71, "total": 4944.0, "estado": "ACTIVA", "cobros": [{"id": "bi37x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "167359", "monto": 4944.0}]}, {"id": "v38j26", "fecha": "2026-06-11", "serie": "8B336FBE", "tipo": "FACT", "cliente": "5504538", "base": 4414.29, "iva": 529.71, "total": 4944.0, "estado": "ACTIVA", "cobros": [{"id": "bi38x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "167359", "monto": 4944.0}]}, {"id": "v39j26", "fecha": "2026-06-11", "serie": "07FC2E6C", "tipo": "FACT", "cliente": "CF", "base": 1160.71, "iva": 139.29, "total": 1300.0, "estado": "ACTIVA", "cobros": [{"id": "rr39x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1890192014", "monto": 1300.0}]}, {"id": "v40j26", "fecha": "2026-06-06", "serie": "3096928C", "tipo": "FACT", "cliente": "107462656", "base": 1071.43, "iva": 128.57, "total": 1200.0, "estado": "ACTIVA", "cobros": [{"id": "rr40x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "2005710533", "monto": 1200.0}]}, {"id": "v41j26", "fecha": "2026-06-05", "serie": "5458ECB0", "tipo": "FACT", "cliente": "CF", "base": 441.43, "iva": 52.97, "total": 494.4, "estado": "ACTIVA", "cobros": [{"id": "bi41x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "229586", "monto": 494.4}]}, {"id": "v42j26", "fecha": "2026-06-05", "serie": "BFDAEB4A", "tipo": "FACT", "cliente": "57140626", "base": 1855.71, "iva": 222.69, "total": 2078.4, "estado": "ACTIVA", "cobros": [{"id": "bi42x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "229586", "monto": 2078.4}]}, {"id": "v43j26", "fecha": "2026-06-05", "serie": "5B792CDD", "tipo": "FACT", "cliente": "55359337", "base": 3102.86, "iva": 372.34, "total": 3475.2, "estado": "ACTIVA", "cobros": [{"id": "bi43x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "129700", "monto": 3475.2}]}, {"id": "v44j26", "fecha": "2026-06-05", "serie": "95291FD3", "tipo": "FACT", "cliente": "13945009", "base": 8062.5, "iva": 967.5, "total": 9030.0, "estado": "ACTIVA", "cobros": [{"id": "rr44x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820493", "monto": 905.0}, {"id": "rr44x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820494", "monto": 905.0}, {"id": "rr44x41", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820495", "monto": 5520.0}, {"id": "rr44x44", "tipo": "banco", "cta": "1.1.17.005", "doc": "56820496", "monto": 1700.0}]}, {"id": "v45j26", "fecha": "2026-06-04", "serie": "9877425D", "tipo": "FACT", "cliente": "CF", "base": 1250.0, "iva": 150.0, "total": 1400.0, "estado": "ACTIVA", "cobros": [{"id": "bi45x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "56404455", "monto": 1400.0}]}, {"id": "v46j26", "fecha": "2026-06-04", "serie": "43090455", "tipo": "FACT", "cliente": "30677114", "base": 3125.0, "iva": 375.0, "total": 3500.0, "estado": "ACTIVA", "cobros": [{"id": "bi46x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "174780", "monto": 3500.0}]}, {"id": "v47j26", "fecha": "2026-06-03", "serie": "AC795D67", "tipo": "FACT", "cliente": "108841014", "base": 1414.29, "iva": 169.71, "total": 1584.0, "estado": "ACTIVA", "cobros": [{"id": "bi47x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "158398", "monto": 1584.0}]}, {"id": "v48j26", "fecha": "2026-06-03", "serie": "7E380EDB", "tipo": "FACT", "cliente": "CF", "base": 655.71, "iva": 78.69, "total": 734.4, "estado": "ACTIVA", "cobros": [{"id": "bi48x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "158398", "monto": 734.4}]}, {"id": "v49j26", "fecha": "2026-06-03", "serie": "81A3472A", "tipo": "FACT", "cliente": "103805311", "base": 527.14, "iva": 63.26, "total": 590.4, "estado": "ACTIVA", "cobros": [{"id": "bi49x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "176574", "monto": 590.4}]}, {"id": "v50j26", "fecha": "2026-06-03", "serie": "F61EFC67", "tipo": "FACT", "cliente": "CF", "base": 484.29, "iva": 58.11, "total": 542.4, "estado": "ACTIVA", "cobros": [{"id": "bi50x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "281266", "monto": 542.4}]}, {"id": "v51j26", "fecha": "2026-06-02", "serie": "7398F13C", "tipo": "FACT", "cliente": "117808105", "base": 4285.71, "iva": 514.29, "total": 4800.0, "estado": "ACTIVA", "cobros": [{"id": "rr51x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1854380599", "monto": 2800.0}, {"id": "rr51x38", "tipo": "banco", "cta": "1.1.17.005", "doc": "1854356679", "monto": 2000.0}]}, {"id": "v52j26", "fecha": "2026-06-02", "serie": "30EC02BE", "tipo": "FACT", "cliente": "CF", "base": 816.96, "iva": 98.04, "total": 915.0, "estado": "ACTIVA", "cobros": [{"id": "rr52x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1816504941", "monto": 915.0}]}, {"id": "v53j26", "fecha": "2026-06-02", "serie": "B01FD61B", "tipo": "FACT", "cliente": "CF", "base": 1964.29, "iva": 235.71, "total": 2200.0, "estado": "ACTIVA", "cobros": [{"id": "rr53x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1816504941", "monto": 2200.0}]}, {"id": "v54j26", "fecha": "2026-06-02", "serie": "2BB26C8B", "tipo": "FACT", "cliente": "CF", "base": 415.18, "iva": 49.82, "total": 465.0, "estado": "ACTIVA", "cobros": [{"id": "rr54x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1837379343", "monto": 465.0}]}, {"id": "v55j26", "fecha": "2026-06-02", "serie": "514F71C4", "tipo": "FACT", "cliente": "68723296", "base": 2589.29, "iva": 310.71, "total": 2900.0, "estado": "ACTIVA", "cobros": [{"id": "bi55x23", "tipo": "banco", "cta": "1.1.17.004", "doc": "164937", "monto": 2900.0}]}, {"id": "v56j26", "fecha": "2026-06-02", "serie": "271C48E7", "tipo": "FACT", "cliente": "CF", "base": 218.75, "iva": 26.25, "total": 245.0, "estado": "ACTIVA", "cobros": [{"id": "rr56x35", "tipo": "banco", "cta": "1.1.17.005", "doc": "1797726191", "monto": 245.0}]}],
  compras: [{"id": "c2j26", "fecha": "2026-06-25", "serie": "774BDFF3", "tipo": "FPEQ", "proveedor": "JULIO JOSE , PALACIOS CASTRO", "cta": "6.1.18", "base": 5000.0, "iva": 0, "total": 5000.0, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c3j26", "fecha": "2026-06-25", "serie": "D3B904BC", "tipo": "FPEQ", "proveedor": "JESSICA MADELINE , GARCÍA MARTÍNEZ DE PALACIO", "cta": "6.1.18", "base": 5000.0, "iva": 0, "total": 5000.0, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c4j26", "fecha": "2026-06-25", "serie": "184474D3", "tipo": "FESP", "proveedor": "CORPORACIÓN NEXO GLOBAL, SOCIEDAD ANÓNIMA", "cta": "6.1.14", "base": 17300.0, "iva": 2076.0, "total": 17300.0, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c5j26", "fecha": "2026-06-23", "serie": "7B2A8EA4", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 55.41, "iva": 6.65, "total": 62.06, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c6j26", "fecha": "2026-06-23", "serie": "D5EAF8CD", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 59.61, "iva": 7.15, "total": 66.76, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c7j26", "fecha": "2026-06-23", "serie": "77CD80EE", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 94.52, "iva": 11.34, "total": 105.86, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c8j26", "fecha": "2026-06-21", "serie": "1AF8BA9D", "tipo": "FACT", "proveedor": "DISTRIBUIDORA DE VEHICULOS IMPORTADOS SOCIEDA", "cta": "6.2.05", "base": 1619.65, "iva": 194.36, "total": 1814.01, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c9j26", "fecha": "2026-06-19", "serie": "F022E098", "tipo": "FACT", "proveedor": "UNO GUATEMALA, SOCIEDAD ANONIMA", "cta": "6.1.10", "base": 376.78, "iva": 45.21, "total": 487.99, "impPet": 66.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c10j26", "fecha": "2026-06-19", "serie": "F463A670", "tipo": "FACT", "proveedor": "KEILA PAOLA LOPEZ ORELLANA", "cta": "6.1.14", "base": 959.82, "iva": 115.18, "total": 1075.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c11j26", "fecha": "2026-06-16", "serie": "0D3AA43C", "tipo": "FACT", "proveedor": "POSFILE SOCIEDAD ANONIMA", "cta": "6.1.17", "base": 234.97, "iva": 28.2, "total": 263.17, "impPet": 0.0, "pagoCta": "1.1.17.004", "estado": "ACTIVA", "registrado": false}, {"id": "c12j26", "fecha": "2026-06-15", "serie": "ED7D7111", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 175.0, "iva": 21.0, "total": 196.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c13j26", "fecha": "2026-06-15", "serie": "59CCE986", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 50.64, "iva": 6.08, "total": 56.72, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c14j26", "fecha": "2026-06-09", "serie": "C0BC36FF", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 58.22, "iva": 6.99, "total": 65.21, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c15j26", "fecha": "2026-06-09", "serie": "41961CDA", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 60.55, "iva": 7.27, "total": 67.82, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c16j26", "fecha": "2026-06-06", "serie": "93949C02", "tipo": "FACT", "proveedor": "UNO GUATEMALA, SOCIEDAD ANONIMA", "cta": "6.1.10", "base": 307.76, "iva": 36.93, "total": 398.32, "impPet": 53.63, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c17j26", "fecha": "2026-06-06", "serie": "BE52BA5D", "tipo": "FACT", "proveedor": "PRICESMART (GUATEMALA) SOCIEDAD ANONIMA", "cta": "6.1.14", "base": 89.24, "iva": 10.71, "total": 99.95, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c18j26", "fecha": "2026-06-05", "serie": "4C9B028C", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 50.91, "iva": 6.11, "total": 57.02, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c19j26", "fecha": "2026-06-05", "serie": "78DA4639", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 50.44, "iva": 6.05, "total": 56.49, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c20j26", "fecha": "2026-06-05", "serie": "C4736B3A", "tipo": "FACT", "proveedor": "HETERIA S.A", "cta": "6.2.06", "base": 131.19, "iva": 15.74, "total": 146.93, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c21j26", "fecha": "2026-06-04", "serie": "A003B67E", "tipo": "FACT", "proveedor": "CARGO EXPRESO, SOCIEDAD ANONIMA", "cta": "6.2.06", "base": 892.86, "iva": 107.14, "total": 1000.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c22j26", "fecha": "2026-06-03", "serie": "E484ED92", "tipo": "FACT", "proveedor": "TATMON, SOCIEDAD ANONIMA", "cta": "6.1.17", "base": 266.96, "iva": 32.04, "total": 299.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}, {"id": "c23j26", "fecha": "2026-06-02", "serie": "CC6AF47C", "tipo": "FACT", "proveedor": "CARGO EXPRESO, SOCIEDAD ANONIMA", "cta": "6.2.06", "base": 204.29, "iva": 24.51, "total": 228.8, "impPet": 0.0, "pagoCta": "2.1.01", "estado": "ACTIVA", "registrado": false}, {"id": "c24j26", "fecha": "2026-06-01", "serie": "0695A96E", "tipo": "FACT", "proveedor": "SERVICIOS INNOVADORES DE COMUNICACIÓN Y ENTRE", "cta": "6.1.17", "base": 178.57, "iva": 21.43, "total": 200.0, "impPet": 0.0, "pagoCta": "1.1.02", "estado": "ACTIVA", "registrado": false}]
};

// ── CATÁLOGO ───────────────────────────────────────────────────
const CUENTAS = {
  "1.1.02":"Caja Chica","1.1.03":"Clientes / Caja General",
  "1.1.04":"Inventario","1.1.05":"Anticipos a Proveedores",
  "1.1.06":"IVA Credito Fiscal","1.1.08":"IVA Retenido x Cobrar",
  "1.1.03":"Clientes / Cuentas x Cobrar",
  "1.1.17.004":"Banco Industrial Q","1.1.17.005":"Banrural Q","1.1.17.006":"BI USD",
  "2.1.01":"Cuentas por Pagar","2.1.02":"Prestamo Socio","2.1.03":"IVA Debito Fiscal",
  "2.1.04":"Anticipos de Clientes","2.1.05":"Depositos x Identificar",
  "2.1.06":"ISR Retenido x Pagar","2.1.07":"Anticipos de Clientes",
  "3.1.01":"Utilidades Retenidas","3.1.02":"Capital Pagado",
  "4.1.01":"Ventas de Bienes","5.1.01":"Costo de Ventas",
  "6.1.10":"Combustible","6.1.13":"Agua, Luz y Telefono",
  "6.1.14":"Insumos de Oficina","6.1.15":"Serv. Contables",
  "6.1.16":"Servicios Legales","6.1.17":"Programas y Software",
  "6.1.18":"Serv. Tecnologicos","6.1.19":"Afiliaciones y Membresias",
  "6.2.02":"Comisiones s/Ventas","6.2.03":"Paqueteria y Envios",
  "6.2.04":"Combustible Vehiculo","6.2.05":"Arrend. y Mant. Vehiculos",
  "6.2.06":"Paqueteria y Fletes","6.3.01":"Gastos Financieros","6.4.01":"ISR",
};

const BANCOS_PAGO = [
  {cta:"1.1.17.004",nom:"Banco Industrial Q"},
  {cta:"1.1.17.005",nom:"Banrural Q"},
  {cta:"1.1.02",nom:"Caja Chica"},
  {cta:"2.1.01",nom:"Credito / CxP"},
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
  {cta:"6.1.10",nom:"Combustible",iva:"petro"},
  {cta:"6.1.13",nom:"Agua, Luz y Telefono",iva:"normal"},
  {cta:"6.1.14",nom:"Insumos de Oficina",iva:"normal"},
  {cta:"6.1.15",nom:"Serv. Contables",iva:"normal"},
  {cta:"6.1.16",nom:"Servicios Legales",iva:"normal"},
  {cta:"6.1.17",nom:"Programas / Software",iva:"normal"},
  {cta:"6.1.18",nom:"Serv. Tecnologicos",iva:"fpeq"},
  {cta:"6.1.19",nom:"Afiliaciones y Membresias",iva:"normal"},
  {cta:"6.2.02",nom:"Comisiones s/Ventas",iva:"normal"},
  {cta:"6.2.03",nom:"Paqueteria y Envios",iva:"normal"},
  {cta:"6.2.04",nom:"Combustible Vehiculo",iva:"petro"},
  {cta:"6.2.05",nom:"Arrend. y Mant. Vehiculos",iva:"normal"},
  {cta:"6.2.06",nom:"Paqueteria y Fletes",iva:"normal"},
  {cta:"6.3.01",nom:"Gastos Financieros",iva:"normal"},
  {cta:"6.4.01",nom:"ISR Trimestral",iva:"fpeq"},
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

  // Agrupar DEBE por cuenta de gasto
  const porCta = {};
  gastos.forEach(g=>{
    const k = g.cta;
    if(!porCta[k]) porCta[k] = {base:0, concs:[]};
    porCta[k].base = r2(porCta[k].base + g.base);
    porCta[k].concs.push(g.proveedor?.slice(0,20)||"");
  });
  Object.entries(porCta).sort().forEach(([cta,d])=>{
    L.push({cta, debe:d.base, haber:0,
      conc:`${CUENTAS[cta]||cta} — ${semLabel}`});
  });

  // IVA CF consolidado
  const tCF = r2(gastos.reduce((a,g)=>a+g.iva,0));
  if(tCF>0) L.push({cta:"1.1.06", debe:tCF, haber:0,
    conc:`IVA CF — ${semLabel}`});

  // ISR FESP consolidado
  const tISR = r2(gastos.filter(isFesp).reduce((a,g)=>a+r2(g.base*0.05),0));
  if(tISR>0) L.push({cta:"6.4.01", debe:tISR, haber:0,
    conc:`ISR Retenido FE 5% — ${semLabel}`});

  // HABER por forma de pago
  const porPago = {};
  gastos.forEach(g=>{
    const k = g.pagoCta||"1.1.02";
    porPago[k] = r2((porPago[k]||0) + g.base + g.iva);
  });
  Object.entries(porPago).forEach(([cta,tot])=>{
    L.push({cta, debe:0, haber:tot,
      conc:`Pago gastos ${semLabel}`});
  });

  // IVA DT FESP y ISR FESP como pasivos
  const tIVAFE = r2(gastos.filter(isFesp).reduce((a,g)=>a+g.iva,0));
  if(tIVAFE>0) L.push({cta:"2.1.03", debe:0, haber:tIVAFE,
    conc:`IVA DT ret. FE — ${semLabel}`});
  if(tISR>0)   L.push({cta:"2.1.06", debe:0, haber:tISR,
    conc:`ISR Retenido FE x Pagar — ${semLabel}`});

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
  const isFesp=c.tipo==="FESP"||c.tipo==="FESP";
  const isr=isFesp?r2(c.base*0.05):0; // ISR 5% solo en Factura Especial
  if(c.base>0) L.push({cta:c.cta,debe:r2(c.base),haber:0,conc:`${CUENTAS[c.cta]||c.cta} - ${c.proveedor}`});
  if(c.iva>0)  L.push({cta:"1.1.06",debe:r2(c.iva),haber:0,conc:`IVA CF - ${c.proveedor}`});
  if(isr>0)    L.push({cta:"6.4.01",debe:isr,haber:0,conc:`ISR Retenido FE 5% - ${c.proveedor}`});
  L.push({cta:c.pagoCta||"1.1.02",debe:0,haber:r2(c.base+c.iva),conc:`Pago - ${c.proveedor}`});
  if(c.iva>0&&isFesp) L.push({cta:"2.1.03",debe:0,haber:r2(c.iva),conc:`IVA DT ret. FE - ${c.proveedor}`});
  if(isr>0)    L.push({cta:"2.1.06",debe:0,haber:isr,conc:`ISR Retenido FE x Pagar - ${c.proveedor}`});
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
function xlsxCompatible(partidas, modo="completo"){
  const wb=XLSX.utils.book_new();
  const rows=[];
  // Encabezado igual al archivo existente
  rows.push(["","","","","","","",""]);
  rows.push(["","","","","","","",""]);
  rows.push(["N°","FECHA","CUENTA","DESCRIPCION","CONCEPTO / REFERENCIA","PARCIAL Q.","DEBE Q.","HABER Q."]);

  let totD=0,totH=0;
  partidas.forEach(p=>{
    const sd=r2(p.lineas.reduce((a,l)=>a+l.debe,0));
    const sh=r2(p.lineas.reduce((a,l)=>a+l.haber,0));
    p.lineas.forEach((l,i)=>{
      rows.push([
        i===0?p.num:"",
        i===0?p.fecha:"",
        l.cta,
        CUENTAS[l.cta]||l.cta,
        i===0?l.conc:l.conc,
        "",
        l.debe||"",
        l.haber||"",
      ]);
    });
    rows.push(["","","","",`TOTALES PARTIDA  ${p.num}`,"",sd,sh]);
    totD+=sd; totH+=sh;
  });
  rows.push(["","","","","TOTALES GENERALES","",r2(totD),r2(totH)]);

  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=[{wch:6},{wch:13},{wch:14},{wch:28},{wch:42},{wch:10},{wch:13},{wch:13}];
  const shName=modo==="pegar"?"P270_en_adelante":"Libro_Diario_Jun2026";
  XLSX.utils.book_append_sheet(wb,ws,shName);

  // Hoja de instrucciones
  const inst=[
    ["INSTRUCCIONES PARA PEGAR EN EL EXCEL EXISTENTE"],
    [""],
    ["1. Abra el archivo Libros_Contables_FINAL.xlsx"],
    ["2. Vaya a la hoja 'Libro_Diario'"],
    ["3. Ubíquese en la fila 1080 (después de TOTALES GENERALES actuales)"],
    ["4. Borre la fila de TOTALES GENERALES actual"],
    ["5. Copie las filas de datos de esta hoja (desde fila 4 en adelante)"],
    ["6. Péguelas a partir de la fila donde estaba el TOTALES GENERALES"],
    ["7. Los TOTALES GENERALES de esta hoja ya incluyen SOLO las partidas nuevas"],
    [""],
    [`Total partidas exportadas: ${partidas.length}`],
    [`Rango: P${partidas[0]?.num||""} — P${partidas[partidas.length-1]?.num||""}`],
    [`Fecha exportación: ${new Date().toLocaleDateString("es-GT")}`],
  ];
  const ws2=XLSX.utils.aoa_to_sheet(inst);
  ws2["!cols"]=[{wch:70}];
  XLSX.utils.book_append_sheet(wb,ws2,"Instrucciones");
  XLSX.writeFile(wb,"NexoGlobal_Junio2026_P270enAdelante.xlsx");
}

function xlsxBanco(ventas){
  const wb=XLSX.utils.book_new();
  const cuentas=[
    {cta:"1.1.17.004",nom:"Banco Industrial Q"},
    {cta:"1.1.17.005",nom:"Banrural Q"},
    {cta:"1.1.03",nom:"CxC Pendiente"},
  ];
  // Hoja resumen
  const resumen=[
    ["CONCILIACIÓN BANCARIA — CORPORACIÓN NEXO GLOBAL S.A.","","",""],
    ["NIT: 120767147 | Período: Junio 2026","","",""],
    [""],
    ["Cuenta","N° Depósitos","Total Q",""],
  ];
  cuentas.forEach(ct=>{
    const movs=ventas.flatMap(v=>(v.cobros||[]).filter(c=>
      (c.tipo==="banco"&&c.cta===ct.cta)||
      (ct.cta==="1.1.17.004"&&c.tipo==="visa")
    ));
    const tot=r2(movs.reduce((s,c)=>s+(c.tipo==="visa"
      ?r2(Number(c.visaNet||0)+Number(c.visaRet||0)+Number(c.visaCom||0)+Number(c.visaMem||0))
      :Number(c.monto||0)),0));
    if(tot>0) resumen.push([ct.nom,movs.length,tot,""]);
  });
  const totG=r2(ventas.flatMap(v=>v.cobros||[]).reduce((s,c)=>s+(c.tipo==="visa"
    ?r2(Number(c.visaNet||0)+Number(c.visaRet||0)+Number(c.visaCom||0)+Number(c.visaMem||0))
    :Number(c.monto||0)),0));
  resumen.push([""],["TOTAL COBROS","",totG,""]);
  const ws0=XLSX.utils.aoa_to_sheet(resumen);
  ws0["!cols"]=[{wch:30},{wch:15},{wch:15},{wch:5}];
  XLSX.utils.book_append_sheet(wb,ws0,"Resumen");

  // Hoja detalle por semana y banco
  const SEM=[{l:"Semana 1",d1:1,d2:9},{l:"Semana 2",d1:10,d2:16},
             {l:"Semana 3",d1:17,d2:23},{l:"Semana 4",d1:24,d2:31}];
  const det=[["Semana","Fecha","Serie DTE","Cliente","Banco/Cuenta","N° Boleta/Referencia","Monto Q"]];
  const getDay=f=>{const m=String(f||"").match(/-(\d{2})(?:T|$)/);return m?parseInt(m[1]):0;};
  ventas.filter(v=>v.estado!=="ANULADO").forEach(v=>{
    const sem=SEM.find(s=>{const d=getDay(v.fecha);return d>=s.d1&&d<=s.d2;});
    (v.cobros||[]).forEach(c=>{
      if(c.tipo==="banco"){
        det.push([sem?.l||"",v.fecha?.slice(0,10),v.serie,v.cliente,
          CUENTAS[c.cta]||c.cta,c.doc||"",Number(c.monto||0)]);
      }
      if(c.tipo==="visa"){
        const tot=r2(Number(c.visaNet||0)+Number(c.visaRet||0)+Number(c.visaCom||0)+Number(c.visaMem||0));
        det.push([sem?.l||"",v.fecha?.slice(0,10),v.serie,v.cliente,
          "VISA/Tarjeta",c.doc||"",tot]);
      }
    });
  });
  det.push(["","","","","","TOTAL",totG]);
  const ws1=XLSX.utils.aoa_to_sheet(det);
  ws1["!cols"]=[{wch:10},{wch:12},{wch:12},{wch:30},{wch:22},{wch:20},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws1,"Detalle Depositos");
  XLSX.writeFile(wb,"ConciliacionBancaria_Junio2026.xlsx");
}

function xlsxMayor(partidas){
  // Una sola hoja con TODAS las cuentas — separadas por encabezado
  const ctas=[...new Set(partidas.flatMap(p=>p.lineas.map(l=>l.cta)))].sort();
  const wb=XLSX.utils.book_new();
  const rows=[
    ["LIBRO MAYOR — CORPORACION NEXO GLOBAL S.A.","","","","",""],
    ["NIT: 120767147 | Período: Junio 2026","","","","",""],
    ["","","","","",""],
  ];

  ctas.forEach(cta=>{
    // Encabezado de cuenta
    rows.push([`CUENTA: ${cta} — ${CUENTAS[cta]||cta}`,"","","","",""]);
    rows.push(["Asiento","Fecha","Concepto","Debe Q","Haber Q","Saldo Q"]);
    let s=0;
    partidas.forEach(p=>p.lineas.filter(l=>l.cta===cta).forEach(l=>{
      s=r2(s+l.debe-l.haber);
      rows.push([p.num,p.fecha,l.conc,l.debe||"",l.haber||"",s]);
    }));
    const td=r2(rows.filter(r=>typeof r[3]==="number").slice(-99)
      .reduce((a,r)=>a+(typeof r[3]==="number"?r[3]:0),0)); // approx
    // Calcular totales correctamente
    let tD=0,tH=0;
    partidas.forEach(p=>p.lineas.filter(l=>l.cta===cta).forEach(l=>{tD+=l.debe;tH+=l.haber;}));
    rows.push(["TOTALES","","",r2(tD),r2(tH),r2(tD-tH)]);
    rows.push(["","","","","",""]);  // línea en blanco entre cuentas
  });

  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=[{wch:10},{wch:12},{wch:42},{wch:13},{wch:13},{wch:14}];
  XLSX.utils.book_append_sheet(wb,ws,"Libro_Mayor");
  XLSX.writeFile(wb,"LibroMayor_Junio2026.xlsx");
}

// ── UI BASE ────────────────────────────────────────────────────
const Btn=({onClick,children,color="blue",size="md",disabled=false,full=false})=>{
  const c={blue:"bg-blue-700 hover:bg-blue-800",green:"bg-green-700 hover:bg-green-800",
    orange:"bg-orange-600 hover:bg-orange-700",red:"bg-red-600 hover:bg-red-700",
    gray:"bg-gray-500 hover:bg-gray-600",purple:"bg-purple-700 hover:bg-purple-800",
    teal:"bg-teal-700 hover:bg-teal-800",yellow:"bg-yellow-600 hover:bg-yellow-700"}[color];
  const s={md:"px-4 py-2 text-sm",sm:"px-3 py-1.5 text-xs",lg:"px-5 py-3 text-base",xs:"px-2 py-1 text-xs"}[size];
  return <button onClick={onClick} disabled={disabled}
    className={`${c} ${s} ${full?"w-full":""} text-white font-semibold rounded-lg disabled:bg-gray-200 disabled:text-gray-400 transition-colors`}>{children}</button>;
};
const Card=({title,value,sub,color="blue"})=>{
  const c={blue:"border-blue-500 bg-blue-50",green:"border-green-500 bg-green-50",orange:"border-orange-500 bg-orange-50",red:"border-red-500 bg-red-50"}[color];
  return <div className={`border-l-4 rounded-lg p-3 ${c}`}>
    <p className="text-xs text-gray-500 font-semibold uppercase tracking-wide">{title}</p>
    <p className="text-lg font-bold text-gray-800 mt-0.5">{value}</p>
    {sub&&<p className="text-xs text-gray-400 mt-0.5">{sub}</p>}
  </div>;
};
const Inp=({label,type="text",value,onChange,placeholder="",min,step,className=""})=>
  <div className={className}>
    {label&&<label className="block text-xs font-semibold text-gray-600 mb-1">{label}</label>}
    <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} min={min} step={step}
      className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"/>
  </div>;
const Sel=({label,value,onChange,options,className=""})=>
  <div className={className}>
    {label&&<label className="block text-xs font-semibold text-gray-600 mb-1">{label}</label>}
    <select value={value} onChange={e=>onChange(e.target.value)}
      className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
      {options.map(o=><option key={o.value||o} value={o.value||o}>{o.label||o}</option>)}
    </select>
  </div>;

const PreviewPartida=({p,colorHdr="bg-blue-900"})=>{
  if(!p||!p.lineas.length) return null;
  const sd=r2(p.lineas.reduce((a,l)=>a+l.debe,0)),sh=r2(p.lineas.reduce((a,l)=>a+l.haber,0));
  const ok=Math.abs(sd-sh)<0.02;
  return <div className="border rounded-lg overflow-hidden text-xs shadow-sm">
    <div className={`${colorHdr} text-white px-3 py-1.5 flex justify-between font-semibold`}>
      <span>Vista previa — P{p.num}</span>
      <span>{ok?<span className="text-green-300">✅ Cuadra</span>:<span className="text-red-300">⚠️ Dif Q{r2(Math.abs(sd-sh)).toFixed(2)}</span>}</span>
    </div>
    {p.lineas.map((l,i)=>(
      <div key={i} className={`grid grid-cols-4 px-3 py-0.5 ${i%2?"bg-white":"bg-gray-50"}`}>
        <span className="font-mono text-blue-700">{l.cta}</span>
        <span className="text-gray-600 truncate">{CUENTAS[l.cta]||l.cta}</span>
        <span className="text-right text-green-700 font-semibold">{l.debe?q(l.debe):""}</span>
        <span className="text-right text-red-700 font-semibold">{l.haber?q(l.haber):""}</span>
      </div>
    ))}
    <div className="grid grid-cols-4 px-3 py-1.5 bg-gray-100 font-bold border-t">
      <span className="col-span-2">TOTALES P{p.num}</span>
      <span className="text-right">{q(sd)}</span><span className="text-right">{q(sh)}</span>
    </div>
  </div>;
};

// ── COBROS EDITOR ─────────────────────────────────────────────
function CobrosEditor({cobros,onChange,totalFactura}){
  const totalCobros=calcTotalCobros(cobros);
  const diff=r2(totalFactura-totalCobros);
  const ok=Math.abs(diff)<0.50; // tolerancia Q0.50 para redondeos bancarios
  const addCobro=(tipo)=>{
    const b={id:uid(),tipo,doc:"",monto:""};
    if(tipo==="banco")b.cta="1.1.17.004";
    if(tipo==="efectivo")b.cta="1.1.02";
    if(tipo==="credito")b.cta="2.1.01";
    if(tipo==="deposito")b.cta="2.1.05";
    if(tipo==="visa")Object.assign(b,{cta:"VISA",visaNet:"",visaRet:"",visaCom:"",visaMem:""});
    onChange([...cobros,b]);
  };
  const upd=(id,f,v)=>onChange(cobros.map(c=>c.id===id?{...c,[f]:v}:c));
  const del=(id)=>onChange(cobros.filter(c=>c.id!==id));
  const autoVisa=(id)=>{
    // totalFactura ya es el total c/IVA del VISA
    const thisCobro=cobros.find(c=>c.id===id);
    const tot=r2(Number(thisCobro?.visaNet||totalFactura)); // total c/IVA del cobro VISA
    const ivaV=r2(tot/1.12*0.12);    // IVA sobre el monto VISA
    const ret=r2(ivaV*0.15);          // retención 15% sobre IVA
    const cb=r2(tot*0.0625);          // comisión base 6.25%
    const ci=r2(cb*0.12);             // IVA de la comisión
    const net=r2(tot-ret-cb-ci-25);   // neto depositado
    onChange(cobros.map(c=>c.id===id?{...c,visaNet:net.toFixed(2),visaRet:ret.toFixed(2),visaCom:(cb+ci).toFixed(2),visaMem:"25.00"}:c));
  };
  return(
    <div className="space-y-3">
      {totalFactura>0&&(
        <div className={`rounded-lg px-4 py-2.5 flex items-center justify-between text-sm ${ok?"bg-green-50 border border-green-300":"bg-red-50 border border-red-300"}`}>
          <span className="text-gray-600">Total factura <b>{q(totalFactura)}</b> · Cobros <b>{q(totalCobros)}</b></span>
          <span className={`font-bold ${ok?"text-green-700":"text-red-600"}`}>
            {ok?(Math.abs(diff)<0.02?"✅ Cuadra":`✅ Cuadra (dif Q${Math.abs(diff).toFixed(2)} redondeo)`):`⚠️ Falta ${q(Math.abs(diff))}`}
          </span>
        </div>
      )}
      {cobros.map((c,idx)=>{
        const med=MEDIOS_PAGO.find(m=>m.tipo===c.tipo);
        return(
          <div key={c.id} className="border border-gray-200 rounded-xl overflow-hidden">
            <div className="flex items-center gap-2 px-3 py-2 bg-gray-50 border-b">
              <span>{med?.icon||"💳"}</span>
              <span className="text-xs font-bold text-gray-700 flex-1">Cobro {idx+1} — {med?.nom||c.tipo}</span>
              {c.tipo==="banco"&&(
                <select value={c.cta} onChange={e=>upd(c.id,"cta",e.target.value)}
                  className="text-xs border rounded px-2 py-1 bg-white">
                  <option value="1.1.17.004">BI Q</option>
                  <option value="1.1.17.005">Banrural Q</option>
                  <option value="1.1.17.006">BI USD</option>
                  <option value="1.1.03">CxC Pendiente</option>
                </select>
              )}
              <button onClick={()=>del(c.id)} className="text-red-400 hover:text-red-600 font-bold text-lg px-1">×</button>
            </div>
            <div className="px-3 py-2 space-y-2">
              {(c.tipo==="banco"||c.tipo==="efectivo"||c.tipo==="credito")&&(
                <div className="grid grid-cols-2 gap-2">
                  <Inp label="N° Boleta / Depósito" value={c.doc} onChange={v=>upd(c.id,"doc",v)} placeholder="Ej: 11236364"/>
                  <Inp label="Monto Q" type="number" step="0.01" value={c.monto} onChange={v=>upd(c.id,"monto",v)} placeholder="0.00"/>
                </div>
              )}
              {c.tipo==="deposito"&&(
                <div className="space-y-1">
                  <div className="grid grid-cols-2 gap-2">
                    <Inp label="N° Referencia" value={c.doc} onChange={v=>upd(c.id,"doc",v)}/>
                    <Inp label="+aplicar / −exceso" type="number" step="0.01" value={c.monto} onChange={v=>upd(c.id,"monto",v)} placeholder="+200 ó −15"/>
                  </div>
                  <p className="text-xs text-gray-500">Positivo = aplicar saldo 2.1.05 · Negativo = registrar exceso</p>
                </div>
              )}
              {c.tipo==="visa"&&(
                <div className="space-y-2">
                  <div className="flex justify-between items-center">
                    <span className="text-xs text-purple-700 font-semibold">Desglose VISA</span>
                    <button onClick={()=>autoVisa(c.id)} className="text-xs bg-purple-600 text-white px-3 py-1 rounded-full">⚡ Auto-calcular</button>
                  </div>
                  <Inp label="N° Referencia / Lote" value={c.doc} onChange={v=>upd(c.id,"doc",v)} placeholder="17425"/>
                  <div className="grid grid-cols-2 gap-2">
                    <Inp label="Neto depositado BI" type="number" step="0.01" value={c.visaNet} onChange={v=>upd(c.id,"visaNet",v)}/>
                    <Inp label="IVA Retenido 15%" type="number" step="0.01" value={c.visaRet} onChange={v=>upd(c.id,"visaRet",v)}/>
                    <Inp label="Comisión bruta" type="number" step="0.01" value={c.visaCom} onChange={v=>upd(c.id,"visaCom",v)}/>
                    <Inp label="Membresía Q25" type="number" step="0.01" value={c.visaMem} onChange={v=>upd(c.id,"visaMem",v)}/>
                  </div>
                </div>
              )}
            </div>
          </div>
        );
      })}
      <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide">+ Agregar forma de cobro</p>
      <div className="grid grid-cols-3 gap-2">
        {[{tipo:"banco",icon:"🏦",label:"Depósito banco"},{tipo:"efectivo",icon:"💵",label:"Efectivo"},
          {tipo:"visa",icon:"💳",label:"Tarjeta VISA"},{tipo:"deposito",icon:"❓",label:"Dep. x Identificar"},
          {tipo:"credito",icon:"📋",label:"Crédito"}].map(({tipo,icon,label})=>(
          <button key={tipo} onClick={()=>addCobro(tipo)}
            className="flex flex-col items-center py-2 border-2 border-dashed border-gray-300 rounded-xl text-xs text-gray-600 hover:border-blue-400 hover:bg-blue-50 transition-all">
            <span className="text-xl mb-0.5">{icon}</span><span className="font-semibold">{label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ── FORM VENTA ────────────────────────────────────────────────
function FormVenta({onSave,num}){
  const[f,setF]=useState({fecha:new Date().toISOString().slice(0,10),serie:"",cliente:"",
    totalConIva:"",cobros:[]});
  // SAT certifica el total c/IVA — app calcula base e IVA automáticamente
  const totalConIva=r2(f.totalConIva);
  const base=r2(totalConIva/1.12);
  const iva=r2(totalConIva-base);   // evita doble redondeo
  const total=totalConIva;           // total exacto = lo que depositó el cliente
  const clienteFinal=f.cliente.trim()||"Consumidor Final";
  const v={...f,cliente:clienteFinal,base,iva,total,totalConIva:total};
  const p=base>0?genPartidaVenta(v,num):null;
  return(
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-3">
        <Inp label="Fecha" type="date" value={f.fecha} onChange={d=>setF(p=>({...p,fecha:d}))}/>
        <Inp label="Serie DTE (8 chars)" value={f.serie} placeholder="BF23B23A"
          onChange={s=>setF(p=>({...p,serie:s.toUpperCase().slice(0,8)}))}/>
      </div>
      <Inp label="Cliente (opcional — deje vacío para CF)" value={f.cliente}
        placeholder="Consumidor Final"
        onChange={c=>setF(p=>({...p,cliente:c}))}/>
      <Inp label="Total c/IVA (Q) — igual al comprobante SAT" type="number" step="0.01"
        value={f.totalConIva} onChange={b=>setF(p=>({...p,totalConIva:b}))}
        placeholder="Ej: 1,120.00"/>
      {totalConIva>0&&<div className="grid grid-cols-3 gap-2 bg-blue-50 rounded-lg p-3 text-sm text-center border border-blue-200">
        <div><div className="text-xs text-gray-500">Base s/IVA</div><b>{q(base)}</b></div>
        <div><div className="text-xs text-gray-500">IVA 12% (auto)</div><b className="text-blue-600">{q(iva)}</b></div>
        <div><div className="text-xs text-gray-500">Total c/IVA</div><b className="text-green-600">{q(total)}</b></div>
      </div>}
      <div className="border-t pt-3">
        <p className="text-sm font-bold text-gray-700 mb-3">💳 Forma(s) de cobro <span className="text-xs font-normal text-gray-400">— puede mezclar varias</span></p>
        <CobrosEditor cobros={f.cobros} totalFactura={total} onChange={c=>setF(p=>({...p,cobros:c}))}/>
      </div>
      {p&&<PreviewPartida p={p}/>}
      <Btn full color="green" disabled={!totalConIva}
        onClick={()=>{if(totalConIva){onSave({...v,cobros:f.cobros,id:`v-${Date.now()}`});
          setF({fecha:f.fecha,serie:"",cliente:"",totalConIva:"",cobros:[]});}}}>
        ✔ Registrar Venta — P{num}
      </Btn>
    </div>
  );
}

// ── FORM GASTO ────────────────────────────────────────────────
function FormGasto({onSave,num}){
  const[g,setG]=useState({fecha:new Date().toISOString().slice(0,10),proveedor:"",
    cta:"6.1.14",total:"",impPet:"",pagoCta:"1.1.02",tipo:"FACT",serie:""});
  const cat=GASTOS_CAT.find(c=>c.cta===g.cta)||{iva:"normal"};
  let base=0,iva=0;
  if(cat.iva==="fpeq"){base=r2(g.total);iva=0;}
  else if(cat.iva==="petro"&&g.impPet){const np=r2(r2(g.total)-r2(g.impPet));base=r2(np/1.12);iva=r2(np/1.12*0.12);}
  else if(g.total){base=r2(r2(g.total)/1.12);iva=r2(r2(g.total)/1.12*0.12);}
  const c={...g,base,iva,total:r2(base+iva),estado:"ACTIVA"};
  const p=base>0?genPartidaGasto(c,num):null;
  return(
    <div className="space-y-3">
      <div className="grid grid-cols-2 gap-3">
        <Inp label="Fecha" type="date" value={g.fecha} onChange={d=>setG(p=>({...p,fecha:d}))}/>
        <Inp label="Serie DTE" value={g.serie} placeholder="EE0ABA9D"
          onChange={s=>setG(p=>({...p,serie:s.toUpperCase().slice(0,8)}))}/>
      </div>
      <Inp label="Proveedor" value={g.proveedor} placeholder="Nombre del proveedor"
        onChange={v=>setG(p=>({...p,proveedor:v}))}/>
      <Sel label="Categoría de gasto" value={g.cta} onChange={v=>setG(p=>({...p,cta:v}))}
        options={GASTOS_CAT.map(c=>({value:c.cta,label:`${c.cta} — ${c.nom}`}))}/>
      <div className="flex gap-2">
        {["FACT","FPEQ","FCAM"].map(t=>(
          <button key={t} onClick={()=>setG(p=>({...p,tipo:t}))}
            className={`flex-1 py-1.5 rounded-lg text-xs font-bold border ${g.tipo===t?"bg-orange-600 text-white border-orange-600":"bg-white text-gray-600 border-gray-300"}`}>{t}</button>
        ))}
      </div>
      <Inp label="Total factura c/IVA (Q)" type="number" step="0.01" value={g.total}
        onChange={v=>setG(p=>({...p,total:v}))}/>
      {cat.iva==="petro"&&<Inp label="⛽ Impuesto al petróleo (col W FEL)" type="number" step="0.01"
        value={g.impPet} placeholder="Ver col W en FEL Compras" onChange={v=>setG(p=>({...p,impPet:v}))}/>}
      {g.total>0&&<div className={`rounded-lg p-3 text-sm ${cat.iva==="fpeq"?"bg-yellow-50":"bg-green-50"}`}>
        <div className="grid grid-cols-3 gap-2 text-center">
          <div><div className="text-xs text-gray-500">Base gasto</div><b>{q(base)}</b></div>
          <div><div className="text-xs text-gray-500">IVA CF</div>
            <b className={cat.iva==="fpeq"?"text-yellow-600":"text-green-600"}>
              {cat.iva==="fpeq"?"Sin IVA (FPEQ)":q(iva)}</b></div>
          <div><div className="text-xs text-gray-500">Total</div><b>{q(r2(base+iva))}</b></div>
        </div>
      </div>}
      <Sel label="Forma de pago" value={g.pagoCta} onChange={v=>setG(p=>({...p,pagoCta:v}))}
        options={BANCOS_PAGO.map(b=>({value:b.cta,label:b.nom}))}/>
      {p&&<PreviewPartida p={p} colorHdr="bg-orange-800"/>}
      <Btn full color="orange" disabled={!g.total||!g.proveedor}
        onClick={()=>{if(g.total&&g.proveedor){
          onSave({...c,id:`c-${Date.now()}`,serie:g.serie});
          setG({fecha:g.fecha,proveedor:"",cta:g.cta,total:"",impPet:"",pagoCta:g.pagoCta,tipo:"FACT",serie:""});}}}>
        ✔ Registrar Gasto — P{num}
      </Btn>
    </div>
  );
}

// ── IMPORTAR FEL ──────────────────────────────────────────────
function ImportarFEL({onVentas,onCompras}){
  const[msg,setMsg]=useState(""),[loading,setLoading]=useState(false);
  const handle=async(e)=>{
    const file=e.target.files[0];if(!file)return;
    setLoading(true);setMsg("");
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const hojas=wb.SheetNames;
      let vn=0,cn=0;

      // Detección flexible: busca por nombre O por posición si solo hay 1-2 hojas
      hojas.forEach(name=>{
        const n=name.toLowerCase().trim();
        const esVenta = n.includes("venta") || n.includes("sale") || n.includes("fel_v")
          || n.includes("emitida") || n==="hoja1" || n==="sheet1";
        const esCompra = n.includes("compra") || n.includes("purchase") || n.includes("fel_c")
          || n.includes("gasto") || n.includes("recibida");

        if(esVenta){
          const r=parseFELVentas(wb.Sheets[name]).filter(x=>x.estado!=="ANULADO");
          if(r.length>0){onVentas(r);vn+=r.length;}
        }
        if(esCompra){
          const r=parseFELCompras(wb.Sheets[name]).filter(x=>x.estado!=="ANULADO");
          if(r.length>0){onCompras(r);cn+=r.length;}
        }
      });

      // Si aún no encontró nada, intentar con TODAS las hojas
      if(vn===0&&cn===0){
        // Diagnóstico extra: mostrar primeras filas de cada hoja
        let diag="";
        hojas.forEach(name=>{
          const data2=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:null}).slice(0,3);
          const fila1=data2[0]?data2[0].filter(Boolean).slice(0,4).join("|"):"vacía";
          diag+=`"${name}": ${fila1} · `;
        });
        setMsg(`⚠️ 0 registros. Estructura detectada: ${diag.slice(0,-3)}`);
      } else {
        setMsg(`✅ Importado: ${vn} ventas · ${cn} compras`);
      }
    }catch(err){setMsg("❌ Error: "+err.message);}
    setLoading(false);e.target.value="";
  };
  return(
    <div className="space-y-4">
      <div className="bg-blue-50 border border-blue-200 rounded-xl p-3 text-xs text-blue-700 space-y-1">
        <p className="font-bold text-sm mb-1">📥 Importar Excel SAT</p>
        {[["Combustible","descuenta imp. petróleo col W"],["FPEQ","sin IVA — detectado automáticamente"],
          ["VISA","extrae neto, comisión, retención, membresía"],
          ["Vivian/Arrendamiento","se clasifica en 6.2.05 automáticamente"]].map(([t,d])=>(
          <div key={t} className="flex gap-1.5"><span className="text-green-600">✓</span><span><b>{t}:</b> {d}</span></div>
        ))}
      </div>
      <label className="block cursor-pointer">
        <div className="border-2 border-dashed border-gray-300 rounded-xl p-8 text-center hover:border-blue-400 hover:bg-blue-50 transition-all">
          <p className="text-4xl mb-2">📂</p>
          <p className="font-bold text-gray-700">Seleccionar archivo FEL</p>
          <p className="text-xs text-gray-400 mt-1">.xlsx exportado del portal SAT</p>
        </div>
        <input type="file" accept=".xlsx,.xls" onChange={handle} className="hidden"/>
      </label>
      {loading&&<p className="text-center text-blue-600 py-2">⏳ Procesando...</p>}
      {msg&&<div className={`rounded-lg px-4 py-3 text-sm font-semibold text-center ${
        msg.startsWith("✅")?"bg-green-100 text-green-800":"bg-red-100 text-red-800"}`}>{msg}</div>}
    </div>
  );
}

// ── PARTIDAS SEMANALES ────────────────────────────────────────
// ── CONCILIACIÓN BANCARIA ─────────────────────────────────────
function ConciliacionBanco({ventas}){
  const[semFiltro,setSemFiltro]=useState("todas");
  const[cuentaFiltro,setCuentaFiltro]=useState("todas");

  const CUENTAS_BANCO={
    "1.1.17.004":{nom:"Banco Industrial Q",color:"blue"},
    "1.1.17.005":{nom:"Banrural Q",color:"teal"},
    "1.1.03":{nom:"CxC Pendiente",color:"orange"},
  };

  // Extraer todos los cobros bancarios con detalle
  const movimientos=ventas.flatMap(v=>{
    const dia=parseInt((String(v.fecha||"").match(/-(\d{2})(?:T|$)/)||[])[1]||0);
    const sem=SEM_RANGES.find(s=>dia>=s.d1&&dia<=s.d2)?.label||"Sin semana";
    return (v.cobros||[]).filter(c=>c.tipo==="banco"||c.tipo==="visa").map(c=>({
      id:c.id, fecha:v.fecha, serie:v.serie, cliente:v.cliente,
      sem, cta:c.tipo==="visa"?"1.1.17.004":c.cta,
      tipo:c.tipo, doc:c.doc||"",
      monto:c.tipo==="visa"
        ?r2(Number(c.visaNet||0)+Number(c.visaRet||0)+Number(c.visaCom||0)+Number(c.visaMem||0))
        :r2(Number(c.monto||0)),
      visaNet:c.visaNet, visaRet:c.visaRet, visaCom:c.visaCom, visaMem:c.visaMem,
    }));
  });

  const filtrados=movimientos.filter(m=>
    (semFiltro==="todas"||m.sem===semFiltro)&&
    (cuentaFiltro==="todas"||m.cta===cuentaFiltro)
  );

  // Totales por cuenta
  const totPorCta={};
  movimientos.forEach(m=>{totPorCta[m.cta]=(totPorCta[m.cta]||0)+m.monto;});

  // Totales por semana
  const totPorSem={};
  movimientos.forEach(m=>{
    if(!totPorSem[m.sem]) totPorSem[m.sem]={};
    totPorSem[m.sem][m.cta]=(totPorSem[m.sem][m.cta]||0)+m.monto;
  });

  const totalGral=r2(Object.values(totPorCta).reduce((a,v)=>a+v,0));

  return(
    <div className="space-y-4">
      <Btn full color="green" onClick={()=>xlsxBanco(ventas)}>
        ⬇ Exportar Conciliación Bancaria Excel
      </Btn>
      {/* Resumen por cuenta */}
      <div className="grid grid-cols-1 gap-2">
        {Object.entries(CUENTAS_BANCO).map(([cta,info])=>{
          const tot=r2(totPorCta[cta]||0);
          if(!tot) return null;
          const colores={blue:"bg-blue-50 border-blue-300 text-blue-800",
            teal:"bg-teal-50 border-teal-300 text-teal-800",
            orange:"bg-orange-50 border-orange-300 text-orange-800"}[info.color];
          return(
            <div key={cta} className={`border rounded-xl p-3 flex justify-between items-center ${colores}`}>
              <div>
                <p className="font-bold text-sm">{info.nom}</p>
                <p className="text-xs opacity-70">{movimientos.filter(m=>m.cta===cta).length} depósitos</p>
              </div>
              <p className="font-bold text-lg">{q(tot)}</p>
            </div>
          );
        })}
        <div className="bg-gray-800 text-white rounded-xl p-3 flex justify-between items-center">
          <span className="font-bold">TOTAL COBROS</span>
          <span className="font-bold text-lg">{q(totalGral)}</span>
        </div>
      </div>

      {/* Resumen semanal */}
      <div className="bg-white border rounded-xl overflow-hidden">
        <div className="bg-gray-700 text-white px-3 py-2 text-xs font-bold">
          Depósitos por semana
        </div>
        <div className="divide-y text-xs">
          {SEM_RANGES.map(s=>{
            const tBI=r2(totPorSem[s.label]?.["1.1.17.004"]||0);
            const tBR=r2(totPorSem[s.label]?.["1.1.17.005"]||0);
            const tTotal=r2(tBI+tBR);
            if(!tTotal) return null;
            return(
              <div key={s.label} className="px-3 py-2">
                <div className="flex justify-between font-semibold mb-1">
                  <span>{s.label}</span><span>{q(tTotal)}</span>
                </div>
                <div className="flex gap-3 text-gray-500">
                  {tBI>0&&<span>🏦 BI {q(tBI)}</span>}
                  {tBR>0&&<span>🏦 BR {q(tBR)}</span>}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Filtros */}
      <div className="grid grid-cols-2 gap-2">
        <Sel label="Semana" value={semFiltro} onChange={setSemFiltro}
          options={[{value:"todas",label:"Todas las semanas"},
            ...SEM_RANGES.map(s=>({value:s.label,label:s.label}))]}/>
        <Sel label="Banco" value={cuentaFiltro} onChange={setCuentaFiltro}
          options={[{value:"todas",label:"Todos los bancos"},
            {value:"1.1.17.004",label:"Banco Industrial"},
            {value:"1.1.17.005",label:"Banrural"}]}/>
      </div>

      {/* Detalle boletas */}
      <div className="border rounded-xl overflow-hidden">
        <div className="bg-blue-900 text-white px-3 py-2 flex justify-between text-xs font-bold">
          <span>Detalle de depósitos ({filtrados.length})</span>
          <span>{q(r2(filtrados.reduce((a,m)=>a+m.monto,0)))}</span>
        </div>
        <div className="divide-y text-xs max-h-96 overflow-y-auto">
          {filtrados.length===0&&<p className="text-center text-gray-400 py-4">Sin movimientos</p>}
          {filtrados.map((m,i)=>(
            <div key={i} className={`px-3 py-2 ${i%2?"bg-white":"bg-gray-50"}`}>
              <div className="flex justify-between items-start">
                <div className="flex-1 min-w-0">
                  <div className="flex gap-2 items-center">
                    <span className="font-mono text-blue-700 shrink-0">{m.serie||"—"}</span>
                    <span className="text-gray-500 truncate">{m.cliente}</span>
                  </div>
                  <div className="flex gap-2 mt-0.5 text-gray-400">
                    <span>{m.fecha?.slice(5)}</span>
                    <span>{m.sem}</span>
                    {m.doc&&<span className="text-blue-600">Doc:{m.doc}</span>}
                    {m.tipo==="visa"&&(
                      <span className="text-purple-600">
                        VISA neto {q(r2(Number(m.visaNet||0)))}
                        {m.visaRet>0&&` ret ${q(r2(Number(m.visaRet||0)))}`}
                      </span>
                    )}
                  </div>
                </div>
                <div className="text-right shrink-0 ml-2">
                  <p className="font-bold">{q(m.monto)}</p>
                  <p className="text-gray-400">
                    {m.cta==="1.1.17.004"?"BI":m.cta==="1.1.17.005"?"BR":"CxC"}
                  </p>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function PartidasSemanales({ventas,compras,nextNum,onRegistrar,onRegistrarGastos,onEliminar}){
  const[abierto,setAbierto]=useState(null);
  const[abiertoG,setAbiertoG]=useState(null);
  // Estado para edición de cuentas contables en gastos
  const[editCta,setEditCta]=useState({}); // {gastoId: nuevaCta}
  const getCta=(g)=>editCta[g.id]||g.cta;
  const setCta=(id,cta)=>setEditCta(prev=>({...prev,[id]:cta}));
  // Gastos con cuentas editadas aplicadas
  const aplicarEdiciones=(gs)=>gs.map(g=>({...g,cta:getCta(g)}));
  // Cobros semanales: el usuario ingresa los totales reales del banco
  const[cobrosSemanales,setCobrosSemanales]=useState({
    0:{bi:"",br:"",visaNet:"",visaRet:"",visaCom:"",visaMem:"",dep:""},
    1:{bi:"",br:"",visaNet:"",visaRet:"",visaCom:"",visaMem:"",dep:""},
    2:{bi:"",br:"",visaNet:"",visaRet:"",visaCom:"",visaMem:"",dep:""},
    3:{bi:"",br:"",visaNet:"",visaRet:"",visaCom:"",visaMem:"",dep:""},
  });
  const upd=(i,campo,val)=>setCobrosSemanales(prev=>({...prev,[i]:{...prev[i],[campo]:val}}));
  const dia=f=>{const m=String(f||"").match(/-(\d{2})$/);return m?parseInt(m[1]):0;};

  const buildCobros=(cs)=>{
    const L=[];
    const bi=r2(Number(cs.bi||0)),br=r2(Number(cs.br||0));
    const vn=r2(Number(cs.visaNet||0)),vr=r2(Number(cs.visaRet||0));
    const vc=r2(Number(cs.visaCom||0)),vm=r2(Number(cs.visaMem||0));
    const dep=r2(Number(cs.dep||0));
    if(bi>0)  L.push({id:"bi",tipo:"banco",cta:"1.1.17.004",monto:bi});
    if(br>0)  L.push({id:"br",tipo:"banco",cta:"1.1.17.005",monto:br});
    if(vn>0)  L.push({id:"vn",tipo:"visa",cta:"VISA",visaNet:vn,visaRet:vr,visaCom:vc,visaMem:vm});
    if(dep>0) L.push({id:"dep",tipo:"deposito",cta:"2.1.05",monto:dep});
    if(dep<0) L.push({id:"dep",tipo:"deposito",cta:"2.1.05",monto:dep});
    return L;
  };

  return(
    <div className="space-y-3">
      <div className="bg-green-50 border border-green-200 rounded-xl p-3 text-xs text-green-700">
        <b>💡 Ingrese los cobros REALES del banco por semana</b> — use los totales del estado de cuenta, no calcule por factura. Así evita diferencias de redondeo.
      </div>
      {SEM_RANGES.map((s,i)=>{
        const vtas=ventas.filter(v=>{const d=dia(v.fecha);return d>=s.d1&&d<=s.d2;});
        const tB=r2(vtas.reduce((a,v)=>a+v.base,0));
        const tT=r2(vtas.reduce((a,v)=>a+v.base+v.iva,0));
        const cs=cobrosSemanales[i];
        // Auto-calcular cobros desde facturas si el usuario no ingresó valores
        const autoBI=r2(vtas.flatMap(v=>v.cobros||[]).filter(c=>c.tipo==="banco"&&c.cta==="1.1.17.004").reduce((s,c)=>s+Number(c.monto||0),0));
        const autoBR=r2(vtas.flatMap(v=>v.cobros||[]).filter(c=>c.tipo==="banco"&&c.cta==="1.1.17.005").reduce((s,c)=>s+Number(c.monto||0),0));
        const autoCxC=r2(vtas.flatMap(v=>v.cobros||[]).filter(c=>c.tipo==="banco"&&c.cta==="1.1.03").reduce((s,c)=>s+Number(c.monto||0),0));
        const autoVN=r2(vtas.flatMap(v=>v.cobros||[]).filter(c=>c.tipo==="visa").reduce((s,c)=>s+Number(c.visaNet||0),0));
        const autoVR=r2(vtas.flatMap(v=>v.cobros||[]).filter(c=>c.tipo==="visa").reduce((s,c)=>s+Number(c.visaRet||0),0));
        const autoVC=r2(vtas.flatMap(v=>v.cobros||[]).filter(c=>c.tipo==="visa").reduce((s,c)=>s+Number(c.visaCom||0),0));
        const autoVM=r2(vtas.flatMap(v=>v.cobros||[]).filter(c=>c.tipo==="visa").reduce((s,c)=>s+Number(c.visaMem||0),0));
        // Usar valor manual si existe, sino usar auto-calculado
        const biVal=cs.bi!==""?Number(cs.bi||0):autoBI;
        const brVal=cs.br!==""?Number(cs.br||0):autoBR;
        const vnVal=cs.visaNet!==""?Number(cs.visaNet||0):autoVN;
        const vrVal=cs.visaRet!==""?Number(cs.visaRet||0):autoVR;
        const vcVal=cs.visaCom!==""?Number(cs.visaCom||0):autoVC;
        const vmVal=cs.visaMem!==""?Number(cs.visaMem||0):autoVM;
        const depVal=Number(cs.dep||0);
        const totalCobros=r2(biVal+brVal+vnVal+vrVal+vcVal+vmVal+depVal+autoCxC);
        const diff=r2(tT-totalCobros);
        const cuadra=Math.abs(diff)<0.10;
        // Generar cobros con valores reales
        const allCobros2=[];
        if(biVal>0) allCobros2.push({id:"bi",tipo:"banco",cta:"1.1.17.004",monto:biVal});
        if(brVal>0) allCobros2.push({id:"br",tipo:"banco",cta:"1.1.17.005",monto:brVal});
        if(autoCxC>0) allCobros2.push({id:"cxc",tipo:"banco",cta:"1.1.03",monto:autoCxC});
        if(vnVal>0) allCobros2.push({id:"vn",tipo:"visa",cta:"VISA",visaNet:vnVal,visaRet:vrVal,visaCom:vcVal,visaMem:vmVal});
        if(depVal>0) allCobros2.push({id:"dep",tipo:"deposito",cta:"2.1.05",monto:depVal});
        const allC=allCobros2;
        const cobros=allCobros2;
        const L=cobrosToLineas(allC,`${s.label} (${vtas.length} facturas)`,"");
        L.push({cta:"4.1.01",debe:0,haber:tB,conc:`Ventas ${s.label} - ${vtas.length} facturas`});
        L.push({cta:"2.1.03",debe:0,haber:r2(vtas.reduce((a,v)=>a+v.iva,0)),conc:`IVA DT 12% - ${s.label}`});
        const p=vtas.length>0?{num:nextNum+i,fecha:vtas[0].fecha,tipo:"VENTA_SEM",
          concepto:`Ventas ${s.label}`,lineas:L}:null;
        return(
          <div key={s.label} className="border rounded-xl overflow-hidden shadow-sm">
            <div className="bg-blue-900 text-white px-4 py-2 flex justify-between items-center">
              <div><span className="font-bold">{s.label}</span><span className="text-blue-300 text-xs ml-2">días {s.d1}–{s.d2}</span></div>
              <span className="text-xs bg-blue-700 px-2 py-0.5 rounded-full">{vtas.length} facturas</span>
            </div>
            {vtas.length===0?<p className="text-xs text-gray-400 text-center py-3">Sin ventas registradas</p>:(
              <div className="p-3 space-y-3">
                {/* Totales de facturas */}
                <div className="grid grid-cols-2 gap-2 text-sm bg-gray-50 rounded-lg p-2 text-center">
                  <div><div className="text-xs text-gray-500">Total facturas c/IVA</div><b className="text-blue-700">{q(tT)}</b></div>
                  <div><div className="text-xs text-gray-500">Base s/IVA</div><b className="text-green-700">{q(tB)}</b></div>
                </div>
                {/* Cobros reales del banco — ingreso manual */}
                <div className="bg-white border border-gray-200 rounded-xl p-3 space-y-2">
                  <p className="text-xs font-bold text-gray-600 uppercase tracking-wide">Cobros reales del banco</p>
                  <div className="grid grid-cols-2 gap-2">
                    <Inp label="BI Q (depósito total)" type="number" step="0.01"
                      value={cs.bi} onChange={v=>upd(i,"bi",v)} placeholder="Q0.00"/>
                    <Inp label="Banrural Q (depósito total)" type="number" step="0.01"
                      value={cs.br} onChange={v=>upd(i,"br",v)} placeholder="Q0.00"/>
                  </div>
                  {/* VISA si aplica */}
                  <div className="grid grid-cols-2 gap-2">
                    <Inp label="VISA neto depositado" type="number" step="0.01"
                      value={cs.visaNet} onChange={v=>upd(i,"visaNet",v)} placeholder="Q0.00"/>
                    <Inp label="IVA retenido VISA" type="number" step="0.01"
                      value={cs.visaRet} onChange={v=>upd(i,"visaRet",v)} placeholder="Q0.00"/>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <Inp label="Comisión VISA bruta" type="number" step="0.01"
                      value={cs.visaCom} onChange={v=>upd(i,"visaCom",v)} placeholder="Q0.00"/>
                    <Inp label="Membresía VISA" type="number" step="0.01"
                      value={cs.visaMem} onChange={v=>upd(i,"visaMem",v)} placeholder="Q0.00"/>
                  </div>
                  <Inp label="Dep. x identificar (+aplicar / −exceso)" type="number" step="0.01"
                    value={cs.dep} onChange={v=>upd(i,"dep",v)} placeholder="Q0.00"/>
                  {/* Verificación cuadre */}
                  <div className={`rounded-lg px-3 py-2 text-xs flex justify-between font-semibold ${
                    cuadra?"bg-green-50 border border-green-300 text-green-700":"bg-red-50 border border-red-300 text-red-600"}`}>
                    <div>
                      <span>Facturas {q(tT)} · Cobros {q(totalCobros)}</span>
                      {autoCxC>0&&<span className="text-orange-600 ml-2">CxC {q(autoCxC)}</span>}
                    </div>
                    <span>{cuadra?"✅ Cuadra":`⚠️ Dif Q${Math.abs(diff).toFixed(2)}`}</span>
                  </div>
                </div>
                <div className="flex gap-2">
                  <button onClick={()=>setAbierto(abierto===i?null:i)}
                    className="flex-1 text-xs border border-blue-300 text-blue-700 py-1.5 rounded-lg hover:bg-blue-50">
                    {abierto===i?"Ocultar partida":"Ver partida"}
                  </button>
                  <button onClick={()=>p&&onRegistrar(p)}
                    className="flex-1 text-xs bg-green-700 text-white py-1.5 rounded-lg hover:bg-green-800 font-semibold">
                    Registrar P{nextNum+i}
                  </button>
                </div>
                {abierto===i&&p&&<PreviewPartida p={p}/>}
                {/* Lista facturas con botón eliminar */}
                <div className="text-xs text-gray-500 max-h-32 overflow-y-auto space-y-0.5 border-t pt-2">
                  {vtas.map((v,j)=>(
                    <div key={j} className="flex justify-between items-center gap-1 hover:bg-red-50 rounded px-1">
                      <span className="truncate flex-1">{v.serie||"—"} {v.cliente}</span>
                      <span className="font-semibold shrink-0">{q(r2(v.base+v.iva))}</span>
                      <button onClick={()=>onEliminar&&onEliminar(v.id)}
                        className="text-red-400 hover:text-red-600 font-bold ml-1 shrink-0 w-4 text-center"
                        title="Eliminar factura">✕</button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        );
      })}

      {/* ── SECCIÓN GASTOS ─────────────────────────── */}
      {compras&&compras.length>0&&(()=>{
        const pendientes=compras.filter(c=>!c.registrado&&c.estado!=="ANULADO");
        if(!pendientes.length) return null;
        return(
          <div className="border-2 border-orange-300 rounded-xl overflow-hidden mt-2">
            <div className="bg-orange-700 text-white px-4 py-2.5 flex justify-between items-center">
              <div>
                <span className="font-bold">🧾 Gastos / Compras pendientes</span>
                <span className="text-orange-200 text-xs ml-2">{pendientes.length} facturas importadas</span>
              </div>
              <button onClick={()=>setAbiertoG(!abiertoG)}
                className="text-xs bg-orange-600 px-3 py-1 rounded-full hover:bg-orange-500">
                {abiertoG?"Ocultar":"Ver detalle"}
              </button>
            </div>
            <div className="p-3 space-y-2">
              {/* Resumen por semana */}
              {SEM_RANGES.map((s,si)=>{
                const dia2=f=>{const m=String(f||"").match(/-(\d{2})(?:T|$)/);return m?parseInt(m[1]):0;};
                const gs=pendientes.filter(c=>{const d=dia2(c.fecha);return d>=s.d1&&d<=s.d2;});
                if(!gs.length) return null;
                const tG=r2(gs.reduce((a,c)=>a+c.base,0));
                const tCF=r2(gs.reduce((a,c)=>a+c.iva,0));
                return(
                  <div key={si} className="bg-orange-50 rounded-lg p-3 border border-orange-200 space-y-2">
                    {/* Encabezado semana */}
                    <div className="flex justify-between items-center">
                      <span className="text-xs font-bold text-orange-800">{s.label} — {gs.length} facturas</span>
                      <span className="text-xs text-orange-600">días {s.d1}–{s.d2}</span>
                    </div>
                    {/* Totales */}
                    <div className="grid grid-cols-3 gap-1 text-xs text-center bg-white rounded p-2">
                      <div><div className="text-gray-400">Base gasto</div><b className="text-orange-700">{q(tG)}</b></div>
                      <div><div className="text-gray-400">IVA CF</div><b className="text-green-700">{q(tCF)}</b></div>
                      <div><div className="text-gray-400">Total pagado</div><b>{q(r2(tG+tCF))}</b></div>
                    </div>
                    {/* Detalle facturas con método de pago */}
                    <div className="space-y-1 max-h-32 overflow-y-auto">
                      {gs.map((c,ci)=>{
                        const pagNom=CUENTAS[c.pagoCta]||c.pagoCta||"Caja";
                        const pagColor=c.pagoCta==="1.1.17.004"?"text-blue-600":
                          c.pagoCta==="1.1.17.005"?"text-teal-600":"text-gray-500";
                        const ctaActual=getCta(c);
                        const ctaEditada=editCta[c.id];
                        return(
                          <div key={ci} className={`text-xs bg-white rounded px-2 py-1.5 border ${ctaEditada?"border-blue-300 bg-blue-50":"border-orange-100"}`}>
                            <div className="flex justify-between items-start">
                              <div className="flex-1 min-w-0 mr-2">
                                <span className="text-gray-700 truncate block font-semibold">{c.proveedor?.slice(0,28)}</span>
                                <div className="flex gap-2 mt-0.5 flex-wrap">
                                  <span className={pagColor}>
                                    {c.pagoCta==="1.1.17.004"?"🏦 BI":
                                     c.pagoCta==="1.1.17.005"?"🏦 BR":
                                     c.pagoCta==="2.1.01"?"📋 CxP":"💵 Caja"}
                                  </span>
                                  {c.tipo==="FESP"&&<span className="text-red-600 font-semibold">FESP+ISR</span>}
                                  {c.tipo==="FPEQ"&&<span className="text-yellow-600 font-semibold">FPEQ</span>}
                                </div>
                              </div>
                              <div className="text-right shrink-0">
                                <div className="font-semibold">Q{Number(c.total||0).toFixed(2)}</div>
                                {c.iva>0&&<div className="text-green-600">CF Q{Number(c.iva||0).toFixed(2)}</div>}
                              </div>
                            </div>
                            {/* Selector de cuenta contable */}
                            <div className="mt-1.5 flex items-center gap-1">
                              <span className="text-gray-400 shrink-0">Cuenta:</span>
                              <select value={ctaActual}
                                onChange={e=>setCta(c.id,e.target.value)}
                                className={`flex-1 text-xs border rounded px-1 py-0.5 ${ctaEditada?"border-blue-400 text-blue-700 bg-blue-50 font-semibold":"border-gray-200 text-gray-600"}`}>
                                {GASTOS_CAT.map(g=>(
                                  <option key={g.cta} value={g.cta}>{g.cta} — {g.nom}</option>
                                ))}
                              </select>
                              {ctaEditada&&(
                                <button onClick={()=>setCta(c.id,c.cta)}
                                  className="text-gray-400 hover:text-red-500 shrink-0 px-1"
                                  title="Restaurar original">↩</button>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    {/* Vista previa partida con cuentas editadas */}
                    {(()=>{
                      const pG=genPartidaGastosSem(aplicarEdiciones(gs),nextNum+100+si,s.label);
                      return pG&&<PreviewPartida p={pG} colorHdr="bg-orange-800"/>;
                    })()}
                    {/* Botón registrar */}
                    <button onClick={()=>onRegistrarGastos(aplicarEdiciones(gs), s.label)}
                      className="w-full text-xs bg-orange-700 text-white py-2 rounded-lg hover:bg-orange-800 font-bold">
                      ✔ Registrar P{nextNum+si} — Gastos {s.label} consolidado
                    </button>
                  </div>
                );
              })}
              <p className="text-xs text-orange-500 text-center">
                1 partida consolidada por semana — igual que las ventas
              </p>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ── LIBROS ────────────────────────────────────────────────────
function LibroDiarioView({partidas}){
  const[f,setF]=useState("");
  const pf=partidas.filter(p=>!f||p.concepto.toLowerCase().includes(f.toLowerCase())||String(p.num).includes(f));
  return(
    <div className="space-y-3">
      <div className="flex gap-2">
        <Inp className="flex-1" value={f} onChange={setF} placeholder="Buscar..."/>
        <Btn size="sm" color="green" onClick={()=>xlsxCompatible(partidas)}>⬇ Para pegar</Btn>
      </div>
      {pf.length===0&&<p className="text-center text-gray-400 py-8 text-sm">Sin partidas registradas</p>}
      {pf.map(p=>{
        const sd=r2(p.lineas.reduce((a,l)=>a+l.debe,0)),sh=r2(p.lineas.reduce((a,l)=>a+l.haber,0));
        return(
          <div key={p.num} className="border rounded-lg overflow-hidden shadow-sm">
            <div className={`px-3 py-1.5 text-xs font-bold flex justify-between text-white ${
              p.tipo.startsWith("VENTA")?"bg-green-800":p.tipo==="COMPRA"?"bg-orange-700":"bg-blue-900"}`}>
              <span className="truncate">P{p.num} · {p.fecha} · {p.concepto.slice(0,40)}</span>
              <span>{Math.abs(sd-sh)<0.02?"✅":"⚠️"}</span>
            </div>
            {p.lineas.map((l,i)=>(
              <div key={i} className={`grid grid-cols-4 px-3 py-0.5 text-xs ${i%2?"bg-white":"bg-gray-50"}`}>
                <span className="font-mono text-blue-700">{l.cta}</span>
                <span className="text-gray-600 truncate">{CUENTAS[l.cta]||l.cta}</span>
                <span className="text-right text-green-700">{l.debe?q(l.debe):""}</span>
                <span className="text-right text-red-700">{l.haber?q(l.haber):""}</span>
              </div>
            ))}
            <div className="grid grid-cols-4 px-3 py-1 bg-gray-100 font-bold text-xs border-t">
              <span className="col-span-2">TOTALES P{p.num}</span>
              <span className="text-right">{q(sd)}</span><span className="text-right">{q(sh)}</span>
            </div>
          </div>
        );
      })}
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
  const[modoVenta,setModoVenta]=useState("semanal"); // "semanal" | "individual"
  const[confirmarBorrar,setConfirmarBorrar]=useState(false);
  const[ready,setReady]=useState(false);
  const[storageOk,setStorageOk]=useState(false);

  useEffect(()=>{
    (async()=>{
      try{
        // CARGA SIMPLIFICADA — datos pre-procesados siempre disponibles
        // 1. Intentar cargar partidas desde storage (lo único que el usuario genera)
        const[lp,lsn,lmv]=await Promise.all([
          cargar("nx4-p"),cargar("nx4-sn"),cargar("nx4-mv")]);

        // 2. Ventas y compras SIEMPRE desde datos pre-procesados (FEL oficial)
        const v=DATOS_INICIALES.ventas;
        const c=DATOS_INICIALES.compras.map(x=>({...x,registrado:false}));

        // 3. Partidas: MERGE de pre-cargadas + adicionales del storage
        // P270-P273 siempre están en código — no se pueden perder
        // P274+ vienen del storage (gastos registrados por el usuario)
        const baseNums=new Set(PARTIDAS_INICIALES.map(x=>x.num));
        let adicionales=[];
        if(lp&&lp.length>0){
          adicionales=lp.filter(x=>!baseNums.has(x.num));
        }
        const p=[...PARTIDAS_INICIALES,...adicionales];
        console.log(`Cargado: ${v.length}v ${c.length}c ${p.length}p (${adicionales.length} adicionales del storage)`);

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
    const num=startNum+partidas.length;
    const p=genPartidaGasto(c,num);
    setCompras(prev=>[...prev,c]);
    setPartidas(prev=>[...prev,p]);
    setTab("dashboard");
  },[partidas,startNum]);

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

  const registrarSem=useCallback(p=>{
    if(!p)return;
    const num=startNum+partidas.length;
    setPartidas(prev=>{
      const nuevas=[...prev,{...p,num}];
      // Auto-exportar al registrar partida semanal
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
              <div><h2 className="font-bold text-gray-800">Registrar Gasto — P{nextNum}</h2>
                <p className="text-xs text-gray-500">IVA calculado por categoría</p></div>
            </div>
            <FormGasto num={nextNum} onSave={addCompra}/>
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
              onEliminar={eliminarVenta}/>
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
