const SPREADSHEET_ID = "1JZgphzXLUAj0hjH69pQ_WzViOT0xRnDJGKPZvcCkQ0U";
const SHEET_NAME = "BIPA RADAR";

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('No existe la pestaña "' + SHEET_NAME + '".');
    }

    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      return jsonResponse(buildEmptyResponse(), e);
    }

    const headers = values[0].map(normalizeHeader);
    const idx = getColumnIndexes(headers);
    validateRequiredColumns(idx);

    const grouped = {};

    values.slice(1).forEach(function(row) {
      const record = readRecord(row, idx);

      if (!record.cliente || !record.articulo || !record.referencia) return;
      if (record.precioUsd <= 0) return;

      const key = [
        record.cliente.toLowerCase(),
        record.referencia.toLowerCase(),
        record.articulo.toLowerCase(),
        record.marca.toLowerCase()
      ].join("|");

      if (!grouped[key] || record.precioUsd > grouped[key].precioUsd) {
        grouped[key] = record;
      }
    });

    const data = enrichRecords(Object.keys(grouped).map(function(key) {
      return grouped[key];
    }));

    return jsonResponse({
      success: true,
      count: data.length,
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
      dataUpdatedAt: getLatestDataTimestamp(data),
      data: data,
      resumen: buildSummary(data),
      filtros: buildFilters(data),
      referencias: buildReferenceSummary(data),
      clientes: buildClientSummary(data),
      marcas: buildBrandSummary(data)
    }, e);

  } catch (error) {
    return jsonResponse({
      success: false,
      message: error.message,
      data: [],
      resumen: {},
      filtros: {
        referencias: [],
        clientes: [],
        marcas: [],
        estados: []
      }
    }, e);
  }
}

function getColumnIndexes(headers) {
  return {
    creadoEn: headers.indexOf("creado_en"),
    id: headers.indexOf("id"),
    radarId: headers.indexOf("radar_id"),
    productoId: headers.indexOf("producto_id"),
    clienteId: headers.indexOf("cliente_id"),
    cliente: headers.indexOf("cliente"),
    estado: headers.indexOf("estado"),
    articulo: headers.indexOf("articulo"),
    precioUsd: headers.indexOf("precio_usd"),
    peso: headers.indexOf("peso"),
    marca: headers.indexOf("marca"),
    referencia: headers.indexOf("referencia_para_comparar"),
    precioMayoristaUsd: headers.indexOf("precio_mayorista_usd"),
    precioGranMayoristaUsd: headers.indexOf("precio_gran_mayorista_usd")
  };
}

function validateRequiredColumns(idx) {
  const labels = {
    cliente: "CLIENTE",
    articulo: "articulo",
    precioUsd: "precio_usd",
    marca: "MARCA",
    referencia: "referencia para comparar"
  };

  const missing = Object.keys(labels).filter(function(key) {
    return idx[key] === -1;
  }).map(function(key) {
    return labels[key];
  });

  if (missing.length) {
    throw new Error("Faltan columnas requeridas: " + missing.join(", "));
  }
}

function readRecord(row, idx) {
  const precioUsd = safeNumber(row[idx.precioUsd]);
  const peso = safeNumber(row[idx.peso]);
  const precioMayoristaUsd = idx.precioMayoristaUsd >= 0
    ? safeNumber(row[idx.precioMayoristaUsd])
    : 0;
  const precioGranMayoristaUsd = idx.precioGranMayoristaUsd >= 0
    ? safeNumber(row[idx.precioGranMayoristaUsd])
    : 0;

  return {
    creadoEn: idx.creadoEn >= 0 ? safeDateText(row[idx.creadoEn]) : "",
    id: idx.id >= 0 ? safeText(row[idx.id]) : "",
    radarId: idx.radarId >= 0 ? safeText(row[idx.radarId]) : "",
    productoId: idx.productoId >= 0 ? safeText(row[idx.productoId]) : "",
    clienteId: idx.clienteId >= 0 ? safeText(row[idx.clienteId]) : "",
    cliente: safeText(row[idx.cliente]),
    estado: idx.estado >= 0 ? safeText(row[idx.estado]) : "",
    articulo: safeText(row[idx.articulo]),
    precioUsd: roundTo(precioUsd, 2),
    peso: peso,
    marca: safeText(row[idx.marca]),
    referencia: safeText(row[idx.referencia]),
    precioMayoristaUsd: precioMayoristaUsd > 0 ? roundTo(precioMayoristaUsd, 2) : null,
    precioGranMayoristaUsd: precioGranMayoristaUsd > 0 ? roundTo(precioGranMayoristaUsd, 2) : null,
    precioPorGramo: peso > 0 ? roundTo(precioUsd / peso, 4) : null
  };
}

function enrichRecords(records) {
  const byReference = groupBy(records, "referencia");
  const enriched = [];

  Object.keys(byReference).forEach(function(reference) {
    const items = byReference[reference].slice().sort(function(a, b) {
      return a.precioUsd - b.precioUsd;
    });

    const prices = items.map(function(item) { return item.precioUsd; });
    const minPrice = Math.min.apply(null, prices);
    const maxPrice = Math.max.apply(null, prices);
    const avgPrice = average(prices);
    const bipaItems = items.filter(function(item) {
      return normalizeBrand(item.marca) === "BIPA";
    });
    const bipaMinPrice = bipaItems.length
      ? Math.min.apply(null, bipaItems.map(function(item) { return item.precioUsd; }))
      : null;
    const brandCount = unique(items.map(function(item) { return item.marca; })).length;
    const clientCount = unique(items.map(function(item) { return item.cliente; })).length;

    items.forEach(function(item, index) {
      const differenceVsMinPct = minPrice > 0
        ? roundTo(((item.precioUsd - minPrice) / minPrice) * 100, 2)
        : 0;

      const brechaVsBipaPct = bipaMinPrice && bipaMinPrice > 0
        ? roundTo(((item.precioUsd - bipaMinPrice) / bipaMinPrice) * 100, 2)
        : null;

      enriched.push(Object.assign({}, item, {
        precioMinReferencia: roundTo(minPrice, 2),
        precioMaxReferencia: roundTo(maxPrice, 2),
        precioPromedioReferencia: roundTo(avgPrice, 2),
        diferenciaVsMinimoPct: differenceVsMinPct,
        brechaVsBipaPct: brechaVsBipaPct,
        rankingPrecio: index + 1,
        cantidadCompetidores: brandCount,
        clientesConReferencia: clientCount,
        estatusCompetitivo: getCompetitiveStatus(item, index + 1, differenceVsMinPct, bipaMinPrice)
      }));
    });
  });

  return enriched.sort(function(a, b) {
    return String(a.cliente).localeCompare(String(b.cliente), "es") ||
      String(a.referencia).localeCompare(String(b.referencia), "es") ||
      a.rankingPrecio - b.rankingPrecio;
  });
}

function getCompetitiveStatus(item, ranking, differenceVsMinPct, bipaMinPrice) {
  const isBipa = normalizeBrand(item.marca) === "BIPA";

  if (!bipaMinPrice) return "Sin BIPA";
  if (ranking === 1 && isBipa) return "BIPA líder";
  if (isBipa && differenceVsMinPct <= 5) return "Competitivo";
  if (isBipa && differenceVsMinPct <= 15) return "Revisar";
  if (isBipa && differenceVsMinPct > 15) return "Riesgo alto";
  if (!isBipa && ranking === 1) return "Competencia líder";
  return "Monitorear";
}

function buildSummary(data) {
  const prices = data.map(function(item) { return item.precioUsd; }).filter(isPositiveNumber);
  const references = unique(data.map(function(item) { return item.referencia; }));
  const clients = unique(data.map(function(item) { return item.cliente; }));
  const brands = unique(data.map(function(item) { return item.marca; }));

  return {
    totalProductos: data.length,
    totalReferencias: references.length,
    totalClientes: clients.length,
    totalMarcas: brands.length,
    precioPromedio: roundTo(average(prices), 2),
    precioMinimo: prices.length ? roundTo(Math.min.apply(null, prices), 2) : 0,
    precioMaximo: prices.length ? roundTo(Math.max.apply(null, prices), 2) : 0,
    registrosBipa: data.filter(function(item) { return normalizeBrand(item.marca) === "BIPA"; }).length,
    registrosCompetencia: data.filter(function(item) { return normalizeBrand(item.marca) !== "BIPA"; }).length,
    bipaLider: data.filter(function(item) { return item.estatusCompetitivo === "BIPA líder"; }).length,
    competenciaLider: data.filter(function(item) { return item.estatusCompetitivo === "Competencia líder"; }).length,
    riesgosAltos: data.filter(function(item) { return item.estatusCompetitivo === "Riesgo alto"; }).length,
    sinPeso: data.filter(function(item) { return !item.peso || item.precioPorGramo === null; }).length
  };
}

function buildFilters(data) {
  return {
    referencias: unique(data.map(function(item) { return item.referencia; })).sort(sortEs),
    clientes: unique(data.map(function(item) { return item.cliente; })).sort(sortEs),
    marcas: unique(data.map(function(item) { return item.marca; })).sort(sortEs),
    estados: unique(data.map(function(item) { return item.estado; })).sort(sortEs)
  };
}

function buildReferenceSummary(data) {
  const grouped = groupBy(data, "referencia");

  return Object.keys(grouped).map(function(reference) {
    const items = grouped[reference];
    const prices = items.map(function(item) { return item.precioUsd; }).filter(isPositiveNumber);

    return {
      referencia: reference,
      productos: items.length,
      clientes: unique(items.map(function(item) { return item.cliente; })).length,
      marcas: unique(items.map(function(item) { return item.marca; })).length,
      precioMinimo: prices.length ? roundTo(Math.min.apply(null, prices), 2) : 0,
      precioMaximo: prices.length ? roundTo(Math.max.apply(null, prices), 2) : 0,
      precioPromedio: roundTo(average(prices), 2),
      bipaPresente: items.some(function(item) { return normalizeBrand(item.marca) === "BIPA"; }),
      lider: items.slice().sort(function(a, b) { return a.precioUsd - b.precioUsd; })[0] || null
    };
  }).sort(function(a, b) {
    return b.productos - a.productos || sortEs(a.referencia, b.referencia);
  });
}

function buildClientSummary(data) {
  const grouped = groupBy(data, "cliente");

  return Object.keys(grouped).map(function(client) {
    const items = grouped[client];
    const prices = items.map(function(item) { return item.precioUsd; }).filter(isPositiveNumber);

    return {
      cliente: client,
      productos: items.length,
      referencias: unique(items.map(function(item) { return item.referencia; })).length,
      marcas: unique(items.map(function(item) { return item.marca; })).length,
      precioPromedio: roundTo(average(prices), 2),
      precioMinimo: prices.length ? roundTo(Math.min.apply(null, prices), 2) : 0,
      precioMaximo: prices.length ? roundTo(Math.max.apply(null, prices), 2) : 0,
      riesgosAltos: items.filter(function(item) { return item.estatusCompetitivo === "Riesgo alto"; }).length
    };
  }).sort(function(a, b) {
    return b.productos - a.productos || sortEs(a.cliente, b.cliente);
  });
}

function buildBrandSummary(data) {
  const grouped = groupBy(data, "marca");

  return Object.keys(grouped).map(function(brand) {
    const items = grouped[brand];
    const prices = items.map(function(item) { return item.precioUsd; }).filter(isPositiveNumber);

    return {
      marca: brand,
      productos: items.length,
      referencias: unique(items.map(function(item) { return item.referencia; })).length,
      precioPromedio: roundTo(average(prices), 2),
      precioMinimo: prices.length ? roundTo(Math.min.apply(null, prices), 2) : 0,
      precioMaximo: prices.length ? roundTo(Math.max.apply(null, prices), 2) : 0
    };
  }).sort(function(a, b) {
    return b.productos - a.productos || sortEs(a.marca, b.marca);
  });
}

function buildEmptyResponse() {
  return {
    success: true,
    count: 0,
    generatedAt: "",
    data: [],
    resumen: {
      totalProductos: 0,
      totalReferencias: 0,
      totalClientes: 0,
      totalMarcas: 0,
      precioPromedio: 0,
      precioMinimo: 0,
      precioMaximo: 0,
      registrosBipa: 0,
      registrosCompetencia: 0,
      bipaLider: 0,
      competenciaLider: 0,
      riesgosAltos: 0,
      sinPeso: 0
    },
    filtros: {
      referencias: [],
      clientes: [],
      marcas: [],
      estados: []
    },
    referencias: [],
    clientes: [],
    marcas: []
  };
}

function groupBy(data, key) {
  return data.reduce(function(acc, item) {
    const value = item[key] || "Sin dato";
    if (!acc[value]) acc[value] = [];
    acc[value].push(item);
    return acc;
  }, {});
}

function unique(values) {
  const seen = {};

  return values.filter(function(value) {
    const clean = safeText(value);
    const key = clean.toLowerCase();

    if (!clean || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function average(values) {
  const nums = values.filter(isPositiveNumber);
  if (!nums.length) return 0;
  return nums.reduce(function(sum, value) { return sum + value; }, 0) / nums.length;
}

function getLatestDataTimestamp(data) {
  const timestamps = data
    .map(function(item) { return item.creadoEn; })
    .filter(Boolean)
    .sort();

  return timestamps.length ? timestamps[timestamps.length - 1] : "";
}

function isPositiveNumber(value) {
  return typeof value === "number" && isFinite(value) && value > 0;
}

function normalizeBrand(value) {
  return safeText(value).toUpperCase();
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_");
}

function safeText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function safeDateText(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }

  return safeText(value);
}

function safeNumber(value) {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return isFinite(value) ? value : 0;
  }

  const text = String(value).trim();
  if (!text) return 0;

  const hasComma = text.indexOf(",") >= 0;
  const hasDot = text.indexOf(".") >= 0;
  let normalized = text.replace(/\s/g, "").replace(/\$/g, "");

  if (hasComma && hasDot) {
    normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
  } else {
    normalized = normalized.replace(/,/g, ".");
  }

  const number = Number(normalized);
  return isFinite(number) ? number : 0;
}

function roundTo(value, decimals) {
  const factor = Math.pow(10, decimals);
  return Math.round((Number(value || 0) + Number.EPSILON) * factor) / factor;
}

function sortEs(a, b) {
  return String(a).localeCompare(String(b), "es");
}

function jsonResponse(payload, e) {
  const callback = e && e.parameter && e.parameter.callback
    ? String(e.parameter.callback).replace(/[^\w.$]/g, "")
    : "";

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
